import argparse
import queue
import re
import threading
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Optional, Tuple, Dict, List, Callable

import requests
from pypdf import PdfReader

try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
except Exception:  # GUI optional for CLI usage
    tk = None
    ttk = None
    filedialog = None
    messagebox = None

INVALID_CHARS = r'<>:"/\\|?*'
DOI_RE = re.compile(r"(10\.\d{4,9}/[-._;()/:A-Z0-9]+)", re.IGNORECASE)
YEAR_RE = re.compile(r"\b(19\d{2}|20\d{2})\b")


@dataclass
class PreviewItem:
    pdf: Path
    old_name: str
    new_path: Optional[Path]
    doi: Optional[str]
    title: Optional[str]
    year: Optional[int]
    status: str
    reason: str
    apply: bool = False


def sanitize_filename(name: str) -> str:
    name = re.sub(r"\s+", " ", name).strip()
    name = name.translate(str.maketrans({ch: "_" for ch in INVALID_CHARS}))
    name = name.rstrip(". ").strip()
    name = name.replace("\u0000", "").strip()
    return name or "untitled"


def clamp_filename(stem: str, suffix: str, max_len: int) -> str:
    max_stem_len = max(1, max_len - len(suffix))
    stem = stem.strip()
    if len(stem) > max_stem_len:
        stem = stem[:max_stem_len].rstrip()
    return stem + suffix


def unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    stem, suffix = path.stem, path.suffix
    i = 2
    while True:
        p = path.with_name(f"{stem} ({i}){suffix}")
        if not p.exists():
            return p
        i += 1


def unique_path_with_reserved(path: Path, reserved: set) -> Path:
    if path.name not in reserved:
        reserved.add(path.name)
        return path
    stem, suffix = path.stem, path.suffix
    i = 2
    while True:
        name = f"{stem} ({i}){suffix}"
        if name not in reserved:
            reserved.add(name)
            return path.with_name(name)
        i += 1


def extract_title_from_metadata(pdf_path: Path) -> Optional[str]:
    try:
        reader = PdfReader(str(pdf_path))
        md = reader.metadata
        if md and md.title:
            t = str(md.title).strip()
            bad = ["microsoft word", "untitled", "doi", "title"]
            if len(t) >= 8 and not any(b in t.lower() for b in bad):
                return t
    except Exception:
        pass
    return None


def extract_text_first_pages(pdf_path: Path, max_pages: int) -> str:
    try:
        reader = PdfReader(str(pdf_path))
        texts = []
        for i in range(min(max_pages, len(reader.pages))):
            try:
                texts.append(reader.pages[i].extract_text() or "")
            except Exception:
                pass
        return "\n".join(texts)
    except Exception:
        return ""


def extract_doi_from_text(text: str) -> Optional[str]:
    if not text:
        return None
    m = DOI_RE.search(text)
    if not m:
        return None
    doi = m.group(1).rstrip(".),;")
    return doi


def extract_year_from_text(text: str) -> Optional[int]:
    if not text:
        return None
    years = [int(y) for y in YEAR_RE.findall(text)]
    if not years:
        return None
    from datetime import datetime
    current_year = datetime.now().year
    years = [y for y in years if 1800 < y <= current_year + 1]
    if not years:
        return None
    return max(years)


def crossref_lookup(doi: str, timeout: int, user_agent: str) -> Tuple[Optional[str], Optional[int]]:
    url = f"https://api.crossref.org/works/{doi}"
    headers = {"User-Agent": user_agent, "Accept": "application/json"}
    r = requests.get(url, headers=headers, timeout=timeout)
    if r.status_code != 200:
        return None, None
    data = r.json().get("message", {})
    title_list = data.get("title") or []
    title = title_list[0].strip() if title_list else None
    year = None
    for key in ["issued", "published-print", "published-online", "created"]:
        part = data.get(key) or {}
        date_parts = part.get("date-parts")
        if isinstance(date_parts, list) and date_parts and isinstance(date_parts[0], list) and date_parts[0]:
            y = date_parts[0][0]
            if isinstance(y, int):
                year = y
                break
    return title, year


def build_new_stem(title: str, year: Optional[int], style: str) -> str:
    title = sanitize_filename(title)
    if year:
        if style == "prefix":
            return f"{year} - {title}"
        return f"{title} ({year})"
    return title


def collect_pdfs(folder: Path, recursive: bool) -> List[Path]:
    pattern = "**/*.pdf" if recursive else "*.pdf"
    return sorted(folder.glob(pattern))


def compute_preview(
    folder: Path,
    pdfs: List[Path],
    pages: int,
    maxlen: int,
    style: str,
    no_crossref: bool,
    sleep: float,
    timeout: int,
    unmatched_dir: str,
    user_agent: str,
    progress_cb: Optional[Callable[[int, int, Path], None]] = None,
    item_cb: Optional[Callable[[int, int, "PreviewItem"], None]] = None,
    cancel_event: Optional[threading.Event] = None,
) -> List[PreviewItem]:
    cache: Dict[str, Tuple[Optional[str], Optional[int]]] = {}
    reserved_by_dir: Dict[Path, set] = {}

    def get_reserved(dir_path: Path) -> set:
        if dir_path not in reserved_by_dir:
            try:
                reserved_by_dir[dir_path] = {p.name for p in dir_path.iterdir() if p.is_file()}
            except Exception:
                reserved_by_dir[dir_path] = set()
        return reserved_by_dir[dir_path]

    unmatched_root = folder / unmatched_dir if unmatched_dir else None
    items: List[PreviewItem] = []

    total = len(pdfs)
    for idx, pdf in enumerate(pdfs, 1):
        if cancel_event and cancel_event.is_set():
            break
        if progress_cb:
            progress_cb(idx, total, pdf)
        old_name = pdf.name
        title = extract_title_from_metadata(pdf)
        text = extract_text_first_pages(pdf, pages)
        doi = extract_doi_from_text(text)
        year_guess = extract_year_from_text(text)

        year = None
        if doi and not no_crossref:
            if doi in cache:
                cf_title, cf_year = cache[doi]
            else:
                try:
                    cf_title, cf_year = crossref_lookup(doi, timeout, user_agent)
                except Exception:
                    cf_title, cf_year = None, None
                cache[doi] = (cf_title, cf_year)
                time.sleep(max(0.0, sleep))
            if cf_title:
                title = cf_title
            if cf_year:
                year = cf_year

        if not year and year_guess:
            year = year_guess

        if not title:
            if unmatched_root:
                reserved = get_reserved(unmatched_root)
                dest = unique_path_with_reserved(unmatched_root / pdf.name, reserved)
                item = PreviewItem(
                    pdf=pdf,
                    old_name=old_name,
                    new_path=dest,
                    doi=doi,
                    title=None,
                    year=year,
                    status="move",
                    reason="no title found",
                    apply=True,
                )
                items.append(item)
                if item_cb:
                    item_cb(idx, total, item)
            else:
                item = PreviewItem(
                    pdf=pdf,
                    old_name=old_name,
                    new_path=None,
                    doi=doi,
                    title=None,
                    year=year,
                    status="skip",
                    reason="no title found",
                    apply=False,
                )
                items.append(item)
                if item_cb:
                    item_cb(idx, total, item)
            continue

        new_stem = build_new_stem(title, year, style)
        new_name = clamp_filename(new_stem, ".pdf", maxlen)
        reserved = get_reserved(pdf.parent)
        new_path = unique_path_with_reserved(pdf.with_name(new_name), reserved)

        if new_path.name == old_name:
            item = PreviewItem(
                pdf=pdf,
                old_name=old_name,
                new_path=new_path,
                doi=doi,
                title=title,
                year=year,
                status="ok",
                reason="already good name",
                apply=False,
            )
            items.append(item)
            if item_cb:
                item_cb(idx, total, item)
        else:
            item = PreviewItem(
                pdf=pdf,
                old_name=old_name,
                new_path=new_path,
                doi=doi,
                title=title,
                year=year,
                status="rename",
                reason="ready",
                apply=True,
            )
            items.append(item)
            if item_cb:
                item_cb(idx, total, item)

    return items


def apply_changes(
    items: List[PreviewItem],
    dry_run: bool,
    log=None,
    progress_cb: Optional[Callable[[int, int, PreviewItem], None]] = None,
    cancel_event: Optional[threading.Event] = None,
) -> Tuple[int, int]:
    renamed = 0
    skipped = 0
    total = len(items)
    for idx, item in enumerate(items, 1):
        if cancel_event and cancel_event.is_set():
            break
        if progress_cb:
            progress_cb(idx, total, item)
        if not item.apply or not item.new_path:
            skipped += 1
            continue
        if dry_run:
            if log:
                log(f"[DRY] {item.pdf.name} -> {item.new_path.name}")
            renamed += 1
            continue
        try:
            item.new_path.parent.mkdir(parents=True, exist_ok=True)
            target = item.new_path
            if target.exists():
                target = unique_path(target)
            item.pdf.rename(target)
            if log:
                log(f"[OK] {item.pdf.name} -> {target.name}")
            renamed += 1
        except Exception as e:
            if log:
                log(f"[FAIL] {item.pdf.name}: {e}")
            skipped += 1
    return renamed, skipped


def run_cli(args) -> None:
    folder = Path(args.folder).expanduser().resolve()
    if not folder.exists() or not folder.is_dir():
        raise SystemExit(f"Folder not found: {folder}")

    pdfs = collect_pdfs(folder, args.recursive)
    print(f"Found {len(pdfs)} PDFs in {folder} (recursive={args.recursive})")

    user_agent = "PDF-Renamer/1.0 (mailto:unknown@example.com)"
    items = compute_preview(
        folder=folder,
        pdfs=pdfs,
        pages=args.pages,
        maxlen=args.maxlen,
        style=args.style,
        no_crossref=args.no_crossref,
        sleep=args.sleep,
        timeout=args.timeout,
        unmatched_dir=args.unmatched_dir,
        user_agent=user_agent,
    )

    for idx, item in enumerate(items, 1):
        print(f"\n[{idx}/{len(items)}] {item.old_name}")
        if item.status == "skip":
            print("  [SKIP] no title found.")
            continue
        if item.status == "ok":
            print("  [OK] already good name.")
            continue
        if args.dry_run:
            print(f"  [DRY] {item.old_name} -> {item.new_path.name}")
        else:
            print(f"  [DO] {item.old_name} -> {item.new_path.name}")

    renamed, skipped = apply_changes(items, args.dry_run)
    print("\nDone.")
    print(f"Renamed: {renamed}")
    print(f"Skipped: {skipped}")


class RenamerGUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("PDF重命名工具")
        self.items: List[PreviewItem] = []
        self.user_agent = "PDF-Renamer/1.0 (mailto:unknown@example.com)"
        self._queue: "queue.Queue[tuple]" = queue.Queue()
        self._worker: Optional[threading.Thread] = None
        self._cancel_event = threading.Event()

        self.var_folder = tk.StringVar()
        self.var_recursive = tk.BooleanVar(value=False)
        self.var_no_crossref = tk.BooleanVar(value=False)
        self.var_pages = tk.StringVar(value="2")
        self.var_maxlen = tk.StringVar(value="140")
        self.var_style = tk.StringVar(value="prefix")
        self.var_unmatched = tk.StringVar(value="")
        self.var_status = tk.StringVar(value="就绪")

        self._build_ui()
        self._set_busy(False)

    def _build_ui(self) -> None:
        self.root.minsize(880, 560)
        self.root.resizable(True, True)
        frm = ttk.Frame(self.root, padding=12)
        frm.grid(row=0, column=0, sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        frm.columnconfigure(0, weight=1)
        frm.columnconfigure(1, weight=1)
        frm.columnconfigure(2, weight=1)

        row = 0
        path_frame = ttk.LabelFrame(frm, text="文件夹")
        path_frame.grid(row=row, column=0, columnspan=3, sticky="ew", padx=2, pady=(0, 8))
        path_frame.columnconfigure(1, weight=1)
        ttk.Label(path_frame, text="路径").grid(row=0, column=0, sticky="w", padx=(8, 6), pady=6)
        ttk.Entry(path_frame, textvariable=self.var_folder).grid(row=0, column=1, sticky="ew", pady=6)
        ttk.Button(path_frame, text="浏览", command=self._browse).grid(row=0, column=2, padx=6, pady=6)
        ttk.Checkbutton(path_frame, text="递归子文件夹", variable=self.var_recursive).grid(
            row=1, column=0, sticky="w", padx=(8, 6), pady=(0, 6)
        )
        ttk.Checkbutton(path_frame, text="不使用Crossref(不联网)", variable=self.var_no_crossref).grid(
            row=1, column=1, sticky="w", pady=(0, 6)
        )

        row += 1
        options = ttk.LabelFrame(frm, text="参数")
        options.grid(row=row, column=0, columnspan=3, sticky="ew", padx=2, pady=(0, 8))
        for col in range(6):
            options.columnconfigure(col, weight=0)
        options.columnconfigure(5, weight=1)
        ttk.Label(options, text="读取页数").grid(row=0, column=0, sticky="w", padx=(8, 6), pady=6)
        ttk.Entry(options, textvariable=self.var_pages, width=8).grid(row=0, column=1, sticky="w", pady=6)
        ttk.Label(options, text="文件名最大长度").grid(row=0, column=2, sticky="w", padx=(12, 6), pady=6)
        ttk.Entry(options, textvariable=self.var_maxlen, width=8).grid(row=0, column=3, sticky="w", pady=6)
        ttk.Label(options, text="未匹配移动到").grid(row=0, column=4, sticky="w", padx=(12, 6), pady=6)
        ttk.Entry(options, textvariable=self.var_unmatched, width=18).grid(row=0, column=5, sticky="w", pady=6)

        ttk.Label(options, text="年份样式").grid(row=1, column=0, sticky="w", padx=(8, 6), pady=(0, 6))
        ttk.Radiobutton(options, text="前缀(年-标题)", variable=self.var_style, value="prefix").grid(
            row=1, column=1, sticky="w", pady=(0, 6)
        )
        ttk.Radiobutton(options, text="后缀(标题-年)", variable=self.var_style, value="suffix").grid(
            row=1, column=2, sticky="w", pady=(0, 6)
        )

        row += 1
        btns = ttk.Frame(frm)
        btns.grid(row=row, column=0, columnspan=3, sticky="ew", padx=2, pady=(0, 8))
        self.btn_list = ttk.Button(btns, text="列出PDF", command=self._list_pdfs)
        self.btn_list.grid(row=0, column=0, padx=2, pady=(0, 4), sticky="w")
        self.btn_preview = ttk.Button(btns, text="预览重命名", command=self._preview)
        self.btn_preview.grid(row=0, column=1, padx=2, pady=(0, 4), sticky="w")
        self.btn_rename = ttk.Button(btns, text="重命名所选", command=self._rename_selected)
        self.btn_rename.grid(row=0, column=2, padx=2, pady=(0, 4), sticky="w")
        self.btn_all = ttk.Button(btns, text="全选", command=self._select_all)
        self.btn_all.grid(row=1, column=0, padx=2, sticky="w")
        self.btn_none = ttk.Button(btns, text="全不选", command=self._select_none)
        self.btn_none.grid(row=1, column=1, padx=2, sticky="w")
        self.btn_invert = ttk.Button(btns, text="反选", command=self._invert)
        self.btn_invert.grid(row=1, column=2, padx=2, sticky="w")
        self.btn_cancel = ttk.Button(btns, text="取消", command=self._cancel)
        self.btn_cancel.grid(row=1, column=3, padx=2, sticky="w")

        row += 1
        tree_frame = ttk.LabelFrame(frm, text="预览列表")
        tree_frame.grid(row=row, column=0, columnspan=3, sticky="nsew", padx=2, pady=(0, 8))
        frm.rowconfigure(row, weight=1)
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(
            tree_frame,
            columns=("apply", "old", "new", "status", "doi", "year"),
            show="headings",
            height=12,
        )
        headings = {
            "apply": "是否",
            "old": "旧文件名",
            "new": "新文件名",
            "status": "状态",
            "doi": "DOI",
            "year": "年份",
        }
        for col, w in [("apply", 60), ("old", 260), ("new", 320), ("status", 80), ("doi", 180), ("year", 60)]:
            self.tree.heading(col, text=headings.get(col, col))
            stretch = col in ("old", "new", "doi")
            self.tree.column(col, width=w, anchor="w", stretch=stretch, minwidth=80)
        self.tree.grid(row=0, column=0, sticky="nsew")
        y_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        x_scroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")
        self.tree.bind("<Double-1>", self._toggle_apply)
        self.tree.bind("<space>", self._toggle_apply)

        row += 1
        status_frame = ttk.Frame(frm)
        status_frame.grid(row=row, column=0, columnspan=3, sticky="ew", padx=2, pady=(0, 6))
        status_frame.columnconfigure(1, weight=1)
        ttk.Label(status_frame, text="进度").grid(row=0, column=0, sticky="w")
        self.progress = ttk.Progressbar(status_frame, mode="determinate")
        self.progress.grid(row=0, column=1, sticky="ew", padx=6)

        row += 1
        ttk.Label(frm, textvariable=self.var_status).grid(row=row, column=0, columnspan=3, sticky="w", padx=4)

        row += 1
        log_frame = ttk.LabelFrame(frm, text="日志")
        log_frame.grid(row=row, column=0, columnspan=3, sticky="nsew", padx=2, pady=(0, 2))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        self.log_text = tk.Text(log_frame, height=7, wrap="word", state="disabled")
        self.log_text.grid(row=0, column=0, sticky="nsew")

    def _log(self, msg: str) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _set_busy(self, busy: bool) -> None:
        state = "disabled" if busy else "normal"
        self.btn_preview.configure(state=state)
        self.btn_list.configure(state=state)
        self.btn_rename.configure(state=state)
        self.btn_all.configure(state=state)
        self.btn_none.configure(state=state)
        self.btn_invert.configure(state=state)
        self.btn_cancel.configure(state="normal" if busy else "disabled")

    def _poll_queue(self) -> None:
        try:
            while True:
                msg = self._queue.get_nowait()
                kind = msg[0]
                if kind == "progress":
                    _, idx, total, name = msg
                    self.progress.configure(maximum=max(1, total))
                    self.progress["value"] = idx
                    self.var_status.set(f"处理中 {idx}/{total}: {name}")
                elif kind == "log":
                    _, text = msg
                    self._log(text)
                elif kind == "done_preview":
                    _, items = msg
                    self.items = items
                    self._refresh_tree()
                    self.var_status.set("预览完成")
                    self._set_busy(False)
                    self._worker = None
                elif kind == "cancelled_preview":
                    self.var_status.set("已取消预览")
                    self._set_busy(False)
                    self._worker = None
                elif kind == "item":
                    _, idx, item = msg
                    if 0 <= idx < len(self.items):
                        self.items[idx] = item
                        self._refresh_row(idx)
                elif kind == "done_rename":
                    _, renamed, skipped = msg
                    self._log(f"完成。重命名: {renamed}, 跳过: {skipped}")
                    self.var_status.set("重命名完成")
                    self._set_busy(False)
                    self._worker = None
                    self.root.after(0, self._preview)
                elif kind == "cancelled_rename":
                    _, renamed, skipped = msg
                    self._log(f"已取消。重命名: {renamed}, 跳过: {skipped}")
                    self.var_status.set("已取消重命名")
                    self._set_busy(False)
                    self._worker = None
                elif kind == "error":
                    _, err = msg
                    if messagebox:
                        messagebox.showerror("错误", err)
                    self.var_status.set("出错")
                    self._set_busy(False)
                    self._worker = None
        except queue.Empty:
            pass

        if self._worker and self._worker.is_alive():
            self.root.after(100, self._poll_queue)

    def _browse(self) -> None:
        if not filedialog:
            return
        path = filedialog.askdirectory()
        if path:
            self.var_folder.set(path)

    def _preview(self) -> None:
        if self._worker and self._worker.is_alive():
            return
        self._cancel_event.clear()
        folder = self.var_folder.get().strip()
        if not folder:
            if messagebox:
                messagebox.showerror("错误", "请选择文件夹。")
            return
        root = Path(folder).expanduser().resolve()
        if not root.exists() or not root.is_dir():
            if messagebox:
                messagebox.showerror("错误", "文件夹不存在。")
            return

        try:
            pages = int(self.var_pages.get())
            maxlen = int(self.var_maxlen.get())
        except ValueError:
            if messagebox:
                messagebox.showerror("错误", "读取页数/最大长度必须是整数。")
            return

        pdfs = collect_pdfs(root, self.var_recursive.get())
        if not pdfs:
            if messagebox:
                messagebox.showinfo("提示", "未检测到PDF。")
            return
        self._log(f"找到 {len(pdfs)} 个 PDF。")
        self.progress.configure(maximum=max(1, len(pdfs)))
        self.progress["value"] = 0
        self.var_status.set("开始预览...")
        self.items = self._pending_items(pdfs)
        self._refresh_tree()
        self._set_busy(True)

        def worker():
            try:
                items = compute_preview(
                    folder=root,
                    pdfs=pdfs,
                    pages=pages,
                    maxlen=maxlen,
                    style=self.var_style.get(),
                    no_crossref=self.var_no_crossref.get(),
                    sleep=0.2,
                    timeout=20,
                    unmatched_dir=self.var_unmatched.get().strip(),
                    user_agent=self.user_agent,
                    progress_cb=lambda i, t, p: self._queue.put(("progress", i, t, p.name)),
                    item_cb=lambda i, t, it: self._queue.put(("item", i - 1, it)),
                    cancel_event=self._cancel_event,
                )
                if self._cancel_event.is_set():
                    self._queue.put(("cancelled_preview",))
                else:
                    self._queue.put(("done_preview", items))
            except Exception as e:
                self._queue.put(("error", str(e)))

        self._worker = threading.Thread(target=worker, daemon=True)
        self._worker.start()
        self._poll_queue()

    def _refresh_tree(self) -> None:
        for i in self.tree.get_children():
            self.tree.delete(i)
        for idx, item in enumerate(self.items):
            apply_text = "是" if item.apply else "否"
            new_name = item.new_path.name if item.new_path else ""
            self.tree.insert(
                "",
                "end",
                iid=str(idx),
                values=(apply_text, item.old_name, new_name, self._status_label(item), item.doi or "", item.year or ""),
            )

    def _pending_items(self, pdfs: List[Path]) -> List[PreviewItem]:
        return [
            PreviewItem(
                pdf=pdf,
                old_name=pdf.name,
                new_path=None,
                doi=None,
                title=None,
                year=None,
                status="pending",
                reason="pending",
                apply=False,
            )
            for pdf in pdfs
        ]

    def _refresh_row(self, idx: int) -> None:
        if str(idx) not in self.tree.get_children():
            return
        item = self.items[idx]
        apply_text = "是" if item.apply else "否"
        new_name = item.new_path.name if item.new_path else ""
        self.tree.item(
            str(idx),
            values=(apply_text, item.old_name, new_name, self._status_label(item), item.doi or "", item.year or ""),
        )

    def _status_label(self, item: PreviewItem) -> str:
        mapping = {
            "pending": "待预览",
            "rename": "重命名",
            "move": "移动",
            "ok": "无需",
            "skip": "跳过",
        }
        return mapping.get(item.status, item.status)

    def _toggle_apply(self, event=None) -> None:
        sel = self.tree.selection()
        if not sel:
            return
        for iid in sel:
            idx = int(iid)
            item = self.items[idx]
            if item.status in ("skip", "ok", "pending"):
                continue
            item.apply = not item.apply
            apply_text = "是" if item.apply else "否"
            new_name = item.new_path.name if item.new_path else ""
            self.tree.item(
                iid,
                values=(apply_text, item.old_name, new_name, self._status_label(item), item.doi or "", item.year or ""),
            )

    def _select_all(self) -> None:
        for idx, item in enumerate(self.items):
            if item.status not in ("skip", "ok", "pending"):
                item.apply = True
        self._refresh_tree()

    def _select_none(self) -> None:
        for item in self.items:
            item.apply = False
        self._refresh_tree()

    def _invert(self) -> None:
        for item in self.items:
            if item.status not in ("skip", "ok", "pending"):
                item.apply = not item.apply
        self._refresh_tree()

    def _cancel(self) -> None:
        if self._worker and self._worker.is_alive():
            self._cancel_event.set()
            self.var_status.set("正在取消...")
            self._log("收到取消请求，正在停止当前任务...")

    def _list_pdfs(self) -> None:
        if self._worker and self._worker.is_alive():
            return
        folder = self.var_folder.get().strip()
        if not folder:
            if messagebox:
                messagebox.showerror("错误", "请选择文件夹。")
            return
        root = Path(folder).expanduser().resolve()
        if not root.exists() or not root.is_dir():
            if messagebox:
                messagebox.showerror("错误", "文件夹不存在。")
            return
        pdfs = collect_pdfs(root, self.var_recursive.get())
        if not pdfs:
            if messagebox:
                messagebox.showinfo("提示", "未检测到PDF。")
            return
        self.items = self._pending_items(pdfs)
        self._refresh_tree()
        self.progress.configure(maximum=max(1, len(pdfs)))
        self.progress["value"] = 0
        self.var_status.set(f"已列出 {len(pdfs)} 个 PDF，点击“预览”开始解析")

    def _rename_selected(self) -> None:
        if self._worker and self._worker.is_alive():
            return
        if not self.items:
            if messagebox:
                messagebox.showinfo("提示", "请先预览。")
            return
        self._cancel_event.clear()
        to_apply = [i for i in self.items if i.apply and i.new_path]
        if not to_apply:
            if messagebox:
                messagebox.showinfo("提示", "没有选中的文件。")
            return
        if messagebox:
            if not messagebox.askyesno("确认", "确认重命名所选文件？"):
                return

        self.progress.configure(maximum=max(1, len(to_apply)))
        self.progress["value"] = 0
        self.var_status.set("开始重命名...")
        self._set_busy(True)

        def worker():
            try:
                renamed, skipped = apply_changes(
                    to_apply,
                    dry_run=False,
                    log=lambda m: self._queue.put(("log", m)),
                    progress_cb=lambda i, t, it: self._queue.put(("progress", i, t, it.pdf.name)),
                    cancel_event=self._cancel_event,
                )
                if self._cancel_event.is_set():
                    self._queue.put(("cancelled_rename", renamed, skipped))
                else:
                    self._queue.put(("done_rename", renamed, skipped))
            except Exception as e:
                self._queue.put(("error", str(e)))

        self._worker = threading.Thread(target=worker, daemon=True)
        self._worker.start()
        self._poll_queue()


def run_gui() -> None:
    if tk is None:
        raise SystemExit("tkinter not available.")
    root = tk.Tk()
    RenamerGUI(root)
    root.mainloop()


def main() -> None:
    ap = argparse.ArgumentParser(
        description="Batch rename PDFs using DOI -> Crossref title + year. Falls back to PDF metadata/title/year guessing."
    )
    ap.add_argument("folder", nargs="?", help="PDF 文件夹路径（可批量）")
    ap.add_argument("--gui", action="store_true", help="Launch GUI")
    ap.add_argument("--recursive", action="store_true", help="递归扫描子文件夹")
    ap.add_argument("--dry-run", action="store_true", help="只打印不改名（强烈建议先用）")
    ap.add_argument("--pages", type=int, default=2, help="读前几页找 DOI/年份（默认2页）")
    ap.add_argument("--maxlen", type=int, default=140, help="文件名最大长度（含 .pdf，默认140）")
    ap.add_argument("--sleep", type=float, default=0.2, help="每次查 Crossref 间隔秒数（默认0.2）")
    ap.add_argument("--timeout", type=int, default=20, help="Crossref 请求超时秒数（默认20）")
    ap.add_argument("--style", choices=["prefix", "suffix"], default="prefix",
                    help="年份位置：prefix=年份在前(默认)，suffix=年份在后")
    ap.add_argument("--no-crossref", action="store_true", help="不联网查 Crossref（只用PDF元数据/页面文本猜）")
    ap.add_argument("--unmatched-dir", default="", help="找不到标题/年份的PDF移动到子目录名（例如: _unmatched），默认不移动")
    args = ap.parse_args()

    if args.gui or not args.folder:
        run_gui()
        return
    run_cli(args)


if __name__ == "__main__":
    main()
