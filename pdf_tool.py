#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF Tool - Trích xuất trang & Gộp PDF
Chạy được trên Windows, macOS (và Linux).

Cài đặt:   pip install pypdf
Chạy:      python pdf_tool.py        (macOS: python3 pdf_tool.py)

Cú pháp chọn trang (đánh số từ 1):
    3           -> trang 3
    1-5         -> trang 1 đến 5
    1,3,7-9     -> trang 1, 3, 7, 8, 9
    10-         -> từ trang 10 đến hết
    -4          -> từ trang 1 đến 4
    all / *     -> tất cả các trang
    5-1         -> đảo ngược: 5,4,3,2,1
"""

import os
import re
import sys
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

from pypdf import PdfReader, PdfWriter
from pypdf.errors import PdfReadError

APP_TITLE = "PDF Tool - Trích xuất & Gộp trang"


# ======================================================================
# PHẦN XỬ LÝ (không phụ thuộc giao diện)
# ======================================================================
class PdfPasswordRequired(Exception):
    """File PDF có mật khẩu và chưa mở được."""


def parse_page_ranges(text, total):
    """Chuyển chuỗi '1-3,5,8-' thành danh sách chỉ số trang (0-based)."""
    text = (text or "").strip().lower()
    if total <= 0:
        raise ValueError("File PDF không có trang nào.")
    if text in ("", "all", "*", "tat ca", "tất cả"):
        return list(range(total))

    pages = []
    for part in re.split(r"[,;\s]+", text):
        if not part:
            continue
        m = re.fullmatch(r"(\d*)\s*[-–:]\s*(\d*)", part)
        if m:
            a, b = m.group(1), m.group(2)
            if a == "" and b == "":
                raise ValueError(f"Khoảng trang không hợp lệ: '{part}'")
            start = int(a) if a else 1
            end = int(b) if b else total
            for v in (start, end):
                if v < 1 or v > total:
                    raise ValueError(
                        f"Trang {v} nằm ngoài phạm vi (file có {total} trang)."
                    )
            step = 1 if start <= end else -1
            pages.extend(range(start - 1, end - 1 + step, step))
        elif part.isdigit():
            v = int(part)
            if v < 1 or v > total:
                raise ValueError(
                    f"Trang {v} nằm ngoài phạm vi (file có {total} trang)."
                )
            pages.append(v - 1)
        else:
            raise ValueError(f"Không hiểu cú pháp: '{part}'")

    if not pages:
        raise ValueError("Chưa chọn trang nào.")
    return pages


def open_reader(path, password_getter=None):
    """Mở PDF, tự xử lý file có mật khẩu."""
    reader = PdfReader(path)
    if reader.is_encrypted:
        ok = reader.decrypt("")  # thử mật khẩu rỗng
        if not ok:
            pw = password_getter(path) if password_getter else None
            if not pw or not reader.decrypt(pw):
                raise PdfPasswordRequired(
                    f"Không mở được file có mật khẩu:\n{os.path.basename(path)}"
                )
    return reader


def extract_pages(reader, indices, out_path):
    """Xuất các trang (theo thứ tự indices) thành 1 file PDF."""
    writer = PdfWriter()
    for i in indices:
        writer.add_page(reader.pages[i])
    with open(out_path, "wb") as f:
        writer.write(f)


def extract_pages_separately(reader, indices, out_dir, stem):
    """Mỗi trang xuất thành 1 file riêng. Trả về danh sách file đã tạo."""
    width = len(str(len(reader.pages)))
    created = []
    for i in indices:
        writer = PdfWriter()
        writer.add_page(reader.pages[i])
        out_path = os.path.join(out_dir, f"{stem}_trang_{str(i + 1).zfill(width)}.pdf")
        with open(out_path, "wb") as f:
            writer.write(f)
        created.append(out_path)
    return created


def merge_pdfs(items, out_path, password_getter=None):
    """
    items: danh sách (đường_dẫn, chuỗi_trang). Gộp theo đúng thứ tự.
    Trả về tổng số trang trong file kết quả.
    """
    writer = PdfWriter()
    total_pages = 0
    for path, range_text in items:
        reader = open_reader(path, password_getter)
        indices = parse_page_ranges(range_text, len(reader.pages))
        for i in indices:
            writer.add_page(reader.pages[i])
        total_pages += len(indices)
    with open(out_path, "wb") as f:
        writer.write(f)
    return total_pages


def open_in_file_manager(path):
    """Mở thư mục / file bằng ứng dụng mặc định của hệ điều hành."""
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception:
        pass


def same_file(a, b):
    try:
        return os.path.samefile(a, b)
    except Exception:
        return os.path.abspath(a) == os.path.abspath(b)


# ======================================================================
# GIAO DIỆN
# ======================================================================
class PdfToolApp(ttk.Frame):
    def __init__(self, master):
        super().__init__(master, padding=10)
        self.master = master
        self.pack(fill="both", expand=True)

        master.title(APP_TITLE)
        master.minsize(680, 520)

        self._setup_style()

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True)
        self.tab_extract = ttk.Frame(nb, padding=14)
        self.tab_merge = ttk.Frame(nb, padding=14)
        nb.add(self.tab_extract, text="  Trích xuất trang  ")
        nb.add(self.tab_merge, text="  Gộp PDF  ")

        self._build_extract_tab()
        self._build_merge_tab()

        self.status = tk.StringVar(value="Sẵn sàng.")
        ttk.Label(self, textvariable=self.status, anchor="w", relief="sunken",
                  padding=(6, 3)).pack(fill="x", pady=(8, 0))

    # ---------- chung ----------
    def _setup_style(self):
        style = ttk.Style()
        # 'clam' hiển thị đồng nhất hơn trên Windows; macOS giữ theme gốc 'aqua'
        if sys.platform.startswith("win"):
            try:
                style.theme_use("vista")
            except tk.TclError:
                pass
        style.configure("Action.TButton", padding=(14, 8), font=("", 11, "bold"))
        style.configure("Hint.TLabel", foreground="#666666")

    def set_status(self, msg):
        self.status.set(msg)
        self.update_idletasks()

    def ask_password(self, path):
        return simpledialog.askstring(
            "File có mật khẩu",
            f"Nhập mật khẩu cho:\n{os.path.basename(path)}",
            show="*", parent=self.master,
        )

    def load_reader(self, path):
        try:
            return open_reader(path, self.ask_password)
        except PdfPasswordRequired as e:
            messagebox.showerror("Lỗi", str(e))
        except (PdfReadError, OSError, ValueError) as e:
            messagebox.showerror("Lỗi", f"Không đọc được file PDF:\n{path}\n\n{e}")
        except Exception as e:  # file hỏng
            messagebox.showerror("Lỗi", f"File PDF bị lỗi:\n{path}\n\n{e}")
        return None

    # ---------- TAB TRÍCH XUẤT ----------
    def _build_extract_tab(self):
        t = self.tab_extract
        t.columnconfigure(1, weight=1)

        self.ex_path = tk.StringVar()
        self.ex_total = tk.IntVar(value=0)
        self.ex_ranges = tk.StringVar()
        self.ex_mode = tk.StringVar(value="combined")
        self.ex_reader = None

        ttk.Label(t, text="1. Chọn file PDF nguồn").grid(row=0, column=0, columnspan=3, sticky="w")
        ttk.Entry(t, textvariable=self.ex_path, state="readonly").grid(
            row=1, column=0, columnspan=2, sticky="ew", pady=(4, 0))
        ttk.Button(t, text="Chọn file...", command=self.ex_choose_file).grid(
            row=1, column=2, padx=(8, 0), pady=(4, 0))

        self.ex_info = ttk.Label(t, text="Chưa chọn file.", style="Hint.TLabel")
        self.ex_info.grid(row=2, column=0, columnspan=3, sticky="w", pady=(4, 14))

        ttk.Label(t, text="2. Nhập các trang muốn trích xuất").grid(
            row=3, column=0, columnspan=3, sticky="w")
        ttk.Entry(t, textvariable=self.ex_ranges).grid(
            row=4, column=0, columnspan=3, sticky="ew", pady=(4, 0))
        ttk.Label(
            t,
            text="Ví dụ:  3   |   1-5   |   1,3,7-9   |   10- (từ trang 10 đến hết)   |   all (tất cả)",
            style="Hint.TLabel",
        ).grid(row=5, column=0, columnspan=3, sticky="w", pady=(4, 14))

        ttk.Label(t, text="3. Cách xuất file").grid(row=6, column=0, columnspan=3, sticky="w")
        ttk.Radiobutton(
            t, value="combined", variable=self.ex_mode,
            text="Gộp các trang đã chọn thành 1 file PDF",
        ).grid(row=7, column=0, columnspan=3, sticky="w", pady=(4, 0))
        ttk.Radiobutton(
            t, value="separate", variable=self.ex_mode,
            text="Tách mỗi trang đã chọn thành 1 file riêng (lưu vào thư mục)",
        ).grid(row=8, column=0, columnspan=3, sticky="w", pady=(2, 18))

        ttk.Button(t, text="Trích xuất", style="Action.TButton",
                   command=self.ex_run).grid(row=9, column=0, columnspan=3, sticky="e")

    def ex_choose_file(self):
        path = filedialog.askopenfilename(
            title="Chọn file PDF",
            filetypes=[("PDF", "*.pdf"), ("Tất cả", "*.*")],
        )
        if not path:
            return
        reader = self.load_reader(path)
        if reader is None:
            return
        self.ex_reader = reader
        self.ex_path.set(path)
        n = len(reader.pages)
        self.ex_total.set(n)
        self.ex_info.config(text=f"File có {n} trang.")
        self.set_status(f"Đã mở: {os.path.basename(path)} ({n} trang)")

    def ex_run(self):
        if self.ex_reader is None or not self.ex_path.get():
            messagebox.showwarning("Thiếu thông tin", "Hãy chọn file PDF nguồn trước.")
            return
        src = self.ex_path.get()
        total = self.ex_total.get()
        try:
            indices = parse_page_ranges(self.ex_ranges.get(), total)
        except ValueError as e:
            messagebox.showerror("Trang không hợp lệ", str(e))
            return

        stem = os.path.splitext(os.path.basename(src))[0]
        try:
            if self.ex_mode.get() == "combined":
                out = filedialog.asksaveasfilename(
                    title="Lưu file PDF kết quả",
                    defaultextension=".pdf",
                    initialfile=f"{stem}_trich_xuat.pdf",
                    initialdir=os.path.dirname(src),
                    filetypes=[("PDF", "*.pdf")],
                )
                if not out:
                    return
                if same_file(out, src):
                    messagebox.showerror("Lỗi", "Không được ghi đè lên file nguồn. Hãy chọn tên khác.")
                    return
                self.set_status("Đang xuất file...")
                extract_pages(self.ex_reader, indices, out)
                self.set_status(f"Xong: {out}")
                self._done(f"Đã xuất {len(indices)} trang ra:\n{out}", out)
            else:
                out_dir = filedialog.askdirectory(
                    title="Chọn thư mục lưu các file", initialdir=os.path.dirname(src))
                if not out_dir:
                    return
                self.set_status("Đang tách trang...")
                created = extract_pages_separately(self.ex_reader, indices, out_dir, stem)
                self.set_status(f"Xong: {len(created)} file trong {out_dir}")
                self._done(f"Đã tạo {len(created)} file trong thư mục:\n{out_dir}", out_dir)
        except PermissionError:
            messagebox.showerror("Lỗi", "Không có quyền ghi file, hoặc file đang được mở ở chương trình khác.")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể xuất file:\n{e}")

    # ---------- TAB GỘP ----------
    def _build_merge_tab(self):
        t = self.tab_merge
        t.columnconfigure(0, weight=1)
        t.rowconfigure(1, weight=1)

        ttk.Label(
            t,
            text="Thêm các file PDF theo thứ tự muốn gộp. Bấm đúp vào một dòng để chọn trang riêng cho file đó.",
            style="Hint.TLabel", wraplength=600, justify="left",
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))

        cols = ("stt", "file", "total", "pages")
        self.tree = ttk.Treeview(t, columns=cols, show="headings", selectmode="extended", height=10)
        self.tree.heading("stt", text="#")
        self.tree.heading("file", text="File")
        self.tree.heading("total", text="Số trang")
        self.tree.heading("pages", text="Trang lấy")
        self.tree.column("stt", width=36, anchor="center", stretch=False)
        self.tree.column("file", width=340, anchor="w")
        self.tree.column("total", width=70, anchor="center", stretch=False)
        self.tree.column("pages", width=120, anchor="center", stretch=False)
        self.tree.grid(row=1, column=0, sticky="nsew")
        sb = ttk.Scrollbar(t, orient="vertical", command=self.tree.yview)
        sb.grid(row=1, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.bind("<Double-1>", lambda e: self.mg_edit_range())

        self.mg_paths = {}  # iid -> đường dẫn đầy đủ

        btns = ttk.Frame(t)
        btns.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        for text, cmd in [
            ("Thêm file...", self.mg_add),
            ("Chọn trang...", self.mg_edit_range),
            ("Lên ↑", lambda: self.mg_move(-1)),
            ("Xuống ↓", lambda: self.mg_move(1)),
            ("Xóa dòng", self.mg_remove),
            ("Xóa hết", self.mg_clear),
        ]:
            ttk.Button(btns, text=text, command=cmd).pack(side="left", padx=(0, 6))

        ttk.Button(t, text="Gộp thành 1 file PDF", style="Action.TButton",
                   command=self.mg_run).grid(row=3, column=0, columnspan=2, sticky="e", pady=(14, 0))

    def _renumber(self):
        for n, iid in enumerate(self.tree.get_children(), 1):
            self.tree.set(iid, "stt", n)

    def mg_add(self):
        paths = filedialog.askopenfilenames(
            title="Chọn các file PDF cần gộp",
            filetypes=[("PDF", "*.pdf"), ("Tất cả", "*.*")],
        )
        for p in paths:
            reader = self.load_reader(p)
            if reader is None:
                continue
            iid = self.tree.insert(
                "", "end",
                values=("", os.path.basename(p), len(reader.pages), "all"),
            )
            self.mg_paths[iid] = p
        self._renumber()
        self.set_status(f"Danh sách có {len(self.tree.get_children())} file.")

    def mg_edit_range(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Chọn trang", "Hãy chọn một dòng trước.")
            return
        iid = sel[0]
        total = int(self.tree.set(iid, "total"))
        current = self.tree.set(iid, "pages")
        while True:
            val = simpledialog.askstring(
                "Chọn trang",
                f"{self.tree.set(iid, 'file')}  ({total} trang)\n\n"
                "Nhập trang muốn lấy (vd: 1-3,5,8-  hoặc  all):",
                initialvalue=current, parent=self.master,
            )
            if val is None:
                return
            val = val.strip() or "all"
            try:
                parse_page_ranges(val, total)
            except ValueError as e:
                messagebox.showerror("Trang không hợp lệ", str(e))
                current = val
                continue
            self.tree.set(iid, "pages", val)
            return

    def mg_move(self, delta):
        sel = list(self.tree.selection())
        if not sel:
            return
        items = self.tree.get_children()
        order = sorted(sel, key=lambda i: items.index(i), reverse=(delta > 0))
        for iid in order:
            idx = self.tree.index(iid)
            new = idx + delta
            if 0 <= new < len(items):
                self.tree.move(iid, "", new)
        self._renumber()

    def mg_remove(self):
        for iid in self.tree.selection():
            self.mg_paths.pop(iid, None)
            self.tree.delete(iid)
        self._renumber()

    def mg_clear(self):
        if self.tree.get_children() and messagebox.askyesno("Xóa hết", "Xóa toàn bộ danh sách?"):
            self.tree.delete(*self.tree.get_children())
            self.mg_paths.clear()

    def mg_run(self):
        iids = self.tree.get_children()
        if len(iids) < 1:
            messagebox.showwarning("Thiếu file", "Hãy thêm ít nhất một file PDF.")
            return
        items = [(self.mg_paths[i], self.tree.set(i, "pages")) for i in iids]

        first_dir = os.path.dirname(items[0][0])
        out = filedialog.asksaveasfilename(
            title="Lưu file PDF đã gộp",
            defaultextension=".pdf",
            initialfile="gop.pdf",
            initialdir=first_dir,
            filetypes=[("PDF", "*.pdf")],
        )
        if not out:
            return
        if any(same_file(out, p) for p, _ in items):
            messagebox.showerror("Lỗi", "File kết quả trùng với một file nguồn. Hãy chọn tên khác.")
            return
        try:
            self.set_status("Đang gộp...")
            n = merge_pdfs(items, out, self.ask_password)
            self.set_status(f"Xong: {out}")
            self._done(f"Đã gộp {len(items)} file, tổng {n} trang:\n{out}", out)
        except PdfPasswordRequired as e:
            messagebox.showerror("Lỗi", str(e))
        except PermissionError:
            messagebox.showerror("Lỗi", "Không có quyền ghi file, hoặc file đang được mở ở chương trình khác.")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể gộp file:\n{e}")

    # ---------- tiện ích ----------
    def _done(self, message, path_to_open):
        if messagebox.askyesno("Hoàn tất", message + "\n\nMở vị trí file?"):
            target = path_to_open if os.path.isdir(path_to_open) else os.path.dirname(path_to_open)
            if sys.platform == "darwin" and os.path.isfile(path_to_open):
                subprocess.Popen(["open", "-R", path_to_open])  # hiển thị file trong Finder
            elif sys.platform.startswith("win") and os.path.isfile(path_to_open):
                subprocess.Popen(["explorer", "/select,", os.path.normpath(path_to_open)])
            else:
                open_in_file_manager(target)


def main():
    root = tk.Tk()
    PdfToolApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
