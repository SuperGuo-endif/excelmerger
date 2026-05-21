# -*- coding: utf-8 -*-
"""
Excel 多表合并工具 v2.0（流式版）
- 文件读取：openpyxl read_only + iter_rows()，流式不过载内存
- 合并存储：内存字典 {pk → {field: val}}，固定内存占用
- 字段合并：每个字段取第一个非空值
- 导出：流式写入 openpyxl，不占内存
"""

import os
import sys
import threading
import tempfile
from tkinter import *
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# =============================================================================
# 数据模型
# =============================================================================

class ExcelFile:
    """Excel文件（流式读取，不全量加载到内存）"""
    def __init__(self, path):
        self.path = path
        self.name = os.path.basename(path)
        self.wb = None
        self.sheets = []
        self.selected_sheet = None
        self.headers = []
        self._load_workbook()

    def _load_workbook(self):
        try:
            # 关键：read_only=True，流式加载不占内存
            self.wb = openpyxl.load_workbook(self.path, data_only=True, read_only=True)
            self.sheets = self.wb.sheetnames
            if self.sheets:
                self.selected_sheet = self.sheets[0]
                self._load_headers()
        except Exception as e:
            raise Exception(f"无法读取文件 {self.name}: {e}")

    def _load_headers(self):
        ws = self.wb[self.selected_sheet]
        rows_iter = ws.iter_rows(values_only=True)
        first_row = next(rows_iter)
        self.headers = [
            str(h).strip() if h is not None else f"列{i}"
            for i, h in enumerate(first_row)
        ]
        # 注意：read_only 模式下 iter 只能用一个，拿到 headers 后建立 generator
        # 用缓存方式重新获取generator
        self._rows_cache = None

    def _get_rows_generator(self):
        """每次调用返回一个新的行迭代器（read_only 模式支持多次迭代）"""
        ws = self.wb[self.selected_sheet]
        rows_iter = ws.iter_rows(values_only=True)
        next(rows_iter)  # 跳过表头
        return rows_iter

    def load_sheet(self, sheet_name):
        """切换sheet"""
        self.selected_sheet = sheet_name
        self._load_headers()
        return self.headers, []

    def iter_rows(self):
        """返回行迭代器"""
        return self._get_rows_generator()

    def close(self):
        if self.wb:
            self.wb.close()


class MergerLogic:
    """流式合并逻辑：边读边合并，不累积原始数据"""

    @staticmethod
    def merge_files_streaming(files, primary_key, progress_callback=None):
        """
        流式合并核心：
        - 逐文件、逐行读取（内存只保留 merged index）
        - 内存占用 = O(合并后行数 × 字段数)，与源文件总大小无关
        """
        # Step 1: 收集所有字段（只扫描表头，不读数据）
        all_fields_ordered = []
        field_set = set()

        for f in files:
            for h in f.headers:
                if h not in field_set:
                    field_set.add(h)
                    all_fields_ordered.append(h)

        # Step 2: 建立主键索引（内存字典）
        # merged_index: {pk_str: {field: (value, source_file_idx)}}
        # source_file_idx 小的优先，用于"取第一个非空值"
        merged_index = {}

        # 估算总行数（用于进度）
        file_row_counts = []
        for f in files:
            try:
                ws = f.wb[f.selected_sheet]
                cnt = ws.max_row if hasattr(ws, 'max_row') and ws.max_row else 0
            except:
                cnt = 0
            file_row_counts.append(max(cnt - 1, 0))

        total_ops = sum(file_row_counts)
        completed_ops = [0]

        for f_idx, f in enumerate(files):
            pk_col_idx = None
            for idx, h in enumerate(f.headers):
                if h == primary_key:
                    pk_col_idx = idx
                    break
            if pk_col_idx is None:
                raise Exception(f"文件 {f.name} 中找不到主键字段「{primary_key}」")

            for row in f.iter_rows():
                completed_ops[0] += 1
                progress_callback and progress_callback(completed_ops[0], total_ops)

                pk_val = row[pk_col_idx]
                pk_str = str(pk_val) if pk_val is not None else "__NULL__"

                # 初始化或更新该主键行
                if pk_str not in merged_index:
                    merged_index[pk_str] = {"_pk_val": pk_val, "_min_fidx": f_idx, "_fields": {}}
                else:
                    # 取更小 file_idx 的值优先
                    if f_idx < merged_index[pk_str]["_min_fidx"]:
                        merged_index[pk_str]["_pk_val"] = pk_val
                        merged_index[pk_str]["_min_fidx"] = f_idx

                # 每个字段：只写入第一个非空值
                for col_idx, header in enumerate(f.headers):
                    if header in merged_index[pk_str]["_fields"]:
                        continue  # 已有值，跳过
                    cell_val = row[col_idx]
                    if cell_val is not None and str(cell_val).strip() != "":
                        merged_index[pk_str]["_fields"][header] = cell_val

        # Step 3: 生成合并结果
        merged_data = []
        for pk_str, entry in merged_index.items():
            row_dict = dict(entry["_fields"])
            row_dict[primary_key] = entry["_pk_val"]
            merged_data.append(row_dict)

        stats = {
            "total_files": len(files),
            "total_rows_before": sum(file_row_counts),
            "total_rows_after": len(merged_data),
            "fields_count": len(all_fields_ordered),
            "pk_duplicates": sum(1 for entry in merged_index.values() if entry["_min_fidx"] > 0),
        }

        return all_fields_ordered, merged_data, stats


# =============================================================================
# GUI 应用
# =============================================================================

class ExcelMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel 多表合并工具 v2.0")
        self.root.geometry("900x700")
        self.root.configure(bg="#f0f0f0")

        self.files = []
        self.all_fields = []
        self.primary_key = None
        self.merged_headers = []
        self.merged_data = []

        self._build_ui()

    def _build_ui(self):
        header_frame = Frame(self.root, bg="#2c3e50", height=50)
        header_frame.pack(fill=X)
        header_frame.pack_propagate(False)
        Label(header_frame, text="Excel 多表合并工具 v2.0", font=("微软雅黑", 16, "bold"),
              fg="white", bg="#2c3e50").pack(pady=12)

        main_frame = Frame(self.root, bg="#f0f0f0")
        main_frame.pack(fill=BOTH, expand=True, padx=10, pady=10)

        # 左侧
        left_frame = Frame(main_frame, bg="white")
        left_frame.pack(side=LEFT, fill=Y, padx=(0, 10))
        Label(left_frame, text="已添加文件", font=("微软雅黑", 11, "bold"),
              bg="white").pack(pady=(10, 5))

        list_frame = Frame(left_frame, bg="white")
        list_frame.pack(fill=Y, expand=True, padx=10)
        self.file_listbox = Listbox(list_frame, width=35, height=20,
                                    font=("微软雅黑", 10), selectmode=EXTENDED)
        self.file_listbox.pack(side=LEFT, fill=Y)
        scroll_file = Scrollbar(list_frame, command=self.file_listbox.yview)
        scroll_file.pack(side=RIGHT, fill=Y)
        self.file_listbox.configure(yscrollcommand=scroll_file.set)

        btn_frame = Frame(left_frame, bg="white")
        btn_frame.pack(pady=10, padx=10)
        Button(btn_frame, text="添加文件", command=self.add_files,
               bg="#27ae60", fg="white", font=("微软雅黑", 10), width=10).pack(side=LEFT, padx=3)
        Button(btn_frame, text="删除选中", command=self.remove_selected,
               bg="#e74c3c", fg="white", font=("微软雅黑", 10), width=10).pack(side=LEFT, padx=3)

        # 中间
        mid_frame = Frame(main_frame, bg="#f0f0f0")
        Label(mid_frame, text="选择每张表的Sheet", font=("微软雅黑", 11, "bold"),
              bg="#f0f0f0").pack(pady=(0, 5))
        self.sheet_container = Frame(mid_frame, bg="#f0f0f0")
        self.sheet_container.pack(fill=X, pady=(0, 10))

        Label(mid_frame, text="所有字段列表", font=("微软雅黑", 11, "bold"),
              bg="#f0f0f0").pack(pady=(0, 5))
        field_list_frame = Frame(mid_frame, bg="white")
        field_list_frame.pack(fill=X)
        self.field_listbox = Listbox(field_list_frame, width=30, height=8,
                                     font=("微软雅黑", 9), selectmode=BROWSE)
        self.field_listbox.pack(side=LEFT, fill=X, expand=True)
        scroll_field = Scrollbar(field_list_frame, command=self.field_listbox.yview)
        scroll_field.pack(side=RIGHT, fill=Y)
        self.field_listbox.configure(yscrollcommand=scroll_field.set)

        pk_frame = Frame(mid_frame, bg="#f0f0f0")
        pk_frame.pack(fill=X, pady=(15, 5))
        Label(pk_frame, text="选择主键字段：", font=("微软雅黑", 11),
              bg="#f0f0f0").pack(side=LEFT)
        self.pk_var = StringVar()
        self.pk_combo = ttk.Combobox(pk_frame, textvariable=self.pk_var,
                                      font=("微软雅黑", 10), width=20, state="readonly")
        self.pk_combo.pack(side=LEFT, padx=5)
        self.pk_combo.bind("<<ComboboxSelected>>", self.on_pk_selected)
        mid_frame.pack(side=LEFT, fill=Y, padx=(0, 10))

        # 右侧
        right_frame = Frame(main_frame, bg="white")
        Label(right_frame, text="合并预览", font=("微软雅黑", 11, "bold"),
              bg="white").pack(pady=(10, 5))
        self.stats_label = Label(right_frame, text="请先添加文件并选择主键",
                                  font=("微软雅黑", 9), bg="white", fg="#7f8c8d",
                                  anchor=W, justify=LEFT)
        self.stats_label.pack(fill=X, padx=10)

        preview_frame = Frame(right_frame)
        preview_frame.pack(fill=BOTH, expand=True, padx=10, pady=5)
        self.preview_tree = ttk.Treeview(preview_frame, show="headings", height=15)
        preview_tree_scroll_y = Scrollbar(preview_frame, command=self.preview_tree.yview)
        preview_tree_scroll_x = Scrollbar(preview_frame, orient=HORIZONTAL,
                                           command=self.preview_tree.xview)
        self.preview_tree.configure(yscrollcommand=preview_tree_scroll_y.set,
                                     xscrollcommand=preview_tree_scroll_x.set)
        self.preview_tree.grid(row=0, column=0, sticky=N+S+E+W)
        preview_tree_scroll_y.grid(row=0, column=1, sticky=N+S)
        preview_tree_scroll_x.grid(row=1, column=0, sticky=E+W)
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)

        op_frame = Frame(right_frame, bg="white")
        op_frame.pack(pady=10, padx=10)
        self.merge_btn = Button(op_frame, text="执行合并", command=self.do_merge,
                                 bg="#2980b9", fg="white", font=("微软雅黑", 11), width=12,
                                 state=DISABLED)
        self.merge_btn.pack(side=LEFT, padx=5)
        self.export_btn = Button(op_frame, text="导出Excel", command=self.export_excel,
               bg="#8e44ad", fg="white", font=("微软雅黑", 11), width=12,
               state=DISABLED)
        self.export_btn.pack(side=LEFT, padx=5)
        right_frame.pack(side=LEFT, fill=BOTH, expand=True)

        # 状态栏
        self.status_var = StringVar()
        self.status_var.set("就绪")
        status_bar = Label(self.root, textvariable=self.status_var,
                           font=("微软雅黑", 9), bd=1, relief=SUNKEN, anchor=W)
        status_bar.pack(side=BOTTOM, fill=X)

        self.progress_frame = Frame(self.root, bg="#f0f0f0", height=30)
        self.progress_frame.pack(fill=X, padx=10, pady=(0, 5))
        self.progress_frame.pack_propagate(False)
        self.progress_label = Label(self.progress_frame, text="", font=("微软雅黑", 9),
                                     bg="#f0f0f0", anchor=W)
        self.progress_label.pack(side=LEFT)
        self.progress_bar = ttk.Progressbar(self.progress_frame, length=300, mode="determinate")
        self.progress_bar.pack(side=RIGHT)

    # -------------------------------------------------------------------------
    # 文件操作
    # -------------------------------------------------------------------------

    def add_files(self):
        paths = filedialog.askopenfilenames(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if not paths:
            return

        for path in paths:
            if any(f.path == path for f in self.files):
                messagebox.showwarning("重复文件", f"已添加：{os.path.basename(path)}")
                continue
            try:
                ef = ExcelFile(path)
                self.files.append(ef)
                self.file_listbox.insert(END, ef.name)
            except Exception as e:
                messagebox.showerror("错误", str(e))

        self.status_var.set(f"已添加 {len(self.files)} 个文件")
        self._refresh_sheet_ui()
        self._load_all_fields()

    def remove_selected(self):
        sel = self.file_listbox.curselection()
        if not sel:
            return
        for idx in reversed(sel):
            self.files[idx].close()
            del self.files[idx]
            self.file_listbox.delete(idx)
        self._refresh_sheet_ui()
        self._load_all_fields()

    # -------------------------------------------------------------------------
    # Sheet选择
    # -------------------------------------------------------------------------

    def _refresh_sheet_ui(self):
        for w in self.sheet_container.winfo_children():
            w.destroy()
        if not self.files:
            return
        for idx, f in enumerate(self.files):
            row = Frame(self.sheet_container, bg="white")
            row.pack(fill=X, pady=2)
            Label(row, text=f.name[:20], font=("微软雅黑", 9),
                  bg="white", width=20, anchor=W).pack(side=LEFT, padx=(5, 3))
            sheet_var = StringVar(value=f.selected_sheet)
            combo = ttk.Combobox(row, textvariable=sheet_var,
                                  values=f.sheets, font=("微软雅黑", 9),
                                  width=15, state="readonly")
            combo.pack(side=LEFT, padx=3)

            def on_change(e, file_idx=idx, combo_ref=combo):
                new_sheet = combo_ref.get()
                self.files[file_idx].load_sheet(new_sheet)
                self._load_all_fields()

            combo.bind("<<ComboboxSelected>>", on_change)

    # -------------------------------------------------------------------------
    # 字段 & 主键
    # -------------------------------------------------------------------------

    def _load_all_fields(self):
        self.all_fields = []
        self.field_listbox.delete(0, END)
        for f in self.files:
            for h in f.headers:
                if h not in self.all_fields:
                    self.all_fields.append(h)
                    self.field_listbox.insert(END, h)
        self.pk_combo["values"] = self.all_fields
        self.pk_var.set("")

    def on_pk_selected(self, event=None):
        self.primary_key = self.pk_var.get()
        if self.primary_key and self.files:
            self.merge_btn.config(state=NORMAL)

    # -------------------------------------------------------------------------
    # 合并
    # -------------------------------------------------------------------------

    def _update_preview(self, headers, data, max_rows=50):
        self.preview_tree.delete(*self.preview_tree.get_children())
        for col in self.preview_tree["columns"]:
            self.preview_tree.heading(col, text="")
            self.preview_tree.column(col, width=0)
        if not headers:
            return
        self.preview_tree["columns"] = headers
        col_width = min(120, max(60, 800 // len(headers)))
        for h in headers:
            self.preview_tree.heading(h, text=h)
            self.preview_tree.column(h, width=col_width, anchor=W)
        for row_data in data[:max_rows]:
            values = [row_data.get(h, "") for h in headers]
            self.preview_tree.insert("", END, values=values)

    def do_merge(self):
        if not self.primary_key or not self.files:
            messagebox.showwarning("条件不足", "请选择主键字段")
            return

        self.status_var.set("正在合并...")
        self.merge_btn.config(state=DISABLED, text="合并中...")

        def merge_thread():
            try:
                headers, data, stats = MergerLogic.merge_files_streaming(
                    self.files,
                    self.primary_key,
                    progress_callback=self._on_merge_progress
                )
                self.merged_headers = headers
                self.merged_data = data

                def update_ui():
                    self._update_preview(headers, data)
                    self.stats_label.config(text=(
                        f"文件数：{stats['total_files']}　"
                        f"合并前总行数：{stats['total_rows_before']}　"
                        f"合并后行数：{stats['total_rows_after']}　"
                        f"主键重复行：{stats['pk_duplicates']}　"
                        f"字段数：{stats['fields_count']}"
                    ))
                    self.export_btn.config(state=NORMAL)
                    self.status_var.set("合并完成")
                    self.merge_btn.config(state=NORMAL, text="执行合并")
                    messagebox.showinfo("完成",
                        f"合并完成！\n合并后 {stats['total_rows_after']} 行，{stats['fields_count']} 个字段")
                self.root.after(0, update_ui)
            except Exception as e:
                def on_error():
                    self.status_var.set("合并失败")
                    self.merge_btn.config(state=NORMAL, text="执行合并")
                    messagebox.showerror("错误", str(e))
                self.root.after(0, on_error)

        threading.Thread(target=merge_thread, daemon=True).start()

    def _on_merge_progress(self, current, total):
        def update_progress():
            pct = int(current / total * 100) if total else 100
            self.progress_bar["value"] = pct
            self.progress_label.config(text=f"正在合并... {current}/{total} ({pct}%)")
        self.root.after(0, update_progress)

    # -------------------------------------------------------------------------
    # 导出（流式写文件）
    # -------------------------------------------------------------------------

    def export_excel(self):
        if not self.merged_data:
            messagebox.showwarning("无数据", "请先执行合并")
            return

        path = filedialog.asksaveasfilename(
            title="导出合并结果",
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx")]
        )
        if not path:
            return

        try:
            self.status_var.set("正在导出...")
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "合并结果"

            header_fill = PatternFill(start_color="2c3e50", end_color="2c3e50", fill_type="solid")
            header_font = Font(bold=True, color="FFFFFF", size=11)
            header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
            border = Border(
                left=Side(style="thin"), right=Side(style="thin"),
                top=Side(style="thin"), bottom=Side(style="thin")
            )
            data_align = Alignment(horizontal="left", vertical="center")

            # 写表头
            for ci, h in enumerate(self.merged_headers, 1):
                cell = ws.cell(row=1, column=ci, value=h)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = header_align
                cell.border = border

            # 流式写数据行
            for ri, row_dict in enumerate(self.merged_data, 2):
                for ci, field in enumerate(self.merged_headers, 1):
                    cell = ws.cell(row=ri, column=ci, value=row_dict.get(field, ""))
                    cell.alignment = data_align
                    cell.border = border

            # 列宽
            for ci, h in enumerate(self.merged_headers, 1):
                ws.column_dimensions[get_column_letter(ci)].width = 20

            ws.freeze_panes = "A2"
            wb.save(path)
            self.status_var.set(f"导出成功：{os.path.basename(path)}")
            messagebox.showinfo("成功", f"已导出：\n{path}")
        except Exception as e:
            self.status_var.set("导出失败")
            messagebox.showerror("导出失败", str(e))


# =============================================================================
# 入口
# =============================================================================

def main():
    root = Tk()
    app = ExcelMergerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()