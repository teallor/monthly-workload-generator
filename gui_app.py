from __future__ import annotations

import argparse
import os
import queue
import re
import subprocess
import sys
import threading
import tkinter as tk
from datetime import date
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from main import load_config
from parsers.common import CourseRecord
from parsers.ocr_parser import ocr_self_check
from workload_service import generate_excel, generate_preview
from workload_writer import inspect_template
from workspace_manager import (
    SUPPORTED_MATERIALS,
    collect_folder_files,
    copy_without_overwrite,
    import_material_files,
    load_settings,
    material_item,
    prepare_session,
    save_settings,
    template_month,
)


APP_NAME = "月度工作量表自动生成器"
RESOURCE_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
APP_DIR = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
WORKSPACE = APP_DIR / "workspace"
SETTINGS_PATH = APP_DIR / "app_settings.json"


def split_aliases(value: str) -> list[str]:
    normalized = value.replace("，", ",").replace("、", ",")
    return [item.strip() for item in normalized.split(",") if item.strip()]


def display_hours(value) -> str:
    if value is None:
        return "需确认"
    return f"{value:g}" if isinstance(value, float) else str(value)


def open_path(path: Path):
    if sys.platform == "win32":
        os.startfile(str(path))
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])


class WorkloadGui:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_NAME)
        self.root.geometry("1280x820")
        self.root.minsize(1080, 700)
        self.config = self._load_config()
        self.settings = load_settings(SETTINGS_PATH)
        self.events: queue.Queue = queue.Queue()
        self.materials: list[dict] = []
        self.last_preview: dict | None = None
        self.last_preview_request: dict | None = None
        self.final_excel: Path | None = None
        self.template_usable = False
        self.busy = False
        self.office_available = self._has_microsoft_excel()
        self._build_style()
        self._build_variables()
        self._build_menu()
        self._build_ui()
        self._restore_settings()
        self.root.protocol("WM_DELETE_WINDOW", self.close)
        self.root.after(120, self._poll_events)
        if not self.office_available:
            self.root.after(700, self._show_office_warning)

    def _load_config(self):
        path = APP_DIR / "config.json"
        if not path.exists():
            path = RESOURCE_DIR / "config.json"
        return load_config(path)

    def _build_style(self):
        style = ttk.Style(self.root)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        style.configure("Header.TFrame", background="#174A7E")
        style.configure("AppTitle.TLabel", background="#174A7E", foreground="white", font=("Microsoft YaHei UI", 18, "bold"))
        style.configure("Subtitle.TLabel", background="#174A7E", foreground="#E4EEF8", font=("Microsoft YaHei UI", 10))
        style.configure("Flow.TLabel", background="#E9F2FA", foreground="#174A7E", font=("Microsoft YaHei UI", 10, "bold"), padding=(12, 7))
        style.configure("Step.TLabelframe.Label", font=("Microsoft YaHei UI", 11, "bold"), foreground="#174A7E")
        style.configure("Primary.TButton", font=("Microsoft YaHei UI", 10, "bold"), padding=(15, 8))
        style.configure("Upload.TButton", font=("Microsoft YaHei UI", 10, "bold"), padding=(12, 10))
        style.configure("Hint.TLabel", foreground="#666666", font=("Microsoft YaHei UI", 8))
        style.configure("Value.TLabel", foreground="#222222", font=("Microsoft YaHei UI", 9))
        style.configure("Treeview", rowheight=25, font=("Microsoft YaHei UI", 9))
        style.configure("Treeview.Heading", font=("Microsoft YaHei UI", 9, "bold"))

    def _build_variables(self):
        today = date.today()
        self.template_var = tk.StringVar()
        self.template_name_var = tk.StringVar(value="尚未上传模板")
        self.template_path_var = tk.StringVar(value="-")
        self.template_month_var = tk.StringVar(value="未识别")
        self.template_status_var = tk.StringVar(value="待上传")
        self.output_var = tk.StringVar(value=str(WORKSPACE / "output"))
        self.year_var = tk.StringVar(value=str(today.year))
        self.month_var = tk.StringVar(value=f"{today.month:02d}")
        self.teacher_var = tk.StringVar(value=self.config.get("teacher_name", "黄佳豪"))
        self.aliases_var = tk.StringVar(value=",".join(self.config.get("teacher_aliases", ["黄"])))
        self.ocr_var = tk.BooleanVar(value=bool(self.config.get("enable_ocr", False)))
        self.flow_var = tk.StringVar(value="待导入材料")
        self.status_var = tk.StringVar(value="就绪")
        self.summary_var = tk.StringVar(value="尚未生成预览")
        self.material_count_var = tk.StringVar(value="已导入 0 个材料")

    def _build_menu(self):
        menu = tk.Menu(self.root)
        file_menu = tk.Menu(menu, tearoff=False)
        file_menu.add_command(label="上传模板", command=self.upload_template)
        file_menu.add_command(label="上传材料", command=self.upload_materials)
        file_menu.add_command(label="导入材料文件夹", command=self.import_material_folder)
        file_menu.add_separator()
        file_menu.add_command(label="打开输出目录", command=self.open_output_dir)
        file_menu.add_command(label="退出", command=self.close)
        menu.add_cascade(label="文件", menu=file_menu)
        tool_menu = tk.Menu(menu, tearoff=False)
        tool_menu.add_command(label="生成预览", command=self.start_preview)
        tool_menu.add_command(label="生成最终 Excel", command=self.confirm_write)
        tool_menu.add_separator()
        tool_menu.add_command(label="清空材料列表", command=self.clear_materials)
        tool_menu.add_command(label="清空日志", command=self.clear_log)
        tool_menu.add_separator()
        tool_menu.add_command(label="OCR 自检", command=self.start_ocr_self_check)
        menu.add_cascade(label="工具", menu=tool_menu)
        help_menu = tk.Menu(menu, tearoff=False)
        help_menu.add_command(label="使用说明", command=self.show_help)
        help_menu.add_command(label="关于", command=self.show_about)
        menu.add_cascade(label="帮助", menu=help_menu)
        self.root.configure(menu=menu)

    def _build_ui(self):
        header = ttk.Frame(self.root, style="Header.TFrame", padding=(18, 12))
        header.pack(fill="x")
        title = ttk.Frame(header, style="Header.TFrame")
        title.pack(side="left", fill="x", expand=True)
        ttk.Label(title, text=APP_NAME, style="AppTitle.TLabel").pack(anchor="w")
        ttk.Label(title, text="自动识别 Word / PDF / Excel / 图片课表，一键生成指定月份工作量表", style="Subtitle.TLabel").pack(anchor="w", pady=(3, 0))
        ttk.Label(header, textvariable=self.flow_var, style="Flow.TLabel").pack(side="right", padx=(15, 0))

        body = ttk.Panedwindow(self.root, orient="horizontal")
        body.pack(fill="both", expand=True, padx=12, pady=10)
        left_holder = ttk.Frame(body)
        left_canvas = tk.Canvas(left_holder, width=400, highlightthickness=0)
        left_scroll = ttk.Scrollbar(left_holder, orient="vertical", command=left_canvas.yview)
        left = ttk.Frame(left_canvas, padding=(2, 0, 8, 0))
        left_window = left_canvas.create_window((0, 0), window=left, anchor="nw")
        left.bind("<Configure>", lambda _: left_canvas.configure(scrollregion=left_canvas.bbox("all")))
        left_canvas.bind("<Configure>", lambda event: left_canvas.itemconfigure(left_window, width=event.width))
        left_canvas.configure(yscrollcommand=left_scroll.set)
        left_canvas.pack(side="left", fill="both", expand=True)
        left_scroll.pack(side="right", fill="y")
        right = ttk.Frame(body)
        body.add(left_holder, weight=0)
        body.add(right, weight=1)

        self._build_template_step(left)
        self._build_material_step(left)
        self._build_settings_step(left)
        self._build_action_step(left)
        self._build_results(right)

        footer = ttk.Frame(self.root, padding=(12, 7))
        footer.pack(fill="x")
        ttk.Label(footer, textvariable=self.status_var, relief="sunken", padding=(9, 5)).pack(side="left", fill="x", expand=True)
        self.exit_button = ttk.Button(footer, text="退出", command=self.close)
        self.open_dir_button = ttk.Button(footer, text="打开输出目录", command=self.open_output_dir)
        self.write_button = ttk.Button(footer, text="生成最终 Excel", style="Primary.TButton", command=self.confirm_write, state="disabled")
        self.preview_button = ttk.Button(footer, text="生成预览", style="Primary.TButton", command=self.start_preview)
        for button in (self.exit_button, self.open_dir_button, self.write_button, self.preview_button):
            button.pack(side="right", padx=(7, 0))

    def _build_template_step(self, parent):
        frame = ttk.LabelFrame(parent, text="步骤 1　上传工作量表模板", style="Step.TLabelframe", padding=10)
        frame.pack(fill="x", pady=(0, 8))
        self.upload_template_button = ttk.Button(frame, text="上传工作量表模板", style="Upload.TButton", command=self.upload_template)
        self.upload_template_button.pack(fill="x", pady=(0, 8))
        info = ttk.Frame(frame)
        info.pack(fill="x")
        for row, (label, variable) in enumerate((("文件名", self.template_name_var), ("模板月份", self.template_month_var), ("识别状态", self.template_status_var))):
            ttk.Label(info, text=f"{label}：", width=10).grid(row=row, column=0, sticky="nw", pady=1)
            ttk.Label(info, textvariable=variable, style="Value.TLabel", wraplength=260).grid(row=row, column=1, sticky="w", pady=1)
        ttk.Label(info, text="文件路径：", width=10).grid(row=3, column=0, sticky="nw", pady=1)
        ttk.Entry(info, textvariable=self.template_path_var, state="readonly").grid(row=3, column=1, sticky="ew", pady=1)
        info.columnconfigure(1, weight=1)
        actions = ttk.Frame(frame)
        actions.pack(fill="x", pady=(8, 0))
        ttk.Button(actions, text="更换模板", command=self.upload_template).pack(side="left")
        ttk.Button(actions, text="打开所在文件夹", command=self.open_template_folder).pack(side="left", padx=5)
        ttk.Button(actions, text="清除模板", command=self.clear_template).pack(side="left")
        ttk.Button(actions, text="按模板生成下月", command=self.set_month_after_template).pack(side="left", padx=(5, 0))

    def _build_material_step(self, parent):
        frame = ttk.LabelFrame(parent, text="步骤 2　上传课表依据材料", style="Step.TLabelframe", padding=10)
        frame.pack(fill="both", pady=(0, 8))
        buttons = ttk.Frame(frame)
        buttons.pack(fill="x")
        self.upload_material_button = ttk.Button(buttons, text="上传课表/通知材料", style="Upload.TButton", command=self.upload_materials)
        self.upload_material_button.pack(side="left", fill="x", expand=True)
        self.folder_button = ttk.Button(buttons, text="导入材料文件夹", style="Upload.TButton", command=self.import_material_folder)
        self.folder_button.pack(side="left", fill="x", expand=True, padx=(6, 0))
        ttk.Label(frame, textvariable=self.material_count_var, style="Hint.TLabel").pack(anchor="w", pady=(6, 3))
        columns = ("序号", "文件名", "类型", "大小", "状态", "路径")
        tree_frame = ttk.Frame(frame)
        tree_frame.pack(fill="both", expand=True)
        self.material_tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=4, selectmode="extended")
        widths = {"序号": 45, "文件名": 165, "类型": 48, "大小": 65, "状态": 70, "路径": 260}
        for column in columns:
            self.material_tree.heading(column, text=column)
            self.material_tree.column(column, width=widths[column], minwidth=40, stretch=column in {"文件名", "路径"})
        xbar = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.material_tree.xview)
        ybar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.material_tree.yview)
        self.material_tree.configure(xscrollcommand=xbar.set, yscrollcommand=ybar.set)
        self.material_tree.grid(row=0, column=0, sticky="nsew")
        ybar.grid(row=0, column=1, sticky="ns")
        xbar.grid(row=1, column=0, sticky="ew")
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)
        tools = ttk.Frame(frame)
        tools.pack(fill="x", pady=(6, 0))
        ttk.Button(tools, text="删除选中", command=self.remove_selected_materials).pack(side="left")
        ttk.Button(tools, text="清空全部", command=self.clear_materials).pack(side="left", padx=5)
        ttk.Button(tools, text="打开所在位置", command=self.open_material_folder).pack(side="left")
        manual = ttk.Frame(frame)
        manual.pack(fill="x", pady=(6, 0))
        ttk.Button(manual, text="手工补录课程", command=self.add_manual_course).pack(side="left", fill="x", expand=True)

    def _build_settings_step(self, parent):
        frame = ttk.LabelFrame(parent, text="步骤 3　设置生成月份和教师", style="Step.TLabelframe", padding=10)
        frame.pack(fill="x", pady=(0, 8))
        month = ttk.Frame(frame)
        month.pack(fill="x", pady=2)
        ttk.Label(month, text="生成月份：", width=10).pack(side="left")
        ttk.Combobox(month, textvariable=self.year_var, width=7, state="readonly", values=[str(year) for year in range(date.today().year - 5, date.today().year + 6)]).pack(side="left")
        ttk.Label(month, text="年").pack(side="left", padx=(3, 8))
        ttk.Combobox(month, textvariable=self.month_var, width=5, state="readonly", values=[f"{m:02d}" for m in range(1, 13)]).pack(side="left")
        ttk.Label(month, text="月").pack(side="left", padx=3)
        teacher = ttk.Frame(frame)
        teacher.pack(fill="x", pady=2)
        ttk.Label(teacher, text="教师姓名：", width=10).pack(side="left")
        ttk.Entry(teacher, textvariable=self.teacher_var, width=14).pack(side="left")
        ttk.Label(teacher, text="简称/别名：").pack(side="left", padx=(10, 3))
        ttk.Entry(teacher, textvariable=self.aliases_var, width=16).pack(side="left", fill="x", expand=True)
        ocr_row = ttk.Frame(frame)
        ocr_row.pack(fill="x", pady=(4, 3))
        ttk.Checkbutton(ocr_row, text="启用图片 OCR（可选，较慢）", variable=self.ocr_var).pack(side="left")
        self.ocr_check_button = ttk.Button(ocr_row, text="检测 OCR", command=self.start_ocr_self_check)
        self.ocr_check_button.pack(side="right")
        ttk.Label(frame, text="黄 = 黄佳豪；黄、王 / 黄/王 均识别为黄佳豪参与\n总复习一天 = 理论8课时；实训一天 = 实训8课时", style="Hint.TLabel", justify="left").pack(anchor="w")

    def _build_action_step(self, parent):
        frame = ttk.LabelFrame(parent, text="步骤 4　生成工作量表", style="Step.TLabelframe", padding=10)
        frame.pack(fill="x")
        output = ttk.Frame(frame)
        output.pack(fill="x", pady=(0, 7))
        ttk.Label(output, text="保存位置：").pack(side="left")
        ttk.Entry(output, textvariable=self.output_var, state="readonly").pack(side="left", fill="x", expand=True, padx=4)
        ttk.Button(output, text="更改", command=self.choose_output_dir).pack(side="left")
        ttk.Label(frame, text="先生成预览并核对结果，确认无误后再生成最终 Excel。", style="Hint.TLabel").pack(anchor="w")

    def _build_results(self, parent):
        ttk.Label(parent, text="结果展示", font=("Microsoft YaHei UI", 12, "bold"), foreground="#174A7E").pack(anchor="w", pady=(0, 5))
        notebook = ttk.Notebook(parent)
        notebook.pack(fill="both", expand=True)
        self.preview_tree = self._tree_tab(notebook, "待写入预览", [
            ("序号", 48), ("日期", 88), ("星期", 65), ("写入模块", 78), ("班级/项目", 210),
            ("课程名称", 250), ("教师", 105), ("课时", 55), ("分类", 62), ("子分类", 62),
            ("来源文件", 210), ("命中方式", 105), ("是否需确认", 85), ("备注", 230),
        ])
        self.preview_tree.tag_configure("training", background="#F3FAF5")
        self.preview_tree.tag_configure("assessment", background="#EEF5FC")
        self.preview_tree.tag_configure("confirm", background="#FFF0D9", foreground="#9A4A00")
        self.excluded_tree = self._tree_tab(notebook, "被排除课程", [
            ("日期", 88), ("课程名称", 260), ("教师", 105), ("来源文件", 230), ("排除原因", 260), ("备注", 230),
        ])
        self.summary_tree = self._tree_tab(notebook, "考核汇总", [
            ("班级/项目", 300), ("子分类", 80), ("明细日期", 230), ("明细数量", 80), ("汇总课时", 80), ("预计写入位置", 140),
        ])
        log_frame = ttk.Frame(notebook)
        notebook.add(log_frame, text="解析日志")
        self.log_text = tk.Text(log_frame, wrap="word", font=("Consolas", 10), padx=8, pady=8)
        ybar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=ybar.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        ybar.pack(side="right", fill="y")
        self.summary_label = ttk.Label(parent, textvariable=self.summary_var, padding=(6, 6), anchor="w")
        self.summary_label.pack(fill="x")

    def _tree_tab(self, notebook, title, columns):
        frame = ttk.Frame(notebook)
        notebook.add(frame, text=title)
        tree = ttk.Treeview(frame, columns=[name for name, _ in columns], show="headings")
        xbar = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        ybar = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(xscrollcommand=xbar.set, yscrollcommand=ybar.set)
        for name, width in columns:
            tree.heading(name, text=name)
            tree.column(name, width=width, minwidth=45, stretch=False)
        tree.grid(row=0, column=0, sticky="nsew")
        ybar.grid(row=0, column=1, sticky="ns")
        xbar.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        return tree

    def _target_month(self) -> str:
        return f"{int(self.year_var.get()):04d}-{int(self.month_var.get()):02d}"

    def upload_template(self):
        source = filedialog.askopenfilename(title="上传工作量表模板", filetypes=[("Excel 工作簿", "*.xls *.xlsx")])
        if not source:
            return
        self._set_busy(True, "正在识别工作量表模板...")
        threading.Thread(target=self._template_worker, args=(Path(source),), daemon=True).start()

    def _template_worker(self, source: Path):
        try:
            destination = copy_without_overwrite(source, WORKSPACE / "templates")
            try:
                layout = inspect_template(destination)
                usable, error = True, ""
            except Exception as exc:
                layout, usable, error = None, False, str(exc)
            self.events.put(("template", {"path": destination, "usable": usable, "layout": layout, "error": error}))
        except Exception as exc:
            self.events.put(("error", self._friendly_error("模板上传失败", exc)))

    def _on_template(self, payload):
        path = payload["path"]
        self.template_var.set(str(path))
        self.template_name_var.set(path.name)
        self.template_path_var.set(str(path))
        month = template_month(path.name)
        self.template_month_var.set(month)
        self.template_usable = payload["usable"]
        if payload["usable"]:
            self.template_status_var.set("已识别为工作量表模板，可用")
            if month != "未识别":
                self.set_month_after_template()
            self.flow_var.set("模板已就绪，待导入材料")
        else:
            self.template_status_var.set("无法识别，不可用")
            messagebox.showwarning("模板识别失败", "该文件不像工作量表模板，请确认是否包含“工作量表”或指定模板表头。\n\n详细信息：" + payload["error"])
        self._set_busy(False, "模板识别完成")
        self._save_settings()

    def set_month_after_template(self):
        month = self.template_month_var.get()
        if month == "未识别":
            messagebox.showwarning("无法推断下月", "模板文件名中没有识别到年份和月份，请手动选择生成月份。")
            return
        year, number = (int(part) for part in month.split("-"))
        if number == 12:
            year, number = year + 1, 1
        else:
            number += 1
        available = [str(value) for value in range(date.today().year - 5, date.today().year + 6)]
        if str(year) not in available:
            messagebox.showwarning("年份超出范围", f"建议月份为 {year}年{number}月，请手动确认年份。")
        self.year_var.set(str(year))
        self.month_var.set(f"{number:02d}")
        self._invalidate_preview()

    def clear_template(self):
        self.template_var.set("")
        self.template_name_var.set("尚未上传模板")
        self.template_path_var.set("-")
        self.template_month_var.set("未识别")
        self.template_status_var.set("待上传")
        self.template_usable = False
        self._invalidate_preview()

    def open_template_folder(self):
        path = Path(self.template_var.get())
        if path.exists():
            self._safe_open(path.parent)
        else:
            messagebox.showwarning("模板不存在", "模板路径不存在，请重新上传模板。")

    def upload_materials(self):
        paths = filedialog.askopenfilenames(
            title="上传课表/通知材料",
            filetypes=[("支持的材料", "*.doc *.docx *.pdf *.xls *.xlsx *.jpg *.jpeg *.png"), ("所有文件", "*.*")],
        )
        if paths:
            self._start_material_import([Path(path) for path in paths])

    def import_material_folder(self):
        folder = filedialog.askdirectory(title="导入材料文件夹")
        if folder:
            self._start_material_import(collect_folder_files(Path(folder)))

    def _start_material_import(self, paths):
        self._set_busy(True, "正在导入材料...")
        target = WORKSPACE / "materials" / self._target_month()
        threading.Thread(target=self._material_worker, args=(paths, target, list(self.materials)), daemon=True).start()

    def _material_worker(self, paths, target, existing):
        try:
            added, skipped = import_material_files(paths, target, existing)
            self.events.put(("materials", {"added": added, "skipped": skipped}))
        except Exception as exc:
            self.events.put(("error", self._friendly_error("材料导入失败", exc)))

    def _on_materials(self, payload):
        for item in payload["added"]:
            path = Path(item["path"])
            if path.suffix.lower() in {".jpg", ".jpeg", ".png"} and not self.ocr_var.get():
                item["status"] = "图片已导入，未启用 OCR，需手工补录"
            elif path.suffix.lower() in {".jpg", ".jpeg", ".png"} and self.ocr_var.get():
                item["status"] = "待 OCR 解析"
        self.materials.extend(payload["added"])
        self._refresh_material_tree()
        self._invalidate_preview()
        self.flow_var.set("已导入材料，待生成预览")
        self._set_busy(False, f"已导入 {len(payload['added'])} 个文件")
        if payload["skipped"]:
            messagebox.showinfo("部分文件未重复导入", "\n".join(payload["skipped"][:15]))
        self._save_settings()

    def _refresh_material_tree(self):
        self.material_tree.delete(*self.material_tree.get_children())
        for index, item in enumerate(self.materials, 1):
            self.material_tree.insert("", "end", iid=str(index - 1), values=(index, item["name"], item["type"], item["size"], item["status"], item["path"]))
        supported = sum(
            item.get("supported", False) and (item.get("is_manual") or Path(item["path"]).exists())
            for item in self.materials
        )
        self.material_count_var.set(f"已导入 {len(self.materials)} 个材料，其中 {supported} 个可解析")

    def remove_selected_materials(self):
        selected = sorted((int(iid) for iid in self.material_tree.selection()), reverse=True)
        for index in selected:
            if 0 <= index < len(self.materials):
                self.materials.pop(index)
        self._refresh_material_tree()
        self._invalidate_preview()
        self._save_settings()

    def clear_materials(self):
        if self.materials and not messagebox.askyesno("清空材料列表", "确认清空当前材料列表？工作区中的文件不会被删除。"):
            return
        self.materials.clear()
        self._refresh_material_tree()
        self._invalidate_preview()
        self.flow_var.set("待导入材料")
        self._save_settings()

    def open_material_folder(self):
        selected = self.material_tree.selection()
        if not selected:
            messagebox.showwarning("未选择材料", "请先在材料列表中选择一个文件。")
            return
        item = self.materials[int(selected[0])]
        if item.get("is_manual"):
            messagebox.showinfo("手工补录", "该记录来自界面手工补录，没有对应的外部文件。")
            return
        path = Path(item["path"])
        self._safe_open(path.parent if path.exists() else WORKSPACE / "materials")

    def _normalized_manual_date(self, raw):
        value = raw.strip()
        if re.fullmatch(r"\d{1,2}", value):
            return f"{self._target_month()}-{int(value):02d}"
        if re.fullmatch(r"\d{4}-\d{1,2}-\d{1,2}", value):
            year, month, day = (int(part) for part in value.split("-"))
            return f"{year:04d}-{month:02d}-{day:02d}"
        raise ValueError("日期请输入日号（如 12）或完整日期（如 2026-08-12）")

    def _append_manual_record(self, record):
        self.materials.append({
            "name": f"{record.date} {record.course_name}", "type": "手工", "size": "-",
            "status": "待解析", "path": "界面手工补录", "original_path": "界面手工补录",
            "supported": True, "is_manual": True, "record": record.to_dict(),
        })

    def add_manual_course(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("手工补录课程")
        dialog.transient(self.root)
        dialog.grab_set()
        fields = {
            "日期": tk.StringVar(), "班级/项目": tk.StringVar(), "课程名称": tk.StringVar(),
            "教师": tk.StringVar(value=self.teacher_var.get()), "课时": tk.StringVar(value="8"),
        }
        ttk.Label(dialog, text="手工补录课程", font=("Microsoft YaHei UI", 14, "bold"), foreground="#174A7E").grid(row=0, column=0, columnspan=2, sticky="w", padx=18, pady=(16, 10))
        for row, (label, variable) in enumerate(fields.items(), 1):
            ttk.Label(dialog, text=label + "：").grid(row=row, column=0, sticky="e", padx=(18, 6), pady=4)
            ttk.Entry(dialog, textvariable=variable, width=48).grid(row=row, column=1, sticky="ew", padx=(0, 18), pady=4)
        ttk.Label(dialog, text="日期可只填日号；培训/考核分类仍由现有业务规则自动判断。", style="Hint.TLabel").grid(row=6, column=0, columnspan=2, padx=18, pady=5)
        buttons = ttk.Frame(dialog)
        buttons.grid(row=7, column=0, columnspan=2, sticky="e", padx=18, pady=(8, 16))
        def save():
            try:
                course_date = self._normalized_manual_date(fields["日期"].get())
                hours = float(fields["课时"].get())
                if hours.is_integer():
                    hours = int(hours)
                if not fields["课程名称"].get().strip() or not fields["教师"].get().strip():
                    raise ValueError("课程名称和教师不能为空")
                record = CourseRecord(
                    source_file="界面手工补录", date=course_date, start_time="", end_time="",
                    course_name=fields["课程名称"].get().strip(), teacher=fields["教师"].get().strip(),
                    hours=hours, project=fields["班级/项目"].get().strip() or "手工补录课程",
                    context="由用户在图形界面手工补录", confidence=1.0,
                )
                self._append_manual_record(record)
                self._refresh_material_tree()
                self._invalidate_preview()
                dialog.destroy()
            except Exception as exc:
                messagebox.showwarning("补录内容不正确", str(exc), parent=dialog)
        ttk.Button(buttons, text="取消", command=dialog.destroy).pack(side="right")
        ttk.Button(buttons, text="保存补录", style="Primary.TButton", command=save).pack(side="right", padx=(0, 7))

    def choose_output_dir(self):
        selected = filedialog.askdirectory(title="选择结果保存位置", initialdir=self.output_var.get())
        if selected:
            self.output_var.set(selected)
            self._invalidate_preview()
            self._save_settings()

    def _validate(self):
        template = Path(self.template_var.get()) if self.template_var.get() else None
        if not template or not template.exists():
            messagebox.showwarning("没有选择模板", "请先点击“上传工作量表模板”。")
            return None
        if not self.template_usable:
            messagebox.showwarning("模板不可用", "当前文件未通过模板识别，请更换工作量表模板。")
            return None
        supported = [item for item in self.materials if item.get("supported") and (item.get("is_manual") or Path(item["path"]).exists())]
        if not supported:
            messagebox.showwarning("没有上传材料", "请先上传课表、通知、PDF、Excel 或图片材料。")
            return None
        if not self.teacher_var.get().strip():
            messagebox.showwarning("没有教师姓名", "请输入要统计的教师姓名。")
            return None
        try:
            target_month = self._target_month()
        except (ValueError, TypeError):
            messagebox.showwarning("月份不正确", "请选择正确的年份和月份。")
            return None
        output = Path(self.output_var.get()).expanduser()
        try:
            output.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            messagebox.showerror("输出目录不可用", f"无法使用结果保存位置：{exc}\n建议选择“文档”中的文件夹。")
            return None
        return {
            "template": template, "materials": supported, "target_month": target_month,
            "output": output.resolve(), "teacher": self.teacher_var.get().strip(),
            "aliases": split_aliases(self.aliases_var.get()), "ocr": self.ocr_var.get(),
        }

    def start_preview(self):
        validated = self._validate()
        if not validated:
            return
        for item in self.materials:
            if item.get("supported"):
                suffix = Path(item["path"]).suffix.lower()
                item["status"] = "待 OCR 解析" if self.ocr_var.get() and suffix in {".jpg", ".jpeg", ".png"} else "待解析"
        self._refresh_material_tree()
        self._set_busy(True, "正在解析材料并生成预览...")
        self.flow_var.set("正在生成预览")
        threading.Thread(target=self._preview_worker, args=(validated,), daemon=True).start()

    def _preview_worker(self, validated):
        try:
            file_materials = [item for item in validated["materials"] if not item.get("is_manual")]
            manual_records = [CourseRecord(**item["record"]) for item in validated["materials"] if item.get("is_manual")]
            session = prepare_session(WORKSPACE, validated["target_month"], file_materials)
            result = generate_preview(
                base_dir=APP_DIR, input_dir=session, output_dir=validated["output"],
                target_month_raw=validated["target_month"], teacher_name=validated["teacher"],
                teacher_aliases=validated["aliases"], template_path=validated["template"],
                template_keyword=self.config.get("template_keyword", "工作量表"), enable_ocr=validated["ocr"],
                extra_records=manual_records,
            )
            result["material_count"] = len(validated["materials"])
            result["request_signature"] = self._current_signature(validated)
            self.events.put(("preview", result))
        except Exception as exc:
            self.events.put(("error", self._friendly_error("生成预览失败", exc)))

    def _current_signature(self, validated=None):
        return {
            "template": str((validated["template"] if validated else Path(self.template_var.get())).resolve()),
            "materials": tuple(sorted(
                ("manual:" + str(item.get("record"))) if item.get("is_manual") else str(Path(item["path"]).resolve())
                for item in self.materials if item.get("supported") and (item.get("is_manual") or Path(item["path"]).exists())
            )),
            "target_month": validated["target_month"] if validated else self._target_month(),
            "output": str((validated["output"] if validated else Path(self.output_var.get())).resolve()),
            "teacher": validated["teacher"] if validated else self.teacher_var.get().strip(),
            "aliases": tuple(validated["aliases"] if validated else split_aliases(self.aliases_var.get())),
            "ocr": validated["ocr"] if validated else self.ocr_var.get(),
        }

    def _on_preview(self, result):
        self.last_preview = result
        self.last_preview_request = result["request_signature"]
        self.final_excel = None
        failed_names = set()
        ocr_failed_names = set()
        for warning in result["warnings"]:
            if "解析失败" in warning:
                failed_names.add(warning.split(":", 1)[0])
            if "OCR失败" in warning or "未能识别出有效课程" in warning:
                ocr_failed_names.add(warning.split(":", 1)[0])
        scanned_names = {path.name for path in result["scanned"]}
        for item in self.materials:
            if not item.get("supported"):
                item["status"] = "不支持"
            elif item.get("is_manual"):
                item["status"] = "已解析"
            elif item["name"] in ocr_failed_names:
                item["status"] = "OCR 失败，可手工补录"
            elif Path(item["path"]).suffix.lower() in {".jpg", ".jpeg", ".png"} and not self.ocr_var.get():
                item["status"] = "图片已导入，未启用 OCR，需手工补录"
            elif Path(item["path"]).suffix.lower() in {".jpg", ".jpeg", ".png"} and self.ocr_var.get() and item["name"] in scanned_names:
                item["status"] = "OCR 已解析"
            elif item["name"] in failed_names:
                item["status"] = "解析失败"
            elif item["name"] in scanned_names or Path(item["path"]).name in scanned_names:
                item["status"] = "已解析"
            elif Path(item["path"]).exists():
                item["status"] = "已解析"
        self._refresh_material_tree()
        self._fill_results(result)
        self.flow_var.set("预览已生成，待确认")
        self._set_busy(False, "预览已完成")
        self.write_button.configure(state="normal")
        self._save_settings()
        references = [record for record in result["excluded"] if record.status == "参考" and record.date]
        ocr_failures = [warning for warning in result["warnings"] if "OCR失败" in warning or "未能识别出有效课程" in warning]
        if ocr_failures:
            messagebox.showwarning(
                "图片 OCR 未识别出有效课程",
                "图片 OCR 未能识别出有效课程，请使用手工补录。\n\n" + "\n".join(ocr_failures),
            )
        elif not result["included"] and references:
            detected = "、".join(sorted({record.date[:7] for record in references}))
            messagebox.showwarning("目标月份可能选错", f"当前月份没有待写入课程，但材料中检测到其他月份：{detected}。\n请检查生成月份后重新预览。")
        else:
            messagebox.showinfo("预览生成完成", f"已识别 {len(result['included'])} 条待写入课程。\n\n请在右侧表格中核对后再生成最终 Excel。")

    def _fill_results(self, result):
        for tree in (self.preview_tree, self.excluded_tree, self.summary_tree):
            tree.delete(*tree.get_children())
        for index, record in enumerate(result["included"], 1):
            weekday = ""
            try:
                weekday = "星期" + "一二三四五六日"[date.fromisoformat(record.date).weekday()]
            except (ValueError, TypeError):
                pass
            tag = "confirm" if record.needs_confirmation or record.hours is None else ("training" if record.category == "培训" else "assessment")
            self.preview_tree.insert("", "end", values=(
                index, record.date, weekday, record.write_module or record.category, record.project,
                record.course_name, record.teacher, display_hours(record.hours), record.category,
                record.subcategory, record.source_file, record.teacher_match_type,
                "需确认" if record.needs_confirmation or record.hours is None else "否",
                record.confirmation_note or record.context,
            ), tags=(tag,))
        for record in result["excluded"]:
            self.excluded_tree.insert("", "end", values=(
                record.date, record.course_name, record.teacher, record.source_file,
                record.exclusion_reason or record.status, record.confirmation_note or record.context,
            ))
        for item in result["assessment_summary"]:
            self.summary_tree.insert("", "end", values=(
                item["班级/项目"], item["子分类"], item.get("明细日期", ""),
                f"{item['明细条数']}条", f"{display_hours(item['课时'])}课时", item["预计写入单元格"],
            ))
        training = [record for record in result["included"] if record.category == "培训"]
        assessment = [record for record in result["included"] if record.category == "考核"]
        training_hours = sum(record.hours or 0 for record in training)
        assessment_hours = sum(record.hours or 0 for record in assessment)
        self.summary_var.set(f"待写入 {len(result['included'])} 条　｜　培训 {len(training)} 条 / {training_hours:g} 课时　｜　考核 {len(assessment)} 条 / {assessment_hours:g} 课时")
        log_lines = [*result["log_lines"], "", f"预览文件: {result['paths']['preview']}", f"排除文件: {result['paths']['excluded']}", f"解析日志: {result['paths']['log']}"]
        self.log_text.delete("1.0", "end")
        self.log_text.insert("end", "\n".join(log_lines))
        if result["warnings"]:
            self.log_text.insert("end", "\n\n提示：\n" + "\n".join(f"- {warning}" for warning in result["warnings"]))
        self.log_text.see("end")

    def confirm_write(self):
        validated = self._validate()
        if not validated:
            return
        if not self.last_preview:
            messagebox.showwarning("找不到预览结果", "请先点击“生成预览”。")
            return
        if self._current_signature(validated) != self.last_preview_request:
            messagebox.showwarning("参数已经改变", "模板、材料、月份、教师或保存位置已经改变，请重新生成预览。")
            return
        unresolved = [record for record in self.last_preview["included"] if record.needs_confirmation or record.hours is None]
        if unresolved:
            messagebox.showwarning("仍有记录需要确认", f"当前仍有 {len(unresolved)} 条记录需要人工确认，请先处理后再生成 Excel。")
            return
        included = self.last_preview["included"]
        training = [record for record in included if record.category == "培训"]
        assessment = [record for record in included if record.category == "考核"]
        training_hours = sum(record.hours or 0 for record in training)
        assessment_hours = sum(record.hours or 0 for record in assessment)
        output_name = self.last_preview["month_output_dir"]
        details = (
            f"目标月份：{int(self.year_var.get())}年{int(self.month_var.get())}月\n"
            f"模板文件：{Path(self.template_var.get()).name}\n"
            f"材料数量：{self.last_preview['material_count']}个\n"
            f"待写入课程：{len(included)}条\n"
            f"培训：{len(training)}条 / {training_hours:g}课时\n"
            f"考核：{len(assessment)}条 / {assessment_hours:g}课时\n"
            f"输出路径：{output_name}\n\n是否确认生成？"
        )
        if not self._confirmation_dialog(details):
            return
        self._set_busy(True, "正在写入 Excel...")
        self.flow_var.set("正在生成 Excel")
        threading.Thread(target=self._write_worker, args=(validated,), daemon=True).start()

    def _confirmation_dialog(self, details):
        dialog = tk.Toplevel(self.root)
        dialog.title("确认生成工作量表")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)
        result = {"ok": False}
        ttk.Label(dialog, text="确认生成工作量表", font=("Microsoft YaHei UI", 14, "bold"), foreground="#174A7E").pack(anchor="w", padx=20, pady=(18, 8))
        ttk.Label(dialog, text=details, justify="left", wraplength=620).pack(anchor="w", padx=20, pady=(0, 15))
        buttons = ttk.Frame(dialog, padding=(20, 0, 20, 18))
        buttons.pack(fill="x")
        def accept():
            result["ok"] = True
            dialog.destroy()
        ttk.Button(buttons, text="取消", command=dialog.destroy).pack(side="right")
        ttk.Button(buttons, text="确认生成", style="Primary.TButton", command=accept).pack(side="right", padx=(0, 8))
        dialog.protocol("WM_DELETE_WINDOW", dialog.destroy)
        dialog.update_idletasks()
        dialog.geometry(f"+{self.root.winfo_rootx()+180}+{self.root.winfo_rooty()+130}")
        self.root.wait_window(dialog)
        return result["ok"]

    def _write_worker(self, validated):
        try:
            result = generate_excel(
                base_dir=APP_DIR, input_dir=self.last_preview["input_dir"], output_dir=validated["output"],
                target_month_raw=validated["target_month"], teacher_name=validated["teacher"],
                template_path=validated["template"], template_keyword=self.config.get("template_keyword", "工作量表"),
            )
            self.events.put(("write", result))
        except Exception as exc:
            self.events.put(("error", self._friendly_error("生成 Excel 失败", exc)))

    def _on_write(self, result):
        self.final_excel = Path(result["output"])
        self.log_text.insert("end", f"\n\nExcel 文件: {self.final_excel}\n写入培训 {result['training_count']} 条，考核 {result['assessment_count']} 条。\n")
        self.log_text.see("end")
        self.flow_var.set("Excel 已生成")
        self._set_busy(False, "已完成")
        self._success_dialog(self.final_excel)

    def _success_dialog(self, path):
        dialog = tk.Toplevel(self.root)
        dialog.title("生成完成")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)
        ttk.Label(dialog, text="生成完成", font=("Microsoft YaHei UI", 15, "bold"), foreground="#19713B").pack(anchor="w", padx=20, pady=(18, 8))
        ttk.Label(dialog, text=f"工作量表已生成：\n\n{path}", justify="left", wraplength=650).pack(anchor="w", padx=20, pady=(0, 15))
        buttons = ttk.Frame(dialog, padding=(20, 0, 20, 18))
        buttons.pack(fill="x")
        ttk.Button(buttons, text="关闭", command=dialog.destroy).pack(side="right")
        ttk.Button(buttons, text="打开输出文件夹", command=lambda: self._safe_open(path.parent)).pack(side="right", padx=7)
        ttk.Button(buttons, text="打开 Excel", style="Primary.TButton", command=lambda: self._safe_open(path)).pack(side="right")
        dialog.update_idletasks()
        dialog.geometry(f"+{self.root.winfo_rootx()+180}+{self.root.winfo_rooty()+140}")

    def _poll_events(self):
        try:
            while True:
                kind, payload = self.events.get_nowait()
                if kind == "template":
                    self._on_template(payload)
                elif kind == "materials":
                    self._on_materials(payload)
                elif kind == "preview":
                    self._on_preview(payload)
                elif kind == "write":
                    self._on_write(payload)
                elif kind == "error":
                    self._on_error(payload)
                elif kind == "ocr_check":
                    self._on_ocr_self_check(payload)
        except queue.Empty:
            pass
        self.root.after(120, self._poll_events)

    def _set_busy(self, busy, status):
        self.busy = busy
        self.status_var.set(status)
        state = "disabled" if busy else "normal"
        for button in (self.preview_button, self.upload_template_button, self.upload_material_button, self.folder_button, self.open_dir_button, self.ocr_check_button):
            button.configure(state=state)
        self.write_button.configure(state="disabled" if busy or not self.last_preview else "normal")

    def _invalidate_preview(self):
        self.last_preview = None
        self.last_preview_request = None
        self.final_excel = None
        if hasattr(self, "write_button"):
            self.write_button.configure(state="disabled")

    def _friendly_error(self, title, exc):
        detail = str(exc).strip() or exc.__class__.__name__
        lowered = detail.lower()
        if "permission" in lowered or "拒绝访问" in detail:
            advice = "建议你选择有写入权限的输出文件夹。"
        elif "0x800a03ec" in lowered or "占用" in detail or "正在使用" in detail:
            advice = "建议你先关闭正在打开的工作量表和 Excel 后重试。"
        elif "preview" in lowered or "预览" in detail:
            advice = "建议你重新点击“生成预览”并核对结果。"
        else:
            advice = "建议你检查所选文件是否完整、格式是否受支持，然后重试。"
        return f"{title}：{detail}\n\n{advice}"

    def _has_microsoft_excel(self):
        if sys.platform != "win32":
            return False
        try:
            import winreg
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Excel.Application\CLSID"):
                return True
        except OSError:
            return False

    def _show_office_warning(self):
        messagebox.showwarning(
            "Microsoft Excel 环境提示",
            "未检测到 Microsoft Excel。\n\n当前功能建议安装 Microsoft Excel；仅 WPS 环境下可能无法严格保持 .xls 模板格式。",
        )

    def _on_error(self, message):
        self.flow_var.set("发生错误")
        self._set_busy(False, "出错")
        self.log_text.insert("end", "\n" + message + "\n")
        self.log_text.see("end")
        messagebox.showerror("操作失败", message)

    def _safe_open(self, path):
        try:
            path = Path(path)
            path.mkdir(parents=True, exist_ok=True) if not path.suffix else None
            open_path(path.resolve())
        except Exception as exc:
            messagebox.showerror("无法打开", f"无法打开：{path}\n\n{exc}")

    def open_output_dir(self):
        path = self.last_preview["month_output_dir"] if self.last_preview else Path(self.output_var.get()) / self._target_month()
        self._safe_open(path)

    def clear_log(self):
        self.log_text.delete("1.0", "end")

    def start_ocr_self_check(self):
        if self.busy:
            return
        self._set_busy(True, "正在检测 OCR 组件...")
        threading.Thread(target=self._ocr_self_check_worker, daemon=True).start()

    def _ocr_self_check_worker(self):
        try:
            self.events.put(("ocr_check", ocr_self_check()))
        except Exception as exc:
            self.events.put(("ocr_check", {
                "available": False, "checks": {}, "errors": [str(exc)],
            }))

    def _on_ocr_self_check(self, result):
        self._set_busy(False, "OCR 自检完成")
        lines = [f"{name}：{value}" for name, value in result.get("checks", {}).items()]
        if result.get("available"):
            lines.extend(["", "结论：OCR 可用，可以识别 JPG/PNG 照片课表。"])
            messagebox.showinfo("OCR 自检", "\n".join(lines))
        else:
            lines.extend(["", *result.get("errors", []), "", "OCR 组件不可用。你仍可以上传图片作为核对材料，并使用手工补录功能。"])
            messagebox.showwarning("OCR 自检", "\n".join(lines))

    def show_help(self):
        messagebox.showinfo("使用说明", "1. 上传工作量表模板\n2. 上传课表、通知、PDF、Excel 或图片材料\n3. 选择年份、月份和教师\n4. 点击“生成预览”并核对右侧结果\n5. 点击“生成最终 Excel”\n\n程序会自动管理 workspace，无需手动整理目录。")

    def show_about(self):
        messagebox.showinfo("关于", "月度工作量表自动生成器\n\n适用于培训资源部月度工作量表自动生成\n支持 Word / PDF / Excel / 图片课表解析")

    def _restore_settings(self):
        settings = self.settings
        self.output_var.set(settings.get("output_dir") or str(WORKSPACE / "output"))
        self.year_var.set(str(settings.get("year") or date.today().year))
        self.month_var.set(f"{int(settings.get('month') or date.today().month):02d}")
        self.teacher_var.set(settings.get("teacher_name") or self.teacher_var.get())
        self.aliases_var.set(settings.get("teacher_aliases") or self.aliases_var.get())
        self.ocr_var.set(bool(settings.get("enable_ocr", self.ocr_var.get())))
        template = Path(settings.get("template", "")) if settings.get("template") else None
        if template and template.exists() and template.suffix.lower() in {".xls", ".xlsx"}:
            self.template_var.set(str(template))
            self.template_name_var.set(template.name)
            self.template_path_var.set(str(template))
            self.template_month_var.set(template_month(template.name))
            self.template_status_var.set("最近使用模板，待校验")
            try:
                inspect_template(template)
                self.template_usable = True
                self.template_status_var.set("已识别为工作量表模板，可用")
            except Exception:
                self.template_usable = False
                self.template_status_var.set("最近模板无法读取，请重新选择")
        elif template:
            self.template_status_var.set("路径不存在，请重新选择")
        for saved in settings.get("materials", []):
            path = Path(saved.get("path", ""))
            if path.exists():
                item = material_item(path, original_path=Path(saved.get("original_path", path)))
                item["status"] = "待解析"
            else:
                item = dict(saved)
                item.update(status="路径不存在", supported=False)
            self.materials.append(item)
        self._refresh_material_tree()
        if self.template_usable and self.materials:
            self.flow_var.set("已恢复最近记录，待生成预览")

    def _save_settings(self):
        try:
            save_settings(SETTINGS_PATH, {
                "template": self.template_var.get(),
                "materials": [{"path": item["path"], "original_path": item.get("original_path", item["path"]), "name": item["name"], "type": item["type"], "size": item["size"], "status": item["status"], "supported": item.get("supported", False)} for item in self.materials if not item.get("is_manual")],
                "output_dir": self.output_var.get(),
                "year": self.year_var.get(),
                "month": self.month_var.get(),
                "teacher_name": self.teacher_var.get(),
                "teacher_aliases": self.aliases_var.get(),
                "enable_ocr": self.ocr_var.get(),
            })
        except OSError:
            pass

    def close(self):
        self._save_settings()
        self.root.destroy()


def ensure_workspace():
    for name in ("templates", "materials", "sessions", "output", "logs"):
        (WORKSPACE / name).mkdir(parents=True, exist_ok=True)


def launch():
    ensure_workspace()
    root = tk.Tk()
    WorkloadGui(root)
    root.mainloop()


def smoke_test():
    ensure_workspace()
    root = tk.Tk()
    root.withdraw()
    app = WorkloadGui(root)
    root.update_idletasks()
    assert app.material_tree.winfo_exists()
    assert app.preview_tree.winfo_exists()
    assert app.summary_tree.winfo_exists()
    app._save_settings()
    root.destroy()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=APP_NAME)
    parser.add_argument("--smoke-test", action="store_true", help=argparse.SUPPRESS)
    args = parser.parse_args()
    smoke_test() if args.smoke_test else launch()
