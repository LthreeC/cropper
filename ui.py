# -*- coding: utf-8 -*-
"""
用户界面模块 - 简洁美观版
"""

import os
import platform
import threading
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog

CURRENT_OS = platform.system()

DETECT_MODE_LABELS = {
    "智能": "smart",
    "简单": "simple",
    "边缘敏感": "edge",
}


def get_output_quality_policy(mode, source_kind, output_format):
    """返回质量控件类型和不会截断的简短说明。"""
    if output_format in ("PDF", "SVG"):
        if output_format == "SVG":
            return (
                "pdf_image_dpi" if mode == "PPT" else "none",
                "仅导出视觉矢量内容；PDF 链接、批注等交互结构不保留。",
            )
        if mode == "PPT":
            return (
                "pdf_image_dpi",
                "文字和形状保持矢量；图片按所选最高 DPI 优化。",
            )
        if source_kind == "raster":
            return "none", "嵌入裁剪后的原图像素，不重新采样。"
        return "none", "只修改页面边界，保留源文件中的矢量和位图。"

    if mode == "FILE" and source_kind == "raster":
        return "none", "保留源图片像素尺寸和原始 DPI，不重新采样。"
    return "dpi", "按所选 DPI 渲染位图；像素尺寸会随页面大小自动计算。"


def parse_processing_numbers(
    padding_value,
    dpi_value,
    pdf_image_dpi_value,
    page_value,
    quality_policy,
    use_page_number,
):
    """只校验当前任务真正生效的数值字段。"""
    from processor import validate_dpi, validate_padding, validate_page_number

    padding = validate_padding(padding_value)
    dpi = validate_dpi(dpi_value) if quality_policy == "dpi" else 300
    pdf_image_dpi = (
        validate_dpi(pdf_image_dpi_value, "PDF 图片 DPI")
        if quality_policy == "pdf_image_dpi"
        else 300
    )
    page_num = validate_page_number(page_value) if use_page_number else 1
    return padding, dpi, pdf_image_dpi, page_num


class CropperApp(ttk.Frame):
    """主应用界面"""
    
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.is_processing = False
        self.source_files = []
        self.transparent_files = []
        
        self._setup_styles()
        self._init_variables()
        self._create_ui()
        self._on_mode_change()
        self._on_format_change()
    
    def _setup_styles(self):
        """设置紧凑、清晰的桌面样式。"""
        style = ttk.Style()

        if CURRENT_OS == "Windows":
            self.font_base = ("Microsoft YaHei UI", 10)
            self.font_title = ("Microsoft YaHei UI", 10, "bold")
            self.font_big = ("Microsoft YaHei UI", 12, "bold")
            self.font_log = ("Consolas", 9)
        else:
            self.font_base = ("SF Pro Display", 12)
            self.font_title = ("SF Pro Display", 12, "bold")
            self.font_big = ("SF Pro Display", 14, "bold")
            self.font_log = ("Menlo", 10)

        style.configure(".", font=self.font_base)
        style.configure("TLabelframe.Label", font=self.font_title)
        style.configure("TButton", font=self.font_base, padding=(10, 6))
        style.configure("Compact.TButton", font=self.font_base, padding=(8, 4))
        style.configure("Big.TButton", font=self.font_big, padding=(24, 10))
        style.configure("Hint.TLabel", foreground="#666666")

    def _init_variables(self):
        """初始化变量"""
        self.mode_var = tk.StringVar(value="PPT")
        self.scope_var = tk.StringVar(value="CURRENT")
        
        self.detect_mode_var = tk.StringVar(value="smart")
        self.detect_mode_display_var = tk.StringVar(value="智能")
        self.threshold_var = tk.IntVar(value=250)
        self.sensitivity_var = tk.IntVar(value=15)
        self.padding_var = tk.StringVar(value="2")
        
        self.output_format_var = tk.StringVar(value="PDF")
        self.output_dir_var = tk.StringVar(value="<与源文件同目录>")
        self.dpi_var = tk.StringVar(value="300")
        self.pdf_image_dpi_var = tk.StringVar(value="300")
        
        self.file_path_var = tk.StringVar()
        self.page_num_var = tk.IntVar(value=1)
        self.total_pages_var = tk.IntVar(value=1)
        
        self.status_var = tk.StringVar(value="就绪")
        self.progress_var = tk.DoubleVar(value=0)

        self.transparent_file_path_var = tk.StringVar()
        self.transparent_info_var = tk.StringVar(value="支持 SVG、PNG、JPG、BMP、TIFF、WebP 等格式")
        self.transparent_color_mode_var = tk.StringVar(value="corners")
        self.transparent_custom_color_var = tk.StringVar(value="#FFFFFF")
        self.transparent_tolerance_var = tk.IntVar(value=18)
        self.transparent_tolerance_label_var = tk.StringVar(value="18")
        self.transparent_feather_var = tk.IntVar(value=1)
        self.transparent_feather_label_var = tk.StringVar(value="1")
        self.transparent_edge_only_var = tk.BooleanVar(value=True)
        self.transparent_format_var = tk.StringVar(value="PNG")
        self.transparent_output_dir_var = tk.StringVar(value="<与源文件同目录>")
        self.transparent_status_var = tk.StringVar(value="就绪")
        self.transparent_progress_var = tk.DoubleVar(value=0)

        self.detect_advanced_visible = False
        self.transparent_advanced_visible = False
        self.main_log_visible = False
        self.transparent_log_visible = False
    
    def _create_ui(self):
        """创建界面。"""
        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True, padx=12, pady=12)
        self._main_frame = notebook

        main = ttk.Frame(notebook, padding=(14, 12))
        transparent_tab = ttk.Frame(notebook, padding=(14, 12))
        notebook.add(main, text="白边裁剪")
        notebook.add(transparent_tab, text="背景透明")

        self._create_source_section(main)
        self._create_detect_section(main)
        self._create_output_section(main)
        self._create_action_section(main)
        self._create_log_section(main)

        self._create_transparency_tab(transparent_tab)

        # 延迟检查 pymupdf 可用性（不阻塞窗口显示）
        self.after(0, self._deferred_pymupdf_check)

    def _create_source_section(self, parent):
        """来源和处理范围合并为一个任务区。"""
        frame = ttk.LabelFrame(parent, text="任务", padding=(14, 10))
        frame.pack(fill="x", pady=(0, 10))

        mode_row = ttk.Frame(frame)
        mode_row.pack(fill="x")
        ttk.Label(mode_row, text="来源:", width=7).pack(side="left")
        ttk.Radiobutton(
            mode_row,
            text="PowerPoint",
            variable=self.mode_var,
            value="PPT",
            command=self._on_mode_change,
        ).pack(side="left", padx=(0, 20))
        ttk.Radiobutton(
            mode_row,
            text="本地文件",
            variable=self.mode_var,
            value="FILE",
            command=self._on_mode_change,
        ).pack(side="left")
        ttk.Label(
            mode_row,
            text="PDF / SVG / 图片",
            style="Hint.TLabel",
        ).pack(side="left", padx=(8, 0))

        self.source_detail_slot = ttk.Frame(frame)
        self.source_detail_slot.pack(fill="x", pady=(8, 0))
        self.source_detail_slot.columnconfigure(0, weight=1)

        self.ppt_frame = ttk.Frame(self.source_detail_slot)
        ppt_row = ttk.Frame(self.ppt_frame)
        ppt_row.pack(fill="x")
        self.btn_connect = ttk.Button(
            ppt_row,
            text="检测连接",
            command=self._check_ppt,
            style="Compact.TButton",
        )
        self.btn_connect.pack(side="left")
        self.ppt_status = ttk.Label(
            ppt_row,
            text="",
            foreground="gray",
            justify="left",
            wraplength=520,
        )
        self.ppt_status.pack(side="left", fill="x", expand=True, padx=(12, 0))

        self.file_frame = ttk.Frame(self.source_detail_slot)
        file_row = ttk.Frame(self.file_frame)
        file_row.pack(fill="x")
        ttk.Label(file_row, text="文件:", width=7).pack(side="left")
        ttk.Entry(
            file_row,
            textvariable=self.file_path_var,
            font=self.font_base,
        ).pack(side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(
            file_row,
            text="浏览",
            command=self._select_files,
            style="Compact.TButton",
        ).pack(side="left")
        self.file_info = ttk.Label(
            self.file_frame,
            text="支持 PDF、SVG、PNG、JPG、BMP、TIFF、WebP、GIF",
            style="Hint.TLabel",
        )
        self.file_info.pack(anchor="w", padx=(62, 0), pady=(5, 0))

        ttk.Separator(frame).pack(fill="x", pady=(10, 8))

        scope_row = ttk.Frame(frame)
        scope_row.pack(fill="x")
        ttk.Label(scope_row, text="范围:", width=7).pack(side="left")
        self.rb_current = ttk.Radiobutton(
            scope_row,
            text="仅当前页",
            variable=self.scope_var,
            value="CURRENT",
            command=self._on_scope_change,
        )
        self.rb_current.pack(side="left", padx=(0, 20))
        ttk.Radiobutton(
            scope_row,
            text="全部页面",
            variable=self.scope_var,
            value="ALL",
            command=self._on_scope_change,
        ).pack(side="left")

        self.page_frame = ttk.Frame(scope_row)
        ttk.Label(self.page_frame, text="页码:").pack(side="left")
        self.page_spin = ttk.Spinbox(
            self.page_frame,
            from_=1,
            to=9999,
            textvariable=self.page_num_var,
            width=6,
            font=self.font_base,
        )
        self.page_spin.pack(side="left", padx=(6, 0))
        self.page_total = ttk.Label(
            self.page_frame,
            text="",
            style="Hint.TLabel",
        )
        self.page_total.pack(side="left", padx=(6, 0))

    def _create_detect_section(self, parent):
        """常用裁剪项直接显示，少用参数默认折叠。"""
        frame = ttk.LabelFrame(parent, text="裁剪设置", padding=(14, 10))
        frame.pack(fill="x", pady=(0, 10))

        row = ttk.Frame(frame)
        row.pack(fill="x")
        ttk.Label(row, text="检测:").pack(side="left")
        self.detect_mode_combo = ttk.Combobox(
            row,
            textvariable=self.detect_mode_display_var,
            values=list(DETECT_MODE_LABELS),
            state="readonly",
            width=10,
            font=self.font_base,
        )
        self.detect_mode_combo.pack(side="left", padx=(6, 18))
        self.detect_mode_combo.bind(
            "<<ComboboxSelected>>",
            lambda event: self._on_detect_mode_change(),
        )

        ttk.Label(row, text="边缘留白:").pack(side="left")
        ttk.Entry(
            row,
            textvariable=self.padding_var,
            width=6,
            font=self.font_base,
        ).pack(side="left", padx=(6, 4))
        ttk.Label(row, text="px", style="Hint.TLabel").pack(side="left")

        self.btn_detect_advanced = ttk.Button(
            row,
            text="高级参数 ▸",
            command=self._toggle_detect_advanced,
            style="Compact.TButton",
        )
        self.btn_detect_advanced.pack(side="right")

        self.detect_advanced_frame = ttk.Frame(frame)

        threshold_row = ttk.Frame(self.detect_advanced_frame)
        threshold_row.pack(fill="x", pady=(0, 8))
        threshold_row.columnconfigure(1, weight=1)
        ttk.Label(threshold_row, text="白色阈值", width=10).grid(
            row=0, column=0, sticky="w"
        )
        ttk.Scale(
            threshold_row,
            from_=200,
            to=255,
            variable=self.threshold_var,
            orient="horizontal",
        ).grid(row=0, column=1, sticky="ew", padx=(8, 10))
        self.threshold_label = ttk.Label(
            threshold_row,
            text="250",
            width=4,
            font=self.font_title,
        )
        self.threshold_label.grid(row=0, column=2, sticky="e")

        sensitivity_row = ttk.Frame(self.detect_advanced_frame)
        sensitivity_row.pack(fill="x")
        sensitivity_row.columnconfigure(1, weight=1)
        ttk.Label(sensitivity_row, text="敏感度", width=10).grid(
            row=0, column=0, sticky="w"
        )
        ttk.Scale(
            sensitivity_row,
            from_=5,
            to=50,
            variable=self.sensitivity_var,
            orient="horizontal",
        ).grid(row=0, column=1, sticky="ew", padx=(8, 10))
        self.sensitivity_label = ttk.Label(
            sensitivity_row,
            text="15",
            width=4,
            font=self.font_title,
        )
        self.sensitivity_label.grid(row=0, column=2, sticky="e")

        self.threshold_var.trace(
            "w",
            lambda *_: self.threshold_label.config(
                text=str(self.threshold_var.get())
            ),
        )
        self.sensitivity_var.trace(
            "w",
            lambda *_: self.sensitivity_label.config(
                text=str(self.sensitivity_var.get())
            ),
        )

    def _create_output_section(self, parent):
        """输出格式、质量和目录使用稳定的两行布局。"""
        frame = ttk.LabelFrame(parent, text="输出", padding=(14, 10))
        frame.pack(fill="x", pady=(0, 10))

        controls = ttk.Frame(frame)
        controls.pack(fill="x")
        ttk.Label(controls, text="格式:").pack(side="left")
        self.format_combo = ttk.Combobox(
            controls,
            textvariable=self.output_format_var,
            values=["PDF", "SVG", "PNG", "TIFF", "JPEG", "WebP", "GIF"],
            state="readonly",
            width=9,
            font=self.font_base,
        )
        self.format_combo.pack(side="left", padx=(6, 20))
        self.format_combo.bind(
            "<<ComboboxSelected>>",
            lambda event: self._on_format_change(),
        )

        self.quality_slot = ttk.Frame(controls)
        self.quality_slot.pack(side="left", fill="x", expand=True)

        self.dpi_frame = ttk.Frame(self.quality_slot)
        ttk.Label(self.dpi_frame, text="输出 DPI:").pack(side="left")
        ttk.Entry(
            self.dpi_frame,
            textvariable=self.dpi_var,
            width=7,
            font=self.font_base,
        ).pack(side="left", padx=(6, 0))

        self.pdf_dpi_frame = ttk.Frame(self.quality_slot)
        ttk.Label(self.pdf_dpi_frame, text="图片 DPI:").pack(side="left")
        ttk.Combobox(
            self.pdf_dpi_frame,
            textvariable=self.pdf_image_dpi_var,
            values=["300", "450", "600"],
            width=7,
            font=self.font_base,
        ).pack(side="left", padx=(6, 0))

        self.vector_hint = ttk.Label(
            frame,
            text="",
            style="Hint.TLabel",
            justify="left",
            wraplength=640,
        )
        self.vector_hint.pack(fill="x", pady=(7, 0))

        directory_row = ttk.Frame(frame)
        directory_row.pack(fill="x", pady=(9, 0))
        ttk.Label(directory_row, text="保存到:").pack(side="left")
        ttk.Entry(
            directory_row,
            textvariable=self.output_dir_var,
            font=self.font_base,
        ).pack(side="left", fill="x", expand=True, padx=(6, 8))
        ttk.Button(
            directory_row,
            text="选择",
            command=self._select_output_dir,
            style="Compact.TButton",
        ).pack(side="left")

    def _create_action_section(self, parent):
        """状态、进度和唯一主操作。"""
        frame = ttk.Frame(parent)
        frame.pack(fill="x", pady=(0, 8))

        self.status_label = ttk.Label(
            frame,
            textvariable=self.status_var,
            foreground="#0066cc",
        )
        self.status_label.pack(fill="x", pady=(0, 5))

        self.progress = ttk.Progressbar(
            frame,
            variable=self.progress_var,
            maximum=100,
        )
        self.progress.pack(fill="x", pady=(0, 8))

        self.btn_start = ttk.Button(
            frame,
            text="开始处理",
            style="Big.TButton",
            command=self._start_processing,
        )
        self.btn_start.pack(fill="x")

    def _create_log_section(self, parent):
        """日志默认折叠，错误或警告时自动展开。"""
        self.btn_log_toggle = ttk.Button(
            parent,
            text="处理日志 ▸",
            command=self._toggle_main_log,
            style="Compact.TButton",
        )
        self.btn_log_toggle.pack(anchor="w")

        self.log_frame = ttk.LabelFrame(
            parent,
            text="处理日志",
            padding=(8, 8),
        )
        container = ttk.Frame(self.log_frame)
        container.pack(fill="both", expand=True)

        self.log_text = tk.Text(
            container,
            height=7,
            font=self.font_log,
            bg="#fafafa",
            relief="flat",
            wrap="word",
            padx=8,
            pady=6,
            spacing3=3,
        )
        self.log_text.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(
            container,
            orient="vertical",
            command=self.log_text.yview,
        )
        scrollbar.pack(side="right", fill="y")
        self.log_text.config(yscrollcommand=scrollbar.set, state="disabled")

        self.log_text.tag_configure("TIME", foreground="#777777")
        self.log_text.tag_configure("INFO", foreground="#222222")
        self.log_text.tag_configure("SUCCESS", foreground="#16833a")
        self.log_text.tag_configure("WARNING", foreground="#9a6700")
        self.log_text.tag_configure("ERROR", foreground="#c62828")

    def _create_transparency_tab(self, parent):
        """创建背景透明处理页。"""
        self._create_transparency_source_section(parent)
        self._create_transparency_params_section(parent)
        self._create_transparency_output_section(parent)
        self._create_transparency_action_section(parent)
        self._create_transparency_log_section(parent)

        self.transparent_tolerance_var.trace(
            "w",
            lambda *_: self.transparent_tolerance_label_var.set(
                str(self.transparent_tolerance_var.get())
            ),
        )
        self.transparent_feather_var.trace(
            "w",
            lambda *_: self.transparent_feather_label_var.set(
                str(self.transparent_feather_var.get())
            ),
        )
        self._on_transparent_color_mode_change()

    def _create_transparency_source_section(self, parent):
        frame = ttk.LabelFrame(parent, text="来源", padding=(14, 10))
        frame.pack(fill="x", pady=(0, 10))

        row = ttk.Frame(frame)
        row.pack(fill="x")
        ttk.Label(row, text="文件:").pack(side="left")
        ttk.Entry(
            row,
            textvariable=self.transparent_file_path_var,
            font=self.font_base,
        ).pack(side="left", fill="x", expand=True, padx=(6, 8))
        ttk.Button(
            row,
            text="浏览",
            command=self._select_transparent_files,
            style="Compact.TButton",
        ).pack(side="left")

        ttk.Label(
            frame,
            textvariable=self.transparent_info_var,
            style="Hint.TLabel",
        ).pack(anchor="w", pady=(5, 0))

    def _create_transparency_params_section(self, parent):
        frame = ttk.LabelFrame(parent, text="透明设置", padding=(14, 10))
        frame.pack(fill="x", pady=(0, 10))

        color_row = ttk.Frame(frame)
        color_row.pack(fill="x")
        ttk.Label(color_row, text="目标颜色:").pack(side="left")
        for label, value in (
            ("四角平均", "corners"),
            ("左上角", "top_left"),
            ("自定义", "custom"),
        ):
            ttk.Radiobutton(
                color_row,
                text=label,
                variable=self.transparent_color_mode_var,
                value=value,
                command=self._on_transparent_color_mode_change,
            ).pack(side="left", padx=(12, 0))
        self.transparent_custom_color_entry = ttk.Entry(
            color_row,
            textvariable=self.transparent_custom_color_var,
            width=10,
            font=self.font_base,
        )

        option_row = ttk.Frame(frame)
        option_row.pack(fill="x", pady=(8, 0))
        ttk.Checkbutton(
            option_row,
            text="只处理连接到图片边缘的背景",
            variable=self.transparent_edge_only_var,
        ).pack(side="left")
        self.btn_transparent_advanced = ttk.Button(
            option_row,
            text="高级参数 ▸",
            command=self._toggle_transparency_advanced,
            style="Compact.TButton",
        )
        self.btn_transparent_advanced.pack(side="right")

        self.transparent_advanced_frame = ttk.Frame(frame)

        tolerance_row = ttk.Frame(self.transparent_advanced_frame)
        tolerance_row.pack(fill="x", pady=(0, 8))
        tolerance_row.columnconfigure(1, weight=1)
        ttk.Label(tolerance_row, text="颜色容差", width=10).grid(
            row=0, column=0, sticky="w"
        )
        ttk.Scale(
            tolerance_row,
            from_=0,
            to=100,
            variable=self.transparent_tolerance_var,
            orient="horizontal",
        ).grid(row=0, column=1, sticky="ew", padx=(8, 10))
        ttk.Label(
            tolerance_row,
            textvariable=self.transparent_tolerance_label_var,
            width=4,
            font=self.font_title,
        ).grid(row=0, column=2, sticky="e")

        feather_row = ttk.Frame(self.transparent_advanced_frame)
        feather_row.pack(fill="x")
        feather_row.columnconfigure(1, weight=1)
        ttk.Label(feather_row, text="边缘羽化", width=10).grid(
            row=0, column=0, sticky="w"
        )
        ttk.Scale(
            feather_row,
            from_=0,
            to=8,
            variable=self.transparent_feather_var,
            orient="horizontal",
        ).grid(row=0, column=1, sticky="ew", padx=(8, 10))
        ttk.Label(
            feather_row,
            textvariable=self.transparent_feather_label_var,
            width=4,
            font=self.font_title,
        ).grid(row=0, column=2, sticky="e")

    def _create_transparency_output_section(self, parent):
        frame = ttk.LabelFrame(parent, text="输出", padding=(14, 10))
        frame.pack(fill="x", pady=(0, 10))

        format_row = ttk.Frame(frame)
        format_row.pack(fill="x")
        ttk.Label(format_row, text="格式:").pack(side="left")
        ttk.Combobox(
            format_row,
            textvariable=self.transparent_format_var,
            values=["PNG", "WebP"],
            state="readonly",
            width=9,
            font=self.font_base,
        ).pack(side="left", padx=(6, 0))
        ttk.Label(
            format_row,
            text="保留透明通道和原始像素尺寸",
            style="Hint.TLabel",
        ).pack(side="left", padx=(14, 0))

        directory_row = ttk.Frame(frame)
        directory_row.pack(fill="x", pady=(9, 0))
        ttk.Label(directory_row, text="保存到:").pack(side="left")
        ttk.Entry(
            directory_row,
            textvariable=self.transparent_output_dir_var,
            font=self.font_base,
        ).pack(side="left", fill="x", expand=True, padx=(6, 8))
        ttk.Button(
            directory_row,
            text="选择",
            command=self._select_transparent_output_dir,
            style="Compact.TButton",
        ).pack(side="left")

    def _create_transparency_action_section(self, parent):
        frame = ttk.Frame(parent)
        frame.pack(fill="x", pady=(0, 8))

        ttk.Label(
            frame,
            textvariable=self.transparent_status_var,
            foreground="#0066cc",
        ).pack(fill="x", pady=(0, 5))
        ttk.Progressbar(
            frame,
            variable=self.transparent_progress_var,
            maximum=100,
        ).pack(fill="x", pady=(0, 8))
        self.btn_transparent_start = ttk.Button(
            frame,
            text="开始处理透明背景",
            style="Big.TButton",
            command=self._start_transparency_processing,
        )
        self.btn_transparent_start.pack(fill="x")

    def _create_transparency_log_section(self, parent):
        self.btn_transparent_log_toggle = ttk.Button(
            parent,
            text="处理日志 ▸",
            command=self._toggle_transparent_log,
            style="Compact.TButton",
        )
        self.btn_transparent_log_toggle.pack(anchor="w")

        self.transparent_log_frame = ttk.LabelFrame(
            parent,
            text="处理日志",
            padding=(8, 8),
        )
        container = ttk.Frame(self.transparent_log_frame)
        container.pack(fill="both", expand=True)

        self.transparent_log_text = tk.Text(
            container,
            height=7,
            font=self.font_log,
            bg="#fafafa",
            relief="flat",
            wrap="word",
            padx=8,
            pady=6,
            spacing3=3,
        )
        self.transparent_log_text.pack(
            side="left",
            fill="both",
            expand=True,
        )

        scrollbar = ttk.Scrollbar(
            container,
            orient="vertical",
            command=self.transparent_log_text.yview,
        )
        scrollbar.pack(side="right", fill="y")
        self.transparent_log_text.config(
            yscrollcommand=scrollbar.set,
            state="disabled",
        )

        self.transparent_log_text.tag_configure(
            "TIME", foreground="#777777"
        )
        self.transparent_log_text.tag_configure(
            "INFO", foreground="#222222"
        )
        self.transparent_log_text.tag_configure(
            "SUCCESS", foreground="#16833a"
        )
        self.transparent_log_text.tag_configure(
            "WARNING", foreground="#9a6700"
        )
        self.transparent_log_text.tag_configure(
            "ERROR", foreground="#c62828"
        )

    # ========== 延迟检查 ==========
    
    def _deferred_pymupdf_check(self):
        """延迟检查 pymupdf 是否可用"""
        from controllers import has_pymupdf
        if not has_pymupdf():
            warn = ttk.Label(self, text="提示: 安装 pymupdf 可支持 PDF/SVG 处理 (pip install pymupdf)",
                           foreground="orange", font=self.font_base)
            warn.pack(pady=(10, 0), padx=20, anchor="w", before=self._main_frame)
    
    # ========== 事件处理 ==========
    
    def _on_mode_change(self):
        """模式切换。"""
        mode = self.mode_var.get()

        self.ppt_frame.grid_remove()
        self.file_frame.grid_remove()
        self.page_frame.pack_forget()

        if mode == "PPT":
            self.ppt_frame.grid(row=0, column=0, sticky="ew")
            self.rb_current.config(text="仅当前幻灯片")
        else:
            self.file_frame.grid(row=0, column=0, sticky="ew")
            self.rb_current.config(text="仅指定页码")

        self._on_scope_change()
        self._on_format_change()

    def _on_scope_change(self):
        """范围切换。"""
        self.page_frame.pack_forget()
        if self.scope_var.get() == "ALL":
            self.page_spin.config(state="disabled")
            return

        self.page_spin.config(state="normal")
        if self.mode_var.get() == "FILE" and self.source_files:
            from controllers import FileController

            if FileController.is_pdf(self.source_files[0]):
                self.page_frame.pack(side="left", padx=(20, 0))

    def _on_detect_mode_change(self):
        self.detect_mode_var.set(
            DETECT_MODE_LABELS.get(
                self.detect_mode_display_var.get(),
                "smart",
            )
        )

    def _toggle_detect_advanced(self):
        self.detect_advanced_visible = not self.detect_advanced_visible
        if self.detect_advanced_visible:
            self.detect_advanced_frame.pack(fill="x", pady=(10, 0))
            self.btn_detect_advanced.config(text="高级参数 ▾")
        else:
            self.detect_advanced_frame.pack_forget()
            self.btn_detect_advanced.config(text="高级参数 ▸")

    def _current_source_kind(self):
        if self.mode_var.get() == "PPT":
            return "ppt"
        if not self.source_files:
            return "unknown"

        from controllers import FileController

        first = self.source_files[0]
        if FileController.is_pdf(first) or FileController.is_svg(first):
            return "document"
        return "raster"

    def _on_format_change(self):
        """按来源和格式只显示真正有效的质量选项。"""
        self.dpi_frame.pack_forget()
        self.pdf_dpi_frame.pack_forget()

        quality, hint = get_output_quality_policy(
            self.mode_var.get(),
            self._current_source_kind(),
            self.output_format_var.get(),
        )
        if quality == "dpi":
            self.dpi_frame.pack(side="left")
        elif quality == "pdf_image_dpi":
            self.pdf_dpi_frame.pack(side="left")
        self.vector_hint.config(text=hint)

    def _on_transparent_color_mode_change(self):
        self.transparent_custom_color_entry.pack_forget()
        if self.transparent_color_mode_var.get() == "custom":
            self.transparent_custom_color_entry.pack(
                side="left",
                padx=(12, 0),
            )

    def _toggle_transparency_advanced(self):
        self.transparent_advanced_visible = (
            not self.transparent_advanced_visible
        )
        if self.transparent_advanced_visible:
            self.transparent_advanced_frame.pack(fill="x", pady=(10, 0))
            self.btn_transparent_advanced.config(text="高级参数 ▾")
        else:
            self.transparent_advanced_frame.pack_forget()
            self.btn_transparent_advanced.config(text="高级参数 ▸")

    def _toggle_main_log(self):
        self.main_log_visible = not self.main_log_visible
        if self.main_log_visible:
            self.log_frame.pack(fill="both", expand=True, pady=(8, 0))
            self.btn_log_toggle.config(text="处理日志 ▾")
        else:
            self.log_frame.pack_forget()
            self.btn_log_toggle.config(text="处理日志 ▸")

    def _toggle_transparent_log(self):
        self.transparent_log_visible = not self.transparent_log_visible
        if self.transparent_log_visible:
            self.transparent_log_frame.pack(
                fill="both",
                expand=True,
                pady=(8, 0),
            )
            self.btn_transparent_log_toggle.config(text="处理日志 ▾")
        else:
            self.transparent_log_frame.pack_forget()
            self.btn_transparent_log_toggle.config(text="处理日志 ▸")

    def _check_ppt(self):
        """检测 PPT 连接"""
        from controllers import get_ppt_controller

        controller = None
        try:
            controller = get_ppt_controller()
            if controller.check_connection():
                idx, name, path = controller.get_info()
                count = controller.get_slide_count()
                self.ppt_status.config(text=f"已连接: {name} (共{count}页)", foreground="green")
                self.log(f"已连接 PPT: {name}，第 {idx} 页", "SUCCESS")
            else:
                self.ppt_status.config(text="未检测到 PPT", foreground="red")
                self.log("未检测到运行中的 PowerPoint", "WARNING")
        except Exception as e:
            self.ppt_status.config(text="连接失败", foreground="red")
            self.log(f"连接失败: {e}", "ERROR")
        finally:
            if controller is not None:
                controller.close()
    
    def _select_files(self):
        """选择文件"""
        from controllers import FileController
        try:
            import pymupdf
        except ImportError:
            pymupdf = None
        
        filetypes = [
            ("支持的格式", "*.pdf *.svg *.png *.jpg *.jpeg *.bmp *.tiff *.tif *.webp *.gif"),
            ("PDF 文件", "*.pdf"),
            ("SVG 文件", "*.svg"),
            ("图片文件", "*.svg *.png *.jpg *.jpeg *.bmp *.tiff *.tif *.webp *.gif"),
            ("所有文件", "*.*")
        ]
        paths = filedialog.askopenfilenames(filetypes=filetypes)
        
        if paths:
            self.source_files = list(paths)
            
            if len(paths) == 1:
                self.file_path_var.set(paths[0])
                
                # PDF 获取页数
                if FileController.is_pdf(paths[0]) and pymupdf:
                    try:
                        doc = pymupdf.open(paths[0])
                        count = len(doc)
                        doc.close()
                        self.total_pages_var.set(count)
                        self.page_spin.config(to=count)
                        self.page_total.config(text=f"(共 {count} 页)")
                        self.file_info.config(text=f"PDF 文件，共 {count} 页")
                    except:
                        pass
                elif FileController.is_svg(paths[0]):
                    self.page_frame.pack_forget()
                    self.file_info.config(text="SVG 文件")
                else:
                    self.page_frame.pack_forget()
                    self.file_info.config(text="图片文件")
            else:
                self.file_path_var.set(f"[已选择 {len(paths)} 个文件]")
                self.page_frame.pack_forget()
                self.file_info.config(text=f"已选择 {len(paths)} 个文件，将批量处理")

            self._on_scope_change()
            self._on_format_change()
    
    def _select_output_dir(self):
        """选择输出目录"""
        path = filedialog.askdirectory()
        if path:
            self.output_dir_var.set(path)

    def _select_transparent_files(self):
        """选择透明背景处理文件"""
        filetypes = [
            ("支持的图片", "*.svg *.png *.jpg *.jpeg *.bmp *.tiff *.tif *.webp *.gif"),
            ("SVG 文件", "*.svg"),
            ("图片文件", "*.png *.jpg *.jpeg *.bmp *.tiff *.tif *.webp *.gif"),
            ("所有文件", "*.*")
        ]
        paths = filedialog.askopenfilenames(filetypes=filetypes)

        if paths:
            self.transparent_files = list(paths)
            if len(paths) == 1:
                self.transparent_file_path_var.set(paths[0])
                ext = os.path.splitext(paths[0])[1].lower().lstrip(".").upper()
                self.transparent_info_var.set(f"{ext or '图片'} 文件")
            else:
                self.transparent_file_path_var.set(f"[已选择 {len(paths)} 个文件]")
                self.transparent_info_var.set(f"已选择 {len(paths)} 个文件，将批量处理")

    def _select_transparent_output_dir(self):
        """选择透明背景输出目录"""
        path = filedialog.askdirectory()
        if path:
            self.transparent_output_dir_var.set(path)
    
    def log(self, message, level="INFO"):
        """记录日志"""
        import datetime

        if level in ("WARNING", "ERROR") and not self.main_log_visible:
            self._toggle_main_log()
        ts = datetime.datetime.now().strftime("%H:%M:%S")
        
        self.log_text.config(state="normal")
        self.log_text.insert("end", f"[{ts}] ", "TIME")
        self.log_text.insert("end", f"{message}\n", level)
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    def transparent_log(self, message, level="INFO"):
        """记录透明背景日志"""
        import datetime

        if level in ("WARNING", "ERROR") and not self.transparent_log_visible:
            self._toggle_transparent_log()
        ts = datetime.datetime.now().strftime("%H:%M:%S")

        self.transparent_log_text.config(state="normal")
        self.transparent_log_text.insert("end", f"[{ts}] ", "TIME")
        self.transparent_log_text.insert("end", f"{message}\n", level)
        self.transparent_log_text.see("end")
        self.transparent_log_text.config(state="disabled")

    def _set_processing_buttons(self, is_processing):
        state = "disabled" if is_processing else "normal"
        if hasattr(self, "btn_start"):
            self.btn_start.config(state=state)
        if hasattr(self, "btn_transparent_start"):
            self.btn_transparent_start.config(state=state)
    
    def _start_processing(self):
        """开始处理"""
        if self.is_processing:
            return
        
        self.is_processing = True
        self._set_processing_buttons(True)
        self.progress_var.set(0)
        
        # 分隔线
        self.log_text.config(state="normal")
        self.log_text.insert("end", "─" * 50 + "\n", "TIME")
        self.log_text.config(state="disabled")
        
        thread = threading.Thread(target=self._run_processing, daemon=True)
        thread.start()

    def _start_transparency_processing(self):
        """开始处理透明背景"""
        if self.is_processing:
            return

        self.is_processing = True
        self._set_processing_buttons(True)
        self.transparent_progress_var.set(0)

        self.transparent_log_text.config(state="normal")
        self.transparent_log_text.insert("end", "─" * 50 + "\n", "TIME")
        self.transparent_log_text.config(state="disabled")

        thread = threading.Thread(target=self._run_transparency_processing, daemon=True)
        thread.start()
    
    def _run_processing(self):
        """处理线程"""
        from processor import CropProcessor
        
        try:
            def callback(status, log_entry, progress):
                if status:
                    self.after(0, lambda: self.status_var.set(status))
                if log_entry:
                    level, msg = log_entry
                    self.after(0, lambda: self.log(msg, level))
                if progress is not None:
                    self.after(0, lambda: self.progress_var.set(progress))
            
            processor = CropProcessor(callback=callback)
            
            # 构建配置
            mode = self.mode_var.get()
            source_kind = self._current_source_kind()
            quality_policy, _ = get_output_quality_policy(
                mode,
                source_kind,
                self.output_format_var.get(),
            )
            use_page_number = (
                mode == "FILE"
                and self.scope_var.get() == "CURRENT"
                and len(self.source_files) == 1
                and os.path.splitext(self.source_files[0])[1].lower() == ".pdf"
            )
            padding, dpi, pdf_image_dpi, page_num = parse_processing_numbers(
                self.padding_var.get(),
                self.dpi_var.get(),
                self.pdf_image_dpi_var.get(),
                self.page_spin.get(),
                quality_policy,
                use_page_number,
            )

            config = {
                "scope": self.scope_var.get(),
                "output_format": self.output_format_var.get(),
                "output_dir": self.output_dir_var.get() if self.output_dir_var.get() != "<与源文件同目录>" else None,
                "detect_mode": self.detect_mode_var.get(),
                "threshold": self.threshold_var.get(),
                "sensitivity": self.sensitivity_var.get(),
                "padding": padding,
                "dpi": dpi,
                "pdf_image_dpi": pdf_image_dpi,
                "page_num": page_num,
                "source_files": self.source_files,
            }
            
            result = None
            
            if mode == "PPT":
                result = processor.process_ppt(config)
            else:
                if not self.source_files:
                    self.after(0, lambda: self.log("请先选择文件", "ERROR"))
                else:
                    result = processor.process_file(config)
            
            if result and not processor.had_errors:
                self.after(0, lambda: self.progress_var.set(100))
            
            # 打开输出目录
            if result:
                self._open_folder(os.path.dirname(result) if os.path.isfile(result) else result)
        
        except Exception as e:
            self.after(0, lambda: self.log(f"处理失败: {e}", "ERROR"))
            import traceback
            traceback.print_exc()
        
        finally:
            self.is_processing = False
            self.after(0, lambda: self._set_processing_buttons(False))
            self.after(0, lambda: self.status_var.set("就绪"))

    def _run_transparency_processing(self):
        """透明背景处理线程"""
        from processor import CropProcessor

        try:
            def callback(status, log_entry, progress):
                if status:
                    self.after(0, lambda: self.transparent_status_var.set(status))
                if log_entry:
                    level, msg = log_entry
                    self.after(0, lambda: self.transparent_log(msg, level))
                if progress is not None:
                    self.after(0, lambda: self.transparent_progress_var.set(progress))

            processor = CropProcessor(callback=callback)

            config = {
                "source_files": self.transparent_files,
                "output_dir": self.transparent_output_dir_var.get()
                if self.transparent_output_dir_var.get() != "<与源文件同目录>" else None,
                "output_format": self.transparent_format_var.get(),
                "color_mode": self.transparent_color_mode_var.get(),
                "custom_color": self.transparent_custom_color_var.get(),
                "tolerance": self.transparent_tolerance_var.get(),
                "edge_only": self.transparent_edge_only_var.get(),
                "feather": self.transparent_feather_var.get(),
                "dpi": 300,
            }

            result = None
            if not self.transparent_files:
                self.after(0, lambda: self.transparent_log("请先选择文件", "ERROR"))
            else:
                result = processor.process_transparency(config)

            if result and not processor.had_errors:
                self.after(0, lambda: self.transparent_progress_var.set(100))
            if result:
                self._open_folder(result)

        except Exception as e:
            self.after(0, lambda: self.transparent_log(f"处理失败: {e}", "ERROR"))
            import traceback
            traceback.print_exc()

        finally:
            self.is_processing = False
            self.after(0, lambda: self._set_processing_buttons(False))
            self.after(0, lambda: self.transparent_status_var.set("就绪"))
    
    def _open_folder(self, path):
        """打开文件夹"""
        try:
            if CURRENT_OS == "Windows":
                os.startfile(path)
            elif CURRENT_OS == "Darwin":
                subprocess.run(["open", path])
            else:
                subprocess.run(["xdg-open", path])
        except:
            pass
