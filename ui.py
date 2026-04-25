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
        """设置样式 - 大字体"""
        style = ttk.Style()
        
        if CURRENT_OS == "Windows":
            self.font_base = ("Microsoft YaHei UI", 12)
            self.font_title = ("Microsoft YaHei UI", 12, "bold")
            self.font_big = ("Microsoft YaHei UI", 14, "bold")
            self.font_log = ("Consolas", 11)
        else:
            self.font_base = ("SF Pro Display", 14)
            self.font_title = ("SF Pro Display", 14, "bold")
            self.font_big = ("SF Pro Display", 16, "bold")
            self.font_log = ("Menlo", 12)
        
        style.configure(".", font=self.font_base)
        style.configure("TLabelframe.Label", font=self.font_title)
        style.configure("TButton", font=self.font_base, padding=10)
        style.configure("Big.TButton", font=self.font_big, padding=(30, 16))
    
    def _init_variables(self):
        """初始化变量"""
        self.mode_var = tk.StringVar(value="PPT")
        self.scope_var = tk.StringVar(value="CURRENT")
        
        self.detect_mode_var = tk.StringVar(value="smart")
        self.threshold_var = tk.IntVar(value=250)
        self.sensitivity_var = tk.IntVar(value=15)
        self.padding_var = tk.StringVar(value="2")
        
        self.output_format_var = tk.StringVar(value="PDF")
        self.output_dir_var = tk.StringVar(value="<与源文件同目录>")
        self.dpi_var = tk.StringVar(value="300")
        
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
    
    def _create_ui(self):
        """创建界面"""
        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True, padx=20, pady=20)
        self._main_frame = notebook

        main = ttk.Frame(notebook, padding=20)
        transparent_tab = ttk.Frame(notebook, padding=20)
        notebook.add(main, text="白边裁剪")
        notebook.add(transparent_tab, text="背景透明")
        
        # 1. 来源
        self._create_source_section(main)
        
        # 2. 范围
        self._create_scope_section(main)
        
        # 3. 检测参数
        self._create_detect_section(main)
        
        # 4. 输出设置
        self._create_output_section(main)
        
        # 5. 操作
        self._create_action_section(main)
        
        # 6. 日志
        self._create_log_section(main)

        self._create_transparency_tab(transparent_tab)
        
        # 延迟检查 pymupdf 可用性（不阻塞窗口显示）
        self.after(0, self._deferred_pymupdf_check)
    
    def _create_source_section(self, parent):
        """来源选择"""
        frame = ttk.LabelFrame(parent, text="来源", padding=(20, 15))
        frame.pack(fill="x", pady=(0, 12))
        
        # 模式选择行
        mode_row = ttk.Frame(frame)
        mode_row.pack(fill="x")
        
        ttk.Radiobutton(mode_row, text="PowerPoint", variable=self.mode_var, 
                       value="PPT", command=self._on_mode_change).pack(side="left", padx=(0, 30))
        ttk.Radiobutton(mode_row, text="本地文件 (PDF/图片/SVG)", variable=self.mode_var,
                       value="FILE", command=self._on_mode_change).pack(side="left")
        
        # PPT 连接区
        self.ppt_frame = ttk.Frame(frame)
        ppt_row = ttk.Frame(self.ppt_frame)
        ppt_row.pack(fill="x", pady=(12, 0))
        
        self.btn_connect = ttk.Button(ppt_row, text="检测连接", command=self._check_ppt)
        self.btn_connect.pack(side="left")
        self.ppt_status = ttk.Label(ppt_row, text="", foreground="gray")
        self.ppt_status.pack(side="left", padx=(15, 0))
        
        # 文件选择区
        self.file_frame = ttk.Frame(frame)
        
        file_row = ttk.Frame(self.file_frame)
        file_row.pack(fill="x", pady=(12, 0))
        
        ttk.Label(file_row, text="文件:").pack(side="left")
        ttk.Entry(file_row, textvariable=self.file_path_var, font=self.font_base).pack(
            side="left", fill="x", expand=True, padx=(10, 10))
        ttk.Button(file_row, text="浏览", command=self._select_files).pack(side="left")
        
        self.file_info = ttk.Label(self.file_frame, text="支持 PDF、SVG、PNG、JPG、BMP、TIFF、WebP 等格式",
                                   foreground="gray")
        self.file_info.pack(anchor="w", pady=(8, 0))
    
    def _create_scope_section(self, parent):
        """处理范围"""
        frame = ttk.LabelFrame(parent, text="处理范围", padding=(20, 15))
        frame.pack(fill="x", pady=(0, 12))
        
        row = ttk.Frame(frame)
        row.pack(fill="x")
        
        self.rb_current = ttk.Radiobutton(row, text="仅当前页", variable=self.scope_var, 
                                          value="CURRENT", command=self._on_scope_change)
        self.rb_current.pack(side="left", padx=(0, 30))
        
        ttk.Radiobutton(row, text="全部页面", variable=self.scope_var, 
                       value="ALL", command=self._on_scope_change).pack(side="left")
        
        # 页码输入（PDF 专用）
        self.page_frame = ttk.Frame(frame)
        page_row = ttk.Frame(self.page_frame)
        page_row.pack(fill="x", pady=(12, 0))
        
        ttk.Label(page_row, text="指定页码:").pack(side="left")
        self.page_spin = ttk.Spinbox(page_row, from_=1, to=9999, textvariable=self.page_num_var, 
                                     width=8, font=self.font_base)
        self.page_spin.pack(side="left", padx=(10, 0))
        self.page_total = ttk.Label(page_row, text="", foreground="gray")
        self.page_total.pack(side="left", padx=(10, 0))
    
    def _create_detect_section(self, parent):
        """检测参数"""
        frame = ttk.LabelFrame(parent, text="检测参数", padding=(20, 15))
        frame.pack(fill="x", pady=(0, 12))
        
        # 模式
        row1 = ttk.Frame(frame)
        row1.pack(fill="x", pady=(0, 12))
        
        ttk.Label(row1, text="模式:").pack(side="left")
        for text, val in [("智能", "smart"), ("简单", "simple"), ("边缘敏感", "edge")]:
            ttk.Radiobutton(row1, text=text, variable=self.detect_mode_var, 
                           value=val).pack(side="left", padx=(15, 0))
        
        # 阈值
        row2 = ttk.Frame(frame)
        row2.pack(fill="x", pady=(0, 12))
        
        ttk.Label(row2, text="白色阈值:", width=10).pack(side="left")
        ttk.Scale(row2, from_=200, to=255, variable=self.threshold_var,
                 orient="horizontal", length=160).pack(side="left")
        self.threshold_label = ttk.Label(row2, text="250", width=4, font=self.font_title)
        self.threshold_label.pack(side="left", padx=(10, 0))
        ttk.Label(row2, text="越高=只裁纯白", foreground="#888888").pack(side="left", padx=(15, 0))
        self.threshold_var.trace("w", lambda *_: self.threshold_label.config(
            text=str(self.threshold_var.get())))
        
        # 敏感度
        row3 = ttk.Frame(frame)
        row3.pack(fill="x", pady=(0, 12))
        
        ttk.Label(row3, text="敏感度:", width=10).pack(side="left")
        ttk.Scale(row3, from_=5, to=50, variable=self.sensitivity_var,
                 orient="horizontal", length=160).pack(side="left")
        self.sensitivity_label = ttk.Label(row3, text="15", width=4, font=self.font_title)
        self.sensitivity_label.pack(side="left", padx=(10, 0))
        ttk.Label(row3, text="越低=裁剪更多", foreground="#888888").pack(side="left", padx=(15, 0))
        self.sensitivity_var.trace("w", lambda *_: self.sensitivity_label.config(
            text=str(self.sensitivity_var.get())))
        
        # 留白
        row4 = ttk.Frame(frame)
        row4.pack(fill="x")
        
        ttk.Label(row4, text="边缘留白:", width=10).pack(side="left")
        ttk.Entry(row4, textvariable=self.padding_var, width=8, font=self.font_base).pack(side="left")
        ttk.Label(row4, text="px", foreground="gray").pack(side="left", padx=(5, 0))
    
    def _create_output_section(self, parent):
        """输出设置"""
        frame = ttk.LabelFrame(parent, text="输出设置", padding=(20, 15))
        frame.pack(fill="x", pady=(0, 12))
        
        # 格式 + DPI
        row1 = ttk.Frame(frame)
        row1.pack(fill="x", pady=(0, 12))
        
        ttk.Label(row1, text="格式:").pack(side="left")
        self.format_combo = ttk.Combobox(row1, textvariable=self.output_format_var,
                                         values=["PDF", "SVG", "PNG", "TIFF", "JPEG", "WebP"],
                                         state="readonly", width=8, font=self.font_base)
        self.format_combo.pack(side="left", padx=(10, 30))
        self.format_combo.bind("<<ComboboxSelected>>", lambda e: self._on_format_change())
        
        # DPI 区域（可隐藏）
        self.dpi_frame = ttk.Frame(row1)
        ttk.Label(self.dpi_frame, text="DPI:").pack(side="left")
        ttk.Entry(self.dpi_frame, textvariable=self.dpi_var, width=8, font=self.font_base).pack(side="left", padx=(10, 0))
        
        # 矢量提示（可隐藏）
        self.vector_hint = ttk.Label(row1, text="矢量格式，保持原始画质", foreground="gray")
        
        # 输出目录
        row2 = ttk.Frame(frame)
        row2.pack(fill="x")
        
        ttk.Label(row2, text="保存到:").pack(side="left")
        ttk.Entry(row2, textvariable=self.output_dir_var, font=self.font_base).pack(
            side="left", fill="x", expand=True, padx=(10, 10))
        ttk.Button(row2, text="选择", command=self._select_output_dir).pack(side="left")
    
    def _create_action_section(self, parent):
        """操作区"""
        frame = ttk.Frame(parent)
        frame.pack(fill="x", pady=(0, 12))
        
        # 状态
        self.status_label = ttk.Label(frame, textvariable=self.status_var, foreground="#0066cc")
        self.status_label.pack(fill="x", pady=(0, 8))
        
        # 进度条
        self.progress = ttk.Progressbar(frame, variable=self.progress_var, maximum=100)
        self.progress.pack(fill="x", pady=(0, 12))
        
        # 开始按钮
        self.btn_start = ttk.Button(frame, text="开始处理", style="Big.TButton",
                                   command=self._start_processing)
        self.btn_start.pack(fill="x", ipady=6)
    
    def _create_log_section(self, parent):
        """日志区"""
        frame = ttk.LabelFrame(parent, text="日志", padding=(10, 10))
        frame.pack(fill="both", expand=True)
        
        container = ttk.Frame(frame)
        container.pack(fill="both", expand=True)
        
        self.log_text = tk.Text(container, height=14, font=self.font_log,
                               bg="#fafafa", relief="flat", wrap="word",
                               padx=10, pady=8, spacing3=4)
        self.log_text.pack(side="left", fill="both", expand=True)
        
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=self.log_text.yview)
        scrollbar.pack(side="right", fill="y")
        self.log_text.config(yscrollcommand=scrollbar.set, state="disabled")
        
        self.log_text.tag_configure("TIME", foreground="#999999")
        self.log_text.tag_configure("INFO", foreground="#333333")
        self.log_text.tag_configure("SUCCESS", foreground="#28a745")
        self.log_text.tag_configure("WARNING", foreground="#f39c12")
        self.log_text.tag_configure("ERROR", foreground="#dc3545")

    def _create_transparency_tab(self, parent):
        """创建背景透明处理页"""
        self._create_transparency_source_section(parent)
        self._create_transparency_params_section(parent)
        self._create_transparency_output_section(parent)
        self._create_transparency_action_section(parent)
        self._create_transparency_log_section(parent)

        self.transparent_tolerance_var.trace("w", lambda *_: self.transparent_tolerance_label_var.set(
            str(self.transparent_tolerance_var.get())))
        self.transparent_feather_var.trace("w", lambda *_: self.transparent_feather_label_var.set(
            str(self.transparent_feather_var.get())))

    def _create_transparency_source_section(self, parent):
        frame = ttk.LabelFrame(parent, text="来源", padding=(20, 15))
        frame.pack(fill="x", pady=(0, 12))

        row = ttk.Frame(frame)
        row.pack(fill="x")

        ttk.Label(row, text="文件:").pack(side="left")
        ttk.Entry(row, textvariable=self.transparent_file_path_var, font=self.font_base).pack(
            side="left", fill="x", expand=True, padx=(10, 10))
        ttk.Button(row, text="浏览", command=self._select_transparent_files).pack(side="left")

        ttk.Label(frame, textvariable=self.transparent_info_var, foreground="gray").pack(anchor="w", pady=(8, 0))

    def _create_transparency_params_section(self, parent):
        frame = ttk.LabelFrame(parent, text="透明参数", padding=(20, 15))
        frame.pack(fill="x", pady=(0, 12))

        row1 = ttk.Frame(frame)
        row1.pack(fill="x", pady=(0, 12))
        ttk.Label(row1, text="目标颜色:", width=10).pack(side="left")
        for text, val in [("四角平均", "corners"), ("左上角", "top_left"), ("自定义", "custom")]:
            ttk.Radiobutton(row1, text=text, variable=self.transparent_color_mode_var,
                            value=val).pack(side="left", padx=(15, 0))
        ttk.Entry(row1, textvariable=self.transparent_custom_color_var, width=10,
                  font=self.font_base).pack(side="left", padx=(20, 0))

        row2 = ttk.Frame(frame)
        row2.pack(fill="x", pady=(0, 12))
        ttk.Label(row2, text="容差:", width=10).pack(side="left")
        ttk.Scale(row2, from_=0, to=100, variable=self.transparent_tolerance_var,
                  orient="horizontal", length=160).pack(side="left")
        ttk.Label(row2, textvariable=self.transparent_tolerance_label_var,
                  width=4, font=self.font_title).pack(side="left", padx=(10, 0))
        ttk.Label(row2, text="越高=去除更多相近颜色", foreground="#888888").pack(side="left", padx=(15, 0))

        row3 = ttk.Frame(frame)
        row3.pack(fill="x", pady=(0, 12))
        ttk.Label(row3, text="边缘羽化:", width=10).pack(side="left")
        ttk.Scale(row3, from_=0, to=8, variable=self.transparent_feather_var,
                  orient="horizontal", length=160).pack(side="left")
        ttk.Label(row3, textvariable=self.transparent_feather_label_var,
                  width=4, font=self.font_title).pack(side="left", padx=(10, 0))
        ttk.Label(row3, text="px", foreground="gray").pack(side="left", padx=(5, 0))

        row4 = ttk.Frame(frame)
        row4.pack(fill="x")
        ttk.Checkbutton(row4, text="只处理连到图片边缘的背景",
                        variable=self.transparent_edge_only_var).pack(side="left")

    def _create_transparency_output_section(self, parent):
        frame = ttk.LabelFrame(parent, text="输出设置", padding=(20, 15))
        frame.pack(fill="x", pady=(0, 12))

        row1 = ttk.Frame(frame)
        row1.pack(fill="x", pady=(0, 12))
        ttk.Label(row1, text="格式:").pack(side="left")
        ttk.Combobox(row1, textvariable=self.transparent_format_var,
                     values=["PNG", "WebP"], state="readonly",
                     width=8, font=self.font_base).pack(side="left", padx=(10, 0))

        row2 = ttk.Frame(frame)
        row2.pack(fill="x")
        ttk.Label(row2, text="保存到:").pack(side="left")
        ttk.Entry(row2, textvariable=self.transparent_output_dir_var, font=self.font_base).pack(
            side="left", fill="x", expand=True, padx=(10, 10))
        ttk.Button(row2, text="选择", command=self._select_transparent_output_dir).pack(side="left")

    def _create_transparency_action_section(self, parent):
        frame = ttk.Frame(parent)
        frame.pack(fill="x", pady=(0, 12))

        ttk.Label(frame, textvariable=self.transparent_status_var,
                  foreground="#0066cc").pack(fill="x", pady=(0, 8))
        ttk.Progressbar(frame, variable=self.transparent_progress_var,
                        maximum=100).pack(fill="x", pady=(0, 12))
        self.btn_transparent_start = ttk.Button(
            frame, text="开始处理透明背景", style="Big.TButton",
            command=self._start_transparency_processing)
        self.btn_transparent_start.pack(fill="x", ipady=6)

    def _create_transparency_log_section(self, parent):
        frame = ttk.LabelFrame(parent, text="日志", padding=(10, 10))
        frame.pack(fill="both", expand=True)

        container = ttk.Frame(frame)
        container.pack(fill="both", expand=True)

        self.transparent_log_text = tk.Text(container, height=14, font=self.font_log,
                                            bg="#fafafa", relief="flat", wrap="word",
                                            padx=10, pady=8, spacing3=4)
        self.transparent_log_text.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(container, orient="vertical", command=self.transparent_log_text.yview)
        scrollbar.pack(side="right", fill="y")
        self.transparent_log_text.config(yscrollcommand=scrollbar.set, state="disabled")

        self.transparent_log_text.tag_configure("TIME", foreground="#999999")
        self.transparent_log_text.tag_configure("INFO", foreground="#333333")
        self.transparent_log_text.tag_configure("SUCCESS", foreground="#28a745")
        self.transparent_log_text.tag_configure("WARNING", foreground="#f39c12")
        self.transparent_log_text.tag_configure("ERROR", foreground="#dc3545")
    
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
        """模式切换"""
        mode = self.mode_var.get()
        
        self.ppt_frame.pack_forget()
        self.file_frame.pack_forget()
        self.page_frame.pack_forget()
        
        if mode == "PPT":
            self.ppt_frame.pack(fill="x")
            self.rb_current.config(text="仅当前幻灯片")
        else:
            self.file_frame.pack(fill="x")
            self.rb_current.config(text="仅指定页码")
    
    def _on_scope_change(self):
        """范围切换"""
        scope = self.scope_var.get()
        
        if scope == "ALL":
            self.page_frame.pack_forget()
            self.page_spin.config(state="disabled")
        else:
            if self.mode_var.get() == "FILE" and self.source_files:
                from controllers import FileController
                first = self.source_files[0]
                if FileController.is_pdf(first):
                    self.page_frame.pack(fill="x")
            self.page_spin.config(state="normal")
    
    def _on_format_change(self):
        """输出格式切换"""
        fmt = self.output_format_var.get()
        
        if fmt in ("PDF", "SVG"):
            # 矢量格式，隐藏 DPI，显示提示
            self.dpi_frame.pack_forget()
            self.vector_hint.pack(side="left")
        else:
            # 位图格式，显示 DPI，隐藏提示
            self.vector_hint.pack_forget()
            self.dpi_frame.pack(side="left")
    
    def _check_ppt(self):
        """检测 PPT 连接"""
        from controllers import get_ppt_controller
        
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
                        self.page_frame.pack(fill="x")
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
        ts = datetime.datetime.now().strftime("%H:%M:%S")
        
        self.log_text.config(state="normal")
        self.log_text.insert("end", f"[{ts}] ", "TIME")
        self.log_text.insert("end", f"{message}\n", level)
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    def transparent_log(self, message, level="INFO"):
        """记录透明背景日志"""
        import datetime
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
            try:
                padding = float(self.padding_var.get())
            except:
                padding = 2.0
            
            try:
                dpi = int(self.dpi_var.get())
            except:
                dpi = 300
            
            config = {
                "scope": self.scope_var.get(),
                "output_format": self.output_format_var.get(),
                "output_dir": self.output_dir_var.get() if self.output_dir_var.get() != "<与源文件同目录>" else None,
                "detect_mode": self.detect_mode_var.get(),
                "threshold": self.threshold_var.get(),
                "sensitivity": self.sensitivity_var.get(),
                "padding": padding,
                "dpi": dpi,
                "page_num": self.page_num_var.get(),
                "source_files": self.source_files,
            }
            
            mode = self.mode_var.get()
            result = None
            
            if mode == "PPT":
                result = processor.process_ppt(config)
            else:
                if not self.source_files:
                    self.after(0, lambda: self.log("请先选择文件", "ERROR"))
                else:
                    result = processor.process_file(config)
            
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
