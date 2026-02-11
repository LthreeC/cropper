# -*- coding: utf-8 -*-
"""
核心处理模块
"""

import os
import re
import tempfile
from PIL import Image

try:
    import fitz
    HAS_FITZ = True
except ImportError:
    fitz = None
    HAS_FITZ = False

from detector import get_bbox
from controllers import get_ppt_controller, FileController, CURRENT_OS


class CropProcessor:
    """裁剪处理器"""
    
    def __init__(self, callback=None):
        self.callback = callback or (lambda *a, **k: None)
    
    def log(self, msg, level="INFO"):
        self.callback(None, (level, msg), None)
    
    def set_status(self, msg):
        self.callback(msg, None, None)
    
    def set_progress(self, val):
        self.callback(None, None, val)
    
    def process_ppt(self, config):
        """处理 PPT"""
        controller = get_ppt_controller()
        
        if not controller.check_connection():
            self.log("未找到运行中的 PPT", "ERROR")
            return None
        
        current_index, ppt_name, src_path = controller.get_info()
        total_pages = controller.get_slide_count() if config["scope"] == "ALL" else 1
        
        self.log(f"已连接: {ppt_name} (第 {current_index} 页)")
        
        output_dir = config.get("output_dir") or os.path.dirname(src_path) or os.path.expanduser("~/Desktop")
        base_name = self._safe_name(os.path.splitext(ppt_name)[0])
        out_fmt = config["output_format"]
        
        if out_fmt in ["PDF", "SVG"]:
            return self._process_ppt_vector(controller, config, current_index, output_dir, base_name, total_pages)
        else:
            return self._process_ppt_raster(controller, config, current_index, output_dir, base_name, total_pages)
    
    def _process_ppt_vector(self, controller, config, current_index, output_dir, base_name, total_pages):
        """PPT 矢量输出"""
        scope = config["scope"]
        out_fmt = config["output_format"]
        padding = config["padding"]
        threshold = config["threshold"]
        sensitivity = config["sensitivity"]
        detect_mode = config["detect_mode"]
        
        temp_pdf = os.path.join(tempfile.gettempdir(), f"temp_{os.getpid()}.pdf")
        
        try:
            self.set_status("导出 PDF...")
            controller.export_temp_pdf(temp_pdf, scope=scope, index=current_index)
            
            if scope == "ALL":
                final_path = os.path.join(output_dir, f"{base_name}_Cropped.{out_fmt.lower()}")
            else:
                final_path = os.path.join(output_dir, f"{base_name}_p{current_index}.{out_fmt.lower()}")
            
            self._process_vector_crop(temp_pdf, out_fmt, final_path, padding, threshold, sensitivity, detect_mode, base_name)
            return final_path
        finally:
            if os.path.exists(temp_pdf):
                os.remove(temp_pdf)
    
    def _process_ppt_raster(self, controller, config, current_index, output_dir, base_name, total_pages):
        """PPT 光栅输出"""
        scope = config["scope"]
        out_fmt = config["output_format"]
        padding = config["padding"]
        threshold = config["threshold"]
        sensitivity = config["sensitivity"]
        detect_mode = config["detect_mode"]
        dpi = config.get("dpi", 300)
        
        ppt_w, ppt_h = controller.get_page_setup()
        target_width = int((ppt_w / 72.0) * dpi)
        
        pages = range(1, total_pages + 1) if scope == "ALL" else [current_index]
        
        if scope == "ALL":
            final_path = os.path.join(output_dir, f"{base_name}_Images")
            os.makedirs(final_path, exist_ok=True)
        else:
            final_path = os.path.join(output_dir, f"{base_name}_p{current_index}.{out_fmt.lower()}")
        
        for i, idx in enumerate(pages):
            self.set_status(f"处理第 {idx} 页...")
            self.set_progress((i + 1) / len(pages) * 100)
            
            temp_img = os.path.join(tempfile.gettempdir(), f"temp_{os.getpid()}_{idx}.png")
            
            try:
                controller.export_single_image(temp_img, target_width, index=idx)
                
                if os.path.exists(temp_img):
                    img = Image.open(temp_img)
                    cropped = self._crop_image(img, padding, threshold, sensitivity, detect_mode)
                    
                    if cropped:
                        if scope == "ALL":
                            save_path = os.path.join(final_path, f"{base_name}_p{idx:03d}.{out_fmt.lower()}")
                        else:
                            save_path = final_path
                        self._save_image(cropped, save_path, out_fmt, dpi)
                        self.log(f"第 {idx} 页完成")
                    else:
                        self.log(f"第 {idx} 页: 未检测到内容", "WARNING")
            finally:
                if os.path.exists(temp_img):
                    os.remove(temp_img)
        
        self.log(f"完成: {final_path}", "SUCCESS")
        return final_path
    
    def process_file(self, config):
        """处理本地文件 (PDF 或图片)"""
        paths = config.get("source_files", [])
        if not paths:
            self.log("未选择文件", "ERROR")
            return None
        
        first_path = paths[0]
        
        if FileController.is_pdf(first_path):
            return self._process_pdf(config, first_path)
        else:
            return self._process_images(config, paths)
    
    def _process_pdf(self, config, pdf_path):
        """处理 PDF"""
        if not HAS_FITZ:
            self.log("未安装 pymupdf", "ERROR")
            return None
        
        controller = FileController()
        total_pages = controller.get_pdf_page_count(pdf_path)
        
        self.log(f"加载 PDF: {os.path.basename(pdf_path)} ({total_pages} 页)")
        
        output_dir = config.get("output_dir") or os.path.dirname(pdf_path)
        base_name = self._safe_name(os.path.splitext(os.path.basename(pdf_path))[0])
        out_fmt = config["output_format"]
        scope = config["scope"]
        page_num = config.get("page_num", 1)
        padding = config["padding"]
        threshold = config["threshold"]
        sensitivity = config["sensitivity"]
        detect_mode = config["detect_mode"]
        dpi = config.get("dpi", 300)
        
        if out_fmt in ["PDF", "SVG"]:
            temp_pdf = os.path.join(tempfile.gettempdir(), f"temp_{os.getpid()}.pdf")
            try:
                doc = fitz.open(pdf_path)
                if scope != "ALL":
                    doc.select([page_num - 1])
                doc.save(temp_pdf)
                doc.close()
                
                if scope == "ALL":
                    final_path = os.path.join(output_dir, f"{base_name}_Cropped.{out_fmt.lower()}")
                else:
                    final_path = os.path.join(output_dir, f"{base_name}_p{page_num}.{out_fmt.lower()}")
                
                self._process_vector_crop(temp_pdf, out_fmt, final_path, padding, threshold, sensitivity, detect_mode, base_name)
                return final_path
            finally:
                if os.path.exists(temp_pdf):
                    os.remove(temp_pdf)
        else:
            pages = range(1, total_pages + 1) if scope == "ALL" else [page_num]
            
            if scope == "ALL":
                final_path = os.path.join(output_dir, f"{base_name}_Images")
                os.makedirs(final_path, exist_ok=True)
            else:
                final_path = os.path.join(output_dir, f"{base_name}_p{page_num}.{out_fmt.lower()}")
            
            for i, idx in enumerate(pages):
                self.set_status(f"处理第 {idx} 页...")
                self.set_progress((i + 1) / len(pages) * 100)
                
                img = controller.render_pdf_page(pdf_path, idx - 1, dpi=dpi)
                cropped = self._crop_image(img, padding, threshold, sensitivity, detect_mode)
                
                if cropped:
                    if scope == "ALL":
                        save_path = os.path.join(final_path, f"{base_name}_p{idx:03d}.{out_fmt.lower()}")
                    else:
                        save_path = final_path
                    self._save_image(cropped, save_path, out_fmt, dpi)
                    self.log(f"第 {idx} 页完成")
                else:
                    self.log(f"第 {idx} 页: 未检测到内容", "WARNING")
            
            self.log(f"完成: {final_path}", "SUCCESS")
            return final_path
    
    def _process_images(self, config, paths):
        """处理图片"""
        output_dir = config.get("output_dir") or os.path.dirname(paths[0])
        out_fmt = config["output_format"]
        if out_fmt in ["PDF", "SVG"]:
            out_fmt = "PNG"
        
        padding = config["padding"]
        threshold = config["threshold"]
        sensitivity = config["sensitivity"]
        detect_mode = config["detect_mode"]
        
        results = []
        controller = FileController()
        
        for i, path in enumerate(paths):
            self.set_status(f"处理 {i+1}/{len(paths)}...")
            self.set_progress((i + 1) / len(paths) * 100)
            
            try:
                img = controller.load_image(path)
                original_dpi = img.info.get('dpi', (300, 300))
                if isinstance(original_dpi, tuple):
                    dpi = original_dpi[0]
                else:
                    dpi = original_dpi
                
                cropped = self._crop_image(img, padding, threshold, sensitivity, detect_mode)
                
                if cropped:
                    base_name = self._safe_name(os.path.splitext(os.path.basename(path))[0])
                    save_path = os.path.join(output_dir, f"{base_name}_cropped.{out_fmt.lower()}")
                    self._save_image(cropped, save_path, out_fmt, dpi)
                    results.append(save_path)
                    self.log(f"{os.path.basename(path)} 完成")
                else:
                    self.log(f"{os.path.basename(path)}: 未检测到内容", "WARNING")
            except Exception as e:
                self.log(f"{os.path.basename(path)} 失败: {e}", "ERROR")
        
        if results:
            self.log(f"完成 {len(results)} 个文件", "SUCCESS")
            return output_dir
        return None
    
    def _crop_image(self, img, padding, threshold, sensitivity, detect_mode):
        """裁剪图片"""
        bbox = get_bbox(img, threshold, sensitivity, detect_mode)
        if bbox is None:
            return None
        
        x0, y0, x1, y1 = bbox
        crop_box = (
            max(0, x0 - padding),
            max(0, y0 - padding),
            min(img.width, x1 + padding),
            min(img.height, y1 + padding)
        )
        return img.crop(crop_box)
    
    def _process_vector_crop(self, source_pdf, out_fmt, final_path, padding, threshold, sensitivity, detect_mode, base_name):
        """矢量裁剪"""
        doc = fitz.open(source_pdf)
        total = len(doc)
        
        for i, page in enumerate(doc):
            self.set_status(f"分析第 {i+1}/{total} 页...")
            self.set_progress((i + 1) / total * 100)
            
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            bbox = get_bbox(img, threshold, sensitivity, detect_mode)
            
            if bbox:
                scale = 2.0
                page.set_cropbox(fitz.Rect(
                    max(0, bbox[0] - padding * scale) / scale,
                    max(0, bbox[1] - padding * scale) / scale,
                    min(pix.width, bbox[2] + padding * scale) / scale,
                    min(pix.height, bbox[3] + padding * scale) / scale
                ))
        
        self.set_status("保存中...")
        
        if out_fmt == "PDF":
            doc.save(final_path, garbage=4, deflate=True)
            self.log(f"PDF 已保存: {final_path}", "SUCCESS")
        elif out_fmt == "SVG":
            svg_dir = os.path.dirname(final_path)
            for i, page in enumerate(doc):
                svg_path = os.path.join(svg_dir, f"{base_name}_p{i+1}.svg")
                with open(svg_path, "w", encoding="utf-8") as f:
                    f.write(page.get_svg_image())
            self.log("SVG 已导出", "SUCCESS")
        
        doc.close()
    
    def _save_image(self, img, path, fmt, dpi):
        """保存图片"""
        kwargs = {"dpi": (dpi, dpi)}
        
        if fmt.upper() in ["JPEG", "JPG"]:
            if img.mode == "RGBA":
                img = img.convert("RGB")
            kwargs["quality"] = 95
        elif fmt.upper() == "TIFF":
            kwargs["compression"] = "tiff_lzw"
        elif fmt.upper() == "WEBP":
            kwargs["quality"] = 95
        
        img.save(path, **kwargs)
    
    def _safe_name(self, name):
        return re.sub(r'[\\/*?:"<>|]', "", name) or "Export"
