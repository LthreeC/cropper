# -*- coding: utf-8 -*-
"""
核心处理模块
"""

import os
import re
import tempfile
import time
from math import ceil

from units import pixels_to_points, points_to_pixels


VECTOR_PADDING_DPI = 300
VECTOR_REFINE_MAX_SCALE = 8.0
VECTOR_REFINE_MAX_LONG_EDGE = 8192
VECTOR_VERIFY_DPI = 300
VECTOR_VERIFY_ITERATIONS = 3
# 半个 576 DPI 检测像素再加少量取整余量，消除不同 PDF 渲染器的末行差异。
VECTOR_STRICT_INSET_POINTS = 0.075
# 位图重采样最多会跨一个 300 DPI 像素，仅在严格零留白时补偿。
RASTER_STRICT_INSET_POINTS = 0.24
# 含位图的边缘在不同 PDF/SVG 渲染器间最多相差约两个 300 DPI 像素。
RASTER_PADDING_COMPENSATION_POINTS = 0.48


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

    def _cleanup_temp_file(self, path):
        """清理可能仍被 PowerPoint 短暂占用的临时文件"""
        for attempt in range(5):
            if not os.path.exists(path):
                return
            try:
                os.remove(path)
                return
            except PermissionError:
                if attempt < 4:
                    time.sleep(0.1 * (attempt + 1))

        self.log(f"临时文件仍被占用，稍后可手动删除: {path}", "WARNING")
    
    def process_ppt(self, config):
        """处理 PPT"""
        from controllers import get_ppt_controller
        controller = get_ppt_controller()

        try:
            if not controller.check_connection():
                self.log("未找到运行中的 PPT", "ERROR")
                return None

            current_index, ppt_name, src_path = controller.get_info()
            total_pages = controller.get_slide_count() if config["scope"] == "ALL" else 1

            self.log(f"已连接: {ppt_name} (第 {current_index} 页)")

            source_dir = ""
            if src_path and not re.match(r"^https?://", src_path, re.IGNORECASE):
                source_dir = os.path.dirname(src_path)
            output_dir = config.get("output_dir") or source_dir or os.path.expanduser("~/Desktop")
            if not config.get("output_dir") and not source_dir and src_path:
                self.log("PPT 位于云端，输出目录已改为本机桌面", "WARNING")

            base_name = self._safe_name(os.path.splitext(ppt_name)[0])
            out_fmt = config["output_format"]

            if out_fmt in ["PDF", "SVG"]:
                return self._process_ppt_vector(controller, config, current_index, output_dir, base_name, total_pages)
            else:
                return self._process_ppt_raster(controller, config, current_index, output_dir, base_name, total_pages)
        finally:
            controller.close()
    
    def _process_ppt_vector(self, controller, config, current_index, output_dir, base_name, total_pages):
        """PPT 矢量输出"""
        scope = config["scope"]
        out_fmt = config["output_format"]
        padding = config["padding"]
        threshold = config["threshold"]
        sensitivity = config["sensitivity"]
        detect_mode = config["detect_mode"]
        try:
            pdf_image_dpi = max(1, int(config.get("pdf_image_dpi", 300)))
        except (TypeError, ValueError):
            pdf_image_dpi = 300
        
        temp_pdf = self._make_temp_path(".pdf")
        temp_pptx = self._make_temp_path(".pptx")
        vector_doc = None
        
        try:
            self.set_status("导出 PDF...")
            controller.export_temp_pdf(
                temp_pdf,
                scope=scope,
                index=current_index,
            )

            vector_source = temp_pdf
            try:
                import pymupdf
                from ppt_image_restore import (
                    restore_pptx_images_in_document,
                )

                self.set_status("恢复 PPT 原始图片...")
                self._cleanup_temp_file(temp_pptx)
                controller.export_source_copy(temp_pptx)
                vector_doc = pymupdf.open(temp_pdf)
                restored = restore_pptx_images_in_document(
                    vector_doc,
                    temp_pptx,
                    max_image_dpi=pdf_image_dpi,
                )
                vector_source = vector_doc
                if restored:
                    self.log(
                        f"已优化 {len(restored)} 个 PDF 位图，"
                        f"最高约 {pdf_image_dpi} DPI"
                    )
            except Exception as e:
                if vector_doc is not None:
                    vector_doc.close()
                    vector_doc = None
                self.log(
                    "未能恢复 PPT 原始位图，将保留 PowerPoint "
                    f"导出结果: {e}",
                    "WARNING",
                )
            
            if scope == "ALL":
                final_path = os.path.join(
                    output_dir,
                    f"{base_name}_Cropped.{out_fmt.lower()}",
                )
            else:
                final_path = os.path.join(
                    output_dir,
                    f"{base_name}_p{current_index}.{out_fmt.lower()}",
                )
            
            return self._process_vector_crop(
                vector_source,
                out_fmt,
                final_path,
                padding,
                threshold,
                sensitivity,
                detect_mode,
                base_name,
                pdf_garbage=4,
            )
        finally:
            if vector_doc is not None:
                vector_doc.close()
            self._cleanup_temp_file(temp_pdf)
            self._cleanup_temp_file(temp_pptx)
    
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
        target_width = int(points_to_pixels(ppt_w, dpi))
        
        pages = range(1, total_pages + 1) if scope == "ALL" else [current_index]
        
        if scope == "ALL":
            final_path = os.path.join(output_dir, f"{base_name}_Images")
            os.makedirs(final_path, exist_ok=True)
        else:
            final_path = os.path.join(output_dir, f"{base_name}_p{current_index}.{out_fmt.lower()}")
        
        saved_paths = []
        for i, idx in enumerate(pages):
            self.set_status(f"处理第 {idx} 页...")
            self.set_progress((i + 1) / len(pages) * 100)
            
            temp_img = os.path.join(tempfile.gettempdir(), f"temp_{os.getpid()}_{idx}.png")
            
            try:
                controller.export_single_image(temp_img, target_width, index=idx)
                
                if os.path.exists(temp_img):
                    from PIL import Image
                    img = Image.open(temp_img)
                    cropped = self._crop_image(img, padding, threshold, sensitivity, detect_mode)
                    
                    if cropped:
                        if scope == "ALL":
                            save_path = os.path.join(final_path, f"{base_name}_p{idx:03d}.{out_fmt.lower()}")
                        else:
                            save_path = final_path
                        self._save_image(cropped, save_path, out_fmt, dpi)
                        saved_paths.append(save_path)
                        self.log(f"第 {idx} 页完成")
                    else:
                        self.log(f"第 {idx} 页: 未检测到内容", "WARNING")
            finally:
                if os.path.exists(temp_img):
                    os.remove(temp_img)
        
        if saved_paths:
            self.log(f"完成: {final_path}", "SUCCESS")
            return final_path
        self.log("未检测到可输出内容", "WARNING")
        return None
    
    def process_file(self, config):
        """处理本地文件 (PDF 或图片)"""
        from controllers import FileController
        
        paths = config.get("source_files", [])
        if not paths:
            self.log("未选择文件", "ERROR")
            return None
        
        pdf_paths = [path for path in paths if FileController.is_pdf(path)]
        if pdf_paths:
            if len(paths) > 1:
                self.log("PDF 文件需单独处理，请一次只选择一个 PDF", "ERROR")
                return None
            return self._process_pdf(config, pdf_paths[0])

        return self._process_images(config, paths)
    
    def _process_pdf(self, config, pdf_path):
        """处理 PDF"""
        try:
            import pymupdf
        except ImportError:
            self.log("未安装 pymupdf，请运行: pip install pymupdf", "ERROR")
            return None
        
        from controllers import FileController
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
            if scope == "ALL":
                final_path = os.path.join(output_dir, f"{base_name}_Cropped.{out_fmt.lower()}")
                page_indices = None
            else:
                final_path = os.path.join(output_dir, f"{base_name}_p{page_num}.{out_fmt.lower()}")
                page_indices = [page_num - 1]

            return self._process_vector_crop(
                pdf_path, out_fmt, final_path, padding, threshold, sensitivity,
                detect_mode, base_name, page_indices=page_indices
            )
        else:
            pages = range(1, total_pages + 1) if scope == "ALL" else [page_num]
            
            if scope == "ALL":
                final_path = os.path.join(output_dir, f"{base_name}_Images")
                os.makedirs(final_path, exist_ok=True)
            else:
                final_path = os.path.join(output_dir, f"{base_name}_p{page_num}.{out_fmt.lower()}")
            
            saved_paths = []
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
                    saved_paths.append(save_path)
                    self.log(f"第 {idx} 页完成")
                else:
                    self.log(f"第 {idx} 页: 未检测到内容", "WARNING")
            
            if saved_paths:
                self.log(f"完成: {final_path}", "SUCCESS")
                return final_path
            self.log("未检测到可输出内容", "WARNING")
            return None
    
    def _process_images(self, config, paths):
        """处理图片"""
        from controllers import FileController
        
        output_dir = config.get("output_dir") or os.path.dirname(paths[0])
        requested_fmt = config["output_format"]
        
        padding = config["padding"]
        threshold = config["threshold"]
        sensitivity = config["sensitivity"]
        detect_mode = config["detect_mode"]
        output_dpi = config.get("dpi", 300)
        
        results = []
        controller = FileController()
        output_bases = self._output_base_names(paths)
        
        for i, (path, output_base) in enumerate(zip(paths, output_bases)):
            self.set_status(f"处理 {i+1}/{len(paths)}...")
            self.set_progress((i + 1) / len(paths) * 100)
            
            try:
                if FileController.is_svg(path):
                    result = self._process_svg_file(
                        path, output_dir, requested_fmt, padding, threshold,
                        sensitivity, detect_mode, output_dpi,
                        base_name=output_base,
                    )
                    if result:
                        results.append(result)
                    continue

                out_fmt = requested_fmt
                img = controller.load_image(path)
                original_dpi = img.info.get('dpi', (300, 300))
                if isinstance(original_dpi, tuple):
                    dpi = original_dpi[0]
                else:
                    dpi = original_dpi
                
                cropped = self._crop_image(img, padding, threshold, sensitivity, detect_mode)
                
                if cropped:
                    save_path = os.path.join(
                        output_dir,
                        f"{output_base}_cropped.{out_fmt.lower()}",
                    )
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
    
    def _process_svg_file(
        self, path, output_dir, out_fmt, padding, threshold, sensitivity,
        detect_mode, dpi, base_name=None,
    ):
        """处理 SVG 输入"""
        if base_name is None:
            base_name = self._safe_name(
                os.path.splitext(os.path.basename(path))[0]
            )

        if out_fmt in ["PDF", "SVG"]:
            temp_pdf = self._make_temp_path(".pdf")
            try:
                self._convert_to_pdf(path, temp_pdf)
                save_path = os.path.join(output_dir, f"{base_name}_cropped.{out_fmt.lower()}")
                result = self._process_vector_crop(
                    temp_pdf, out_fmt, save_path, padding, threshold, sensitivity, detect_mode, base_name
                )
                self.log(f"{os.path.basename(path)} 完成")
                return result
            finally:
                if os.path.exists(temp_pdf):
                    os.remove(temp_pdf)

        from controllers import FileController
        controller = FileController()
        img = controller.render_document_page(path, 0, dpi=dpi)
        cropped = self._crop_image(img, padding, threshold, sensitivity, detect_mode)

        if cropped:
            save_path = os.path.join(output_dir, f"{base_name}_cropped.{out_fmt.lower()}")
            self._save_image(cropped, save_path, out_fmt, dpi)
            self.log(f"{os.path.basename(path)} 完成")
            return save_path

        self.log(f"{os.path.basename(path)}: 未检测到内容", "WARNING")
        return None

    def process_transparency(self, config):
        """将图片背景处理为透明"""
        from controllers import FileController
        from transparency import make_transparent, rgb_to_hex

        paths = config.get("source_files", [])
        if not paths:
            self.log("未选择文件", "ERROR")
            return None

        output_dir = config.get("output_dir") or os.path.dirname(paths[0])
        os.makedirs(output_dir, exist_ok=True)

        out_fmt = config.get("output_format", "PNG").upper()
        if out_fmt not in ("PNG", "WEBP"):
            out_fmt = "PNG"

        controller = FileController()
        dpi = config.get("dpi", 300)
        results = []
        output_bases = self._output_base_names(paths)

        for i, (path, output_base) in enumerate(zip(paths, output_bases)):
            self.set_status(f"透明背景 {i+1}/{len(paths)}...")
            self.set_progress((i + 1) / len(paths) * 100)

            try:
                if FileController.is_pdf(path):
                    self.log(f"{os.path.basename(path)}: 透明背景暂不支持 PDF", "WARNING")
                    continue

                if FileController.is_svg(path):
                    img = controller.render_document_page(path, 0, dpi=dpi)
                    image_dpi = dpi
                else:
                    img = controller.load_image(path)
                    original_dpi = img.info.get("dpi", (dpi, dpi))
                    image_dpi = original_dpi[0] if isinstance(original_dpi, tuple) else original_dpi

                transparent, color, ratio = make_transparent(
                    img,
                    color_mode=config.get("color_mode", "corners"),
                    custom_color=config.get("custom_color", "#FFFFFF"),
                    tolerance=config.get("tolerance", 18),
                    edge_only=config.get("edge_only", True),
                    feather=config.get("feather", 1),
                )

                save_path = os.path.join(
                    output_dir,
                    f"{output_base}_transparent.{out_fmt.lower()}",
                )
                self._save_transparent_image(transparent, save_path, out_fmt, image_dpi)
                results.append(save_path)

                if ratio > 0:
                    self.log(f"{os.path.basename(path)} 完成，已移除 {ratio:.1%}，颜色 {rgb_to_hex(color)}")
                else:
                    self.log(f"{os.path.basename(path)}: 未找到匹配背景，已保存原图透明通道", "WARNING")
            except Exception as e:
                self.log(f"{os.path.basename(path)} 失败: {e}", "ERROR")

        if results:
            self.log(f"透明背景完成 {len(results)} 个文件", "SUCCESS")
            return output_dir
        return None

    def _crop_image(self, img, padding, threshold, sensitivity, detect_mode):
        """裁剪图片"""
        from detector import get_bbox
        padding = max(0.0, float(padding))
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

    @staticmethod
    def _detect_page_clip(
        page,
        scale,
        threshold,
        sensitivity,
        detect_mode,
        clip=None,
        allow_full_content=False,
    ):
        """按原生分辨率检测页面或窄边带，返回页面坐标和局部像素框。"""
        import pymupdf
        from PIL import Image
        from detector import get_bbox

        pixmap = page.get_pixmap(
            matrix=pymupdf.Matrix(scale, scale),
            clip=clip,
            alpha=False,
        )
        if pixmap.width <= 0 or pixmap.height <= 0:
            return None

        image = Image.frombytes(
            "RGB",
            (pixmap.width, pixmap.height),
            pixmap.samples,
        )
        bbox = get_bbox(
            image,
            threshold,
            sensitivity,
            detect_mode,
            max_detection_size=None,
            allow_full_content=allow_full_content,
        )
        if bbox is None:
            return None

        content_rect = pymupdf.Rect(
            (pixmap.x + bbox[0]) / scale,
            (pixmap.y + bbox[1]) / scale,
            (pixmap.x + bbox[2]) / scale,
            (pixmap.y + bbox[3]) / scale,
        )
        return content_rect, bbox, (pixmap.width, pixmap.height)

    @staticmethod
    def _has_excess_vector_border(bbox, size, padding_points, scale):
        """判断当前页面是否仍有超过目标留白的完整边缘像素。"""
        margins = (
            bbox[0],
            bbox[1],
            size[0] - bbox[2],
            size[1] - bbox[3],
        )
        allowed_margin = ceil(padding_points * scale)
        return any(margin > allowed_margin for margin in margins)

    @staticmethod
    def _has_raster_at_page_edge(page, tolerance=1.0):
        """判断裁后页面边缘是否由位图覆盖，避免影响仅含内部图片的页面。"""
        import pymupdf

        page_rect = page.rect
        for image in page.get_image_info():
            image_rect = pymupdf.Rect(image["bbox"])
            if not image_rect.intersects(page_rect):
                continue
            if (
                image_rect.x0 <= page_rect.x0 + tolerance
                or image_rect.y0 <= page_rect.y0 + tolerance
                or image_rect.x1 >= page_rect.x1 - tolerance
                or image_rect.y1 >= page_rect.y1 - tolerance
            ):
                return True
        return False

    def _refine_vector_bbox(
        self,
        page,
        rough_rect,
        coarse_scale,
        threshold,
        sensitivity,
        detect_mode,
    ):
        """只高分辨率渲染四条窄边带，精修粗略内容框。"""
        import pymupdf

        page_rect = page.rect
        long_edge = max(page_rect.width, page_rect.height)
        if long_edge <= 0:
            return rough_rect

        refine_scale = min(
            VECTOR_REFINE_MAX_SCALE,
            VECTOR_REFINE_MAX_LONG_EDGE / long_edge,
        )
        refine_scale = max(coarse_scale, refine_scale)
        band = max(2.0, 4.0 / coarse_scale)

        clips = {
            "left": pymupdf.Rect(
                max(page_rect.x0, rough_rect.x0 - band),
                page_rect.y0,
                min(page_rect.x1, rough_rect.x0 + band),
                page_rect.y1,
            ),
            "right": pymupdf.Rect(
                max(page_rect.x0, rough_rect.x1 - band),
                page_rect.y0,
                min(page_rect.x1, rough_rect.x1 + band),
                page_rect.y1,
            ),
            "top": pymupdf.Rect(
                page_rect.x0,
                max(page_rect.y0, rough_rect.y0 - band),
                page_rect.x1,
                min(page_rect.y1, rough_rect.y0 + band),
            ),
            "bottom": pymupdf.Rect(
                page_rect.x0,
                max(page_rect.y0, rough_rect.y1 - band),
                page_rect.x1,
                min(page_rect.y1, rough_rect.y1 + band),
            ),
        }

        edges = {
            "left": rough_rect.x0,
            "right": rough_rect.x1,
            "top": rough_rect.y0,
            "bottom": rough_rect.y1,
        }
        for name, clip in clips.items():
            if clip.is_empty:
                continue
            detected = self._detect_page_clip(
                page,
                refine_scale,
                threshold,
                sensitivity,
                detect_mode,
                clip=clip,
                allow_full_content=True,
            )
            if detected is None:
                continue
            content_rect = detected[0]
            if name == "left":
                edges[name] = content_rect.x0
            elif name == "right":
                edges[name] = content_rect.x1
            elif name == "top":
                edges[name] = content_rect.y0
            else:
                edges[name] = content_rect.y1

        refined = pymupdf.Rect(
            edges["left"],
            edges["top"],
            edges["right"],
            edges["bottom"],
        )
        refined &= page_rect
        return rough_rect if refined.is_empty else refined

    @staticmethod
    def _visible_rect_to_cropbox(page, visible_rect, current_cropbox):
        """把旋转后的可视坐标转换为未旋转 CropBox 坐标。"""
        import pymupdf

        detected = visible_rect * page.derotation_matrix
        detected = pymupdf.Rect(
            detected.x0 + current_cropbox.x0,
            detected.y0 + current_cropbox.y0,
            detected.x1 + current_cropbox.x0,
            detected.y1 + current_cropbox.y0,
        )
        detected &= current_cropbox
        return detected

    def _tighten_vector_cropbox(
        self,
        page,
        threshold,
        sensitivity,
        detect_mode,
    ):
        """在 300 DPI 下迭代移除四边完整白色像素行或列。"""
        import pymupdf

        scale = VECTOR_VERIFY_DPI / 72.0
        band = max(2.0, 4.0 / scale)

        for _ in range(VECTOR_VERIFY_ITERATIONS):
            page_rect = page.rect
            if page_rect.is_empty:
                break

            clips = {
                "left": pymupdf.Rect(
                    page_rect.x0,
                    page_rect.y0,
                    min(page_rect.x1, page_rect.x0 + band),
                    page_rect.y1,
                ),
                "right": pymupdf.Rect(
                    max(page_rect.x0, page_rect.x1 - band),
                    page_rect.y0,
                    page_rect.x1,
                    page_rect.y1,
                ),
                "top": pymupdf.Rect(
                    page_rect.x0,
                    page_rect.y0,
                    page_rect.x1,
                    min(page_rect.y1, page_rect.y0 + band),
                ),
                "bottom": pymupdf.Rect(
                    page_rect.x0,
                    max(page_rect.y0, page_rect.y1 - band),
                    page_rect.x1,
                    page_rect.y1,
                ),
            }

            shifts = {name: 0.0 for name in clips}
            for name, clip in clips.items():
                detected = self._detect_page_clip(
                    page,
                    scale,
                    threshold,
                    sensitivity,
                    detect_mode,
                    clip=clip,
                    allow_full_content=True,
                )
                if detected is None:
                    shifts[name] = (
                        clip.width if name in ("left", "right")
                        else clip.height
                    )
                    continue

                _, bbox, size = detected
                if name == "left":
                    shifts[name] = bbox[0] / scale
                elif name == "right":
                    shifts[name] = (size[0] - bbox[2]) / scale
                elif name == "top":
                    shifts[name] = bbox[1] / scale
                else:
                    shifts[name] = (size[1] - bbox[3]) / scale

            if not any(value > 0 for value in shifts.values()):
                break

            visible_rect = pymupdf.Rect(
                page_rect.x0 + shifts["left"],
                page_rect.y0 + shifts["top"],
                page_rect.x1 - shifts["right"],
                page_rect.y1 - shifts["bottom"],
            )
            if visible_rect.is_empty:
                break

            current_cropbox = page.cropbox
            tightened = self._visible_rect_to_cropbox(
                page,
                visible_rect,
                current_cropbox,
            )
            if tightened.is_empty or all(
                abs(a - b) < 1e-6
                for a, b in zip(tightened, current_cropbox)
            ):
                break
            page.set_cropbox(tightened)

        return pymupdf.Rect(page.cropbox)

    @staticmethod
    def _cropbox_to_mediabox(cropbox, original_mediabox):
        """把 PyMuPDF 的顶部原点 CropBox 转为 PDF MediaBox 坐标。"""
        import pymupdf

        return pymupdf.Rect(
            cropbox.x0,
            original_mediabox.y1 - cropbox.y1,
            cropbox.x1,
            original_mediabox.y1 - cropbox.y0,
        )
    
    def _process_vector_crop(
        self, source_pdf, out_fmt, final_path, padding, threshold,
        sensitivity, detect_mode, base_name, page_indices=None,
        pdf_garbage=3,
    ):
        """矢量裁剪"""
        import pymupdf
        from detector import MAX_DETECTION_SIZE

        padding_points = pixels_to_points(
            max(0.0, float(padding)),
            VECTOR_PADDING_DPI,
        )
        
        owns_document = not isinstance(source_pdf, pymupdf.Document)
        doc = pymupdf.open(source_pdf) if owns_document else source_pdf
        try:
            if page_indices is not None:
                doc.select(page_indices)
            total = len(doc)
        
            for i, page in enumerate(doc):
                self.set_status(f"分析第 {i+1}/{total} 页...")
                self.set_progress((i + 1) / total * 100)

                original_cropbox = pymupdf.Rect(page.cropbox)
                original_mediabox = pymupdf.Rect(page.mediabox)
                page_max_size = max(page.rect.width, page.rect.height)
                coarse_scale = (
                    min(2.0, MAX_DETECTION_SIZE / page_max_size)
                    if page_max_size else 1.0
                )
                coarse = self._detect_page_clip(
                    page,
                    coarse_scale,
                    threshold,
                    sensitivity,
                    detect_mode,
                )

                if coarse:
                    if not self._has_excess_vector_border(
                        coarse[1],
                        coarse[2],
                        padding_points,
                        coarse_scale,
                    ):
                        continue

                    refined = self._refine_vector_bbox(
                        page,
                        coarse[0],
                        coarse_scale,
                        threshold,
                        sensitivity,
                        detect_mode,
                    )
                    content_cropbox = self._visible_rect_to_cropbox(
                        page,
                        refined,
                        original_cropbox,
                    )
                    if content_cropbox.is_empty:
                        continue

                    page.set_cropbox(content_cropbox)
                    content_cropbox = self._tighten_vector_cropbox(
                        page,
                        threshold,
                        sensitivity,
                        detect_mode,
                    )
                    has_raster_edge = self._has_raster_at_page_edge(page)
                    if padding_points == 0:
                        strict_inset = (
                            RASTER_STRICT_INSET_POINTS
                            if has_raster_edge
                            else VECTOR_STRICT_INSET_POINTS
                        )
                        strict_cropbox = pymupdf.Rect(
                            content_cropbox.x0 + strict_inset,
                            content_cropbox.y0 + strict_inset,
                            content_cropbox.x1 - strict_inset,
                            content_cropbox.y1 - strict_inset,
                        )
                        if not strict_cropbox.is_empty:
                            content_cropbox = strict_cropbox
                    effective_padding = padding_points
                    if padding_points > 0 and has_raster_edge:
                        effective_padding += RASTER_PADDING_COMPENSATION_POINTS
                    final_cropbox = pymupdf.Rect(
                        content_cropbox.x0 - effective_padding,
                        content_cropbox.y0 - effective_padding,
                        content_cropbox.x1 + effective_padding,
                        content_cropbox.y1 + effective_padding,
                    )
                    final_cropbox &= original_cropbox
                    if final_cropbox.is_empty:
                        continue

                    if out_fmt == "PDF":
                        hard_mediabox = self._cropbox_to_mediabox(
                            final_cropbox,
                            original_mediabox,
                        )
                        page.set_mediabox(hard_mediabox)
                    else:
                        page.set_cropbox(final_cropbox)

            self.set_status("保存中...")

            if out_fmt == "PDF":
                doc.save(
                    final_path,
                    garbage=pdf_garbage,
                    deflate=True,
                )
                self.log(f"PDF 已保存: {final_path}", "SUCCESS")
                result_path = final_path
            elif out_fmt == "SVG":
                if total == 1:
                    svg_paths = [final_path]
                    result_path = final_path
                else:
                    svg_dir = os.path.splitext(final_path)[0]
                    os.makedirs(svg_dir, exist_ok=True)
                    svg_paths = [os.path.join(svg_dir, f"{base_name}_p{i+1:03d}.svg") for i in range(total)]
                    result_path = svg_dir
                for i, page in enumerate(doc):
                    svg_path = svg_paths[i]
                    with open(svg_path, "w", encoding="utf-8") as f:
                        f.write(page.get_svg_image())
                self.log(f"SVG 已导出: {result_path}", "SUCCESS")
            else:
                raise ValueError(f"不支持的矢量输出格式: {out_fmt}")
        
            return result_path
        finally:
            if owns_document:
                doc.close()
    
    def _save_image(self, img, path, fmt, dpi):
        """保存图片"""
        fmt_upper = fmt.upper()
        if fmt_upper == "PDF":
            self._save_raster_as_pdf(img, path, dpi)
            return
        if fmt_upper == "SVG":
            self._save_raster_as_svg(img, path, dpi)
            return

        kwargs = {"dpi": (dpi, dpi)}
        
        if fmt_upper in ["JPEG", "JPG"]:
            if img.mode not in ("RGB", "L", "CMYK"):
                if "A" in img.getbands() or "transparency" in img.info:
                    from PIL import Image

                    rgba = img.convert("RGBA")
                    background = Image.new("RGB", rgba.size, "white")
                    background.paste(rgba, mask=rgba.getchannel("A"))
                    img = background
                else:
                    img = img.convert("RGB")
            kwargs["quality"] = 95
        elif fmt_upper == "PNG" and not (
            img.mode in ("1", "L", "LA", "P", "RGB", "RGBA", "I")
            or img.mode.startswith("I;16")
        ):
            img = img.convert("RGB")
        elif fmt_upper == "TIFF":
            kwargs["compression"] = "tiff_lzw"
        elif fmt_upper == "WEBP":
            kwargs["quality"] = 95
        
        img.save(path, **kwargs)

    def _save_transparent_image(self, img, path, fmt, dpi):
        """保存带透明通道的图片"""
        if img.mode != "RGBA":
            img = img.convert("RGBA")

        if fmt.upper() == "WEBP":
            img.save(path, "WEBP", quality=95, lossless=True, method=6)
        else:
            img.save(path, "PNG", dpi=(dpi, dpi), optimize=True)
    
    def _save_raster_as_pdf(self, img, path, dpi):
        """将裁剪后的原始像素无重采样嵌入 PDF。"""
        import io

        import pymupdf

        try:
            width_pt = pixels_to_points(img.width, dpi)
            height_pt = pixels_to_points(img.height, dpi)
        except (TypeError, ValueError):
            width_pt = pixels_to_points(img.width, 300)
            height_pt = pixels_to_points(img.height, 300)

        if img.mode not in ("RGB", "RGBA"):
            img = img.convert("RGBA")

        buffer = io.BytesIO()
        img.save(buffer, "PNG", compress_level=1)

        doc = pymupdf.open()
        try:
            page = doc.new_page(width=width_pt, height=height_pt)
            page.insert_image(page.rect, stream=buffer.getvalue())
            doc.save(path, garbage=3, deflate=True)
        finally:
            doc.close()

    def _save_raster_as_svg(self, img, path, dpi):
        """将位图裁剪结果嵌入 SVG 输出"""
        import base64
        import io

        if img.mode not in ("RGB", "RGBA"):
            img = img.convert("RGBA")

        buf = io.BytesIO()
        img.save(buf, format="PNG", dpi=(dpi, dpi))
        data = base64.b64encode(buf.getvalue()).decode("ascii")
        width, height = img.size
        svg = (
            f'<svg xmlns="http://www.w3.org/2000/svg" '
            f'width="{width}" height="{height}" viewBox="0 0 {width} {height}">\n'
            f'  <image width="{width}" height="{height}" '
            f'href="data:image/png;base64,{data}"/>\n'
            f'</svg>\n'
        )
        with open(path, "w", encoding="utf-8") as f:
            f.write(svg)

    def _convert_to_pdf(self, source_path, target_path):
        """将 SVG 等矢量文档转换成临时 PDF，供统一裁剪流程使用"""
        import pymupdf

        doc = pymupdf.open(source_path)
        try:
            pdf_data = doc.convert_to_pdf()
        finally:
            doc.close()

        with open(target_path, "wb") as f:
            f.write(pdf_data)

    def _make_temp_path(self, suffix):
        fd, path = tempfile.mkstemp(prefix=f"cropper_{os.getpid()}_", suffix=suffix)
        os.close(fd)
        return path

    def _output_base_names(self, paths):
        """为同名不同格式的批处理输入生成稳定且不冲突的输出名。"""
        bases = [
            self._safe_name(os.path.splitext(os.path.basename(path))[0])
            for path in paths
        ]
        counts = {}
        for base in bases:
            key = base.casefold()
            counts[key] = counts.get(key, 0) + 1

        used = set()
        results = []
        for path, base in zip(paths, bases):
            if counts[base.casefold()] > 1:
                extension = os.path.splitext(path)[1].lower().lstrip(".")
                candidate_root = self._safe_name(
                    f"{base}_{extension or 'file'}"
                )
            else:
                candidate_root = base

            candidate = candidate_root
            serial = 2
            while candidate.casefold() in used:
                candidate = f"{candidate_root}_{serial}"
                serial += 1
            used.add(candidate.casefold())
            results.append(candidate)
        return results

    def _safe_name(self, name):
        return re.sub(r'[\\/*?:"<>|]', "", name) or "Export"
