# -*- coding: utf-8 -*-
"""
数据源控制器模块
支持 PPT (Windows/Mac)、PDF 文件、图片文件
"""

import os
import platform
import shutil
import subprocess
import tempfile

from units import POINTS_PER_INCH

CURRENT_OS = platform.system()


def _lazy_import_pymupdf():
    """懒加载 pymupdf"""
    try:
        import pymupdf
        return pymupdf
    except ImportError:
        return None


def has_pymupdf():
    """检查 pymupdf 是否可用"""
    return _lazy_import_pymupdf() is not None


def _pdf_slide_position(
    pdf_page_count,
    total_slide_count,
    requested_index,
    source_path=None,
    source_is_current=False,
):
    """把 PowerPoint 的真实页码映射到导出 PDF 的可见页位置。"""
    requested_slide = int(requested_index) - 1
    total_slide_count = int(total_slide_count)
    if requested_slide < 0 or requested_slide >= total_slide_count:
        raise ValueError(
            f"幻灯片页码超出范围: {requested_index}"
            f"（共 {total_slide_count} 页）"
        )

    if pdf_page_count == total_slide_count:
        exported_indices = tuple(range(total_slide_count))
    else:
        if (
            not source_path
            or source_path == "Unsaved"
            or str(source_path).lower().startswith(("http://", "https://"))
            or not os.path.isfile(source_path)
        ):
            raise RuntimeError(
                "PDF 未包含全部幻灯片，且没有可访问的本地 PPTX，"
                "无法可靠映射隐藏页"
            )
        if not source_is_current:
            raise RuntimeError(
                "演示文稿存在未保存更改，无法用磁盘 PPTX "
                "可靠映射隐藏页；请先保存后重试"
            )
        try:
            from ppt_image_restore import pptx_visible_slide_indices

            exported_indices = pptx_visible_slide_indices(source_path)
        except Exception as exc:
            raise RuntimeError(
                "无法从本地 PPTX 可靠映射导出 PDF 页序"
            ) from exc
        if len(exported_indices) != pdf_page_count:
            raise RuntimeError(
                "PowerPoint 导出 PDF 的页数与 PPTX 可见页数不一致，"
                "已停止以避免导出错误幻灯片"
            )

    try:
        return exported_indices.index(requested_slide)
    except ValueError as exc:
        raise ValueError(
            f"第 {requested_index} 页是隐藏幻灯片，"
            "PowerPoint PDF 导出中不包含该页"
        ) from exc


class BaseController:
    """控制器基类"""
    
    def check_connection(self):
        raise NotImplementedError
    
    def get_info(self):
        raise NotImplementedError
    
    def get_slide_count(self):
        return 1
    
    def get_page_setup(self):
        return 720, 540

    def close(self):
        pass


class WindowsPPTController(BaseController):
    """Windows PPT 控制器"""
    
    def __init__(self):
        self.app = None
        self.app_kind = None
        self._pythoncom = None
    
    def _connect(self):
        try:
            import pythoncom
            import win32com.client
        except ImportError as e:
            raise RuntimeError(
                "未安装 pywin32，无法连接 PowerPoint；请使用项目 venv_clean 启动"
            ) from e
        if self._pythoncom is None:
            pythoncom.CoInitialize()
            self._pythoncom = pythoncom
        try:
            self.app = win32com.client.GetActiveObject("PowerPoint.Application")
            self.app_kind = "PowerPoint"
            return True
        except:
            try:
                self.app = win32com.client.GetActiveObject("Kwpp.Application")
                self.app_kind = "WPS"
                return True
            except:
                return False
    
    def check_connection(self):
        return self._connect()

    def close(self):
        self.app = None
        if self._pythoncom is not None:
            self._pythoncom.CoUninitialize()
            self._pythoncom = None

    def _active_slide_index(self):
        window = self.app.ActiveWindow
        try:
            return int(window.View.Slide.SlideIndex)
        except Exception:
            try:
                return int(window.Selection.SlideRange(1).SlideIndex)
            except Exception as exc:
                raise RuntimeError(
                    "无法从当前 PowerPoint 视图确定活动幻灯片"
                ) from exc
    
    def get_info(self):
        if not self.app or self.app.Windows.Count == 0:
            raise Exception("未检测到活动的 PPT 窗口")
        window = self.app.ActiveWindow
        presentation = window.Presentation
        index = self._active_slide_index()
        return index, presentation.Name, presentation.FullName
    
    def get_slide_count(self):
        return self.app.ActivePresentation.Slides.Count
    
    def get_page_setup(self):
        ps = self.app.ActivePresentation.PageSetup
        return ps.SlideWidth, ps.SlideHeight
    
    def export_single_image(self, save_path, width_px, index=1):
        slide = self.app.ActivePresentation.Slides(index)
        ps = self.app.ActivePresentation.PageSetup
        height_px = round(width_px * ps.SlideHeight / ps.SlideWidth)
        slide.Export(save_path, "PNG", int(width_px), int(height_px))
    
    def export_temp_pdf(self, save_path, scope="CURRENT", index=1):
        presentation = self.app.ActivePresentation
        try:
            if scope == "ALL":
                presentation.ExportAsFixedFormat(
                    Path=save_path, FixedFormatType=2, Intent=2,
                    PrintRange=None, RangeType=1
                )
            else:
                if (
                    self.app_kind in ("PowerPoint", "WPS")
                    and self._active_slide_index() != int(index)
                ):
                    raise RuntimeError(
                        "活动幻灯片已在导出前发生变化，请重试"
                    )
                presentation.ExportAsFixedFormat(
                    Path=save_path, FixedFormatType=2, Intent=2,
                    PrintRange=None, RangeType=3
                )
        except Exception as e:
            if self.app_kind != "WPS":
                raise Exception(f"PowerPoint 高质量导出 PDF 失败: {e}") from e
            fallback_pdf = None
            source_path = getattr(presentation, "FullName", None)
            source_is_current = getattr(
                presentation,
                "Saved",
                None,
            ) in (True, -1)
            try:
                export_path = save_path
                if scope != "ALL":
                    fd, fallback_pdf = tempfile.mkstemp(suffix=".pdf")
                    os.close(fd)
                    os.remove(fallback_pdf)
                    export_path = fallback_pdf
                presentation.SaveAs(export_path, 32)

                if scope != "ALL":
                    pymupdf = _lazy_import_pymupdf()
                    if not pymupdf:
                        raise ImportError(
                            "未安装 pymupdf，无法提取当前幻灯片"
                        )
                    with pymupdf.open(fallback_pdf) as doc:
                        page_position = _pdf_slide_position(
                            len(doc),
                            presentation.Slides.Count,
                            index,
                            source_path,
                            source_is_current,
                        )
                        doc.select([page_position])
                        if os.path.isfile(save_path):
                            os.remove(save_path)
                        doc.save(save_path, garbage=4, deflate=True)
            except Exception as e2:
                raise Exception(f"导出PDF失败: {e} | {e2}") from e2
            finally:
                if fallback_pdf and os.path.exists(fallback_pdf):
                    os.remove(fallback_pdf)

    def export_source_copy(self, save_path):
        """保存临时 PPTX 副本，供恢复 PDF 中的原始位图。"""
        self.app.ActivePresentation.SaveCopyAs(save_path, 24)


class MacPPTController(BaseController):
    """Mac PPT 控制器"""

    @staticmethod
    def _applescript_string(value):
        return str(value).replace("\\", "\\\\").replace('"', '\\"')

    @staticmethod
    def _remove_empty_placeholder(path):
        if os.path.isfile(path) and os.path.getsize(path) == 0:
            os.remove(path)
    
    def _run_applescript(self, script):
        try:
            result = subprocess.run(
                ['osascript', '-e', script],
                capture_output=True, text=True, check=True
            )
            return result.stdout.strip()
        except subprocess.CalledProcessError as e:
            raise Exception(f"AppleScript Error: {e.stderr}")
    
    def check_connection(self):
        script = 'tell application "System Events" to (name of processes) contains "Microsoft PowerPoint"'
        return self._run_applescript(script) == "true"
    
    def get_info(self):
        script = '''
        tell application "Microsoft PowerPoint"
            if (count of windows) is 0 then error "No Active Window"
            set pptName to name of active presentation
            set pptPath to full name of active presentation
            try
                set pptPath to POSIX path of (pptPath as alias)
            end try
            try
                set slideIndex to slide index of slide range of selection of active window
            on error
                error "Unable to determine the active slide"
            end try
            return (slideIndex as string) & (character id 31) & pptName & ¬
                (character id 31) & pptPath
        end tell
        '''
        res = self._run_applescript(script)
        parts = res.split("\x1f")
        path = parts[2] if len(parts) > 2 else "Unsaved"
        return int(parts[0]), parts[1], path
    
    def get_slide_count(self):
        script = 'tell application "Microsoft PowerPoint" to count of slides of active presentation'
        return int(self._run_applescript(script))
    
    def get_page_setup(self):
        script = '''
        tell application "Microsoft PowerPoint"
            set w to width of page setup of active presentation
            set h to height of page setup of active presentation
            return w & "|" & h
        end tell
        '''
        res = self._run_applescript(script).split("|")
        return float(res[0]), float(res[1])

    def _source_presentation_is_saved(self):
        script = (
            'tell application "Microsoft PowerPoint" to '
            'return saved of active presentation'
        )
        return self._run_applescript(script).strip().lower() == "true"
    
    def export_single_image(self, save_path, width_px, index=1):
        """通过 PDF 按目标像素宽度渲染，避免 PowerPoint 忽略 PNG 尺寸。"""
        pymupdf = _lazy_import_pymupdf()
        if not pymupdf:
            raise ImportError("未安装 pymupdf，无法按指定 DPI 导出幻灯片")

        fd, temp_pdf = tempfile.mkstemp(suffix=".pdf")
        os.close(fd)
        try:
            self.export_temp_pdf(temp_pdf, scope="CURRENT", index=index)
            with pymupdf.open(temp_pdf) as doc:
                page = doc[0]
                scale = max(1, int(width_px)) / page.rect.width
                pixmap = page.get_pixmap(
                    matrix=pymupdf.Matrix(scale, scale),
                    alpha=False,
                )
                pixmap.save(save_path)
        finally:
            if os.path.exists(temp_pdf):
                os.remove(temp_pdf)
    
    def export_temp_pdf(self, save_path, scope="CURRENT", index=1):
        export_path = save_path
        temp_pdf = None
        source_path = None
        source_is_current = False
        if scope != "ALL":
            try:
                _, _, source_path = self.get_info()
                source_is_current = self._source_presentation_is_saved()
            except Exception:
                # 仅在 PDF 排除了隐藏页、确实需要磁盘页序时才报错。
                source_path = None
                source_is_current = False
            fd, temp_pdf = tempfile.mkstemp(suffix=".pdf")
            os.close(fd)
            export_path = temp_pdf

        escaped_path = self._applescript_string(os.path.abspath(export_path))
        self._remove_empty_placeholder(export_path)
        script = f'''
        tell application "Microsoft PowerPoint"
            set outPath to POSIX file "{escaped_path}" as string
            save active presentation in outPath as save as PDF
        end tell
        '''
        try:
            self._run_applescript(script)
            if scope == "ALL":
                return

            pymupdf = _lazy_import_pymupdf()
            if not pymupdf:
                raise ImportError("未安装 pymupdf，无法提取当前幻灯片")
            with pymupdf.open(temp_pdf) as doc:
                total_slide_count = self.get_slide_count()
                page_position = _pdf_slide_position(
                    len(doc),
                    total_slide_count,
                    index,
                    source_path,
                    source_is_current,
                )
                doc.select([page_position])
                self._remove_empty_placeholder(save_path)
                doc.save(save_path, garbage=4, deflate=True)
        finally:
            if temp_pdf and os.path.exists(temp_pdf):
                os.remove(temp_pdf)

    def export_source_copy(self, save_path):
        """复制本地 PPTX，供恢复 PDF 中的原始位图。"""
        _, _, source_path = self.get_info()
        if (
            not source_path
            or source_path == "Unsaved"
            or source_path.lower().startswith(("http://", "https://"))
            or not os.path.isfile(source_path)
        ):
            raise RuntimeError("当前演示文稿没有可访问的本地 PPTX 源文件")
        shutil.copy2(source_path, save_path)


class FileController(BaseController):
    """本地文件控制器 (PDF + 图片)"""
    
    RASTER_FORMATS = {'.png', '.jpg', '.jpeg', '.bmp', '.tiff', '.tif', '.webp', '.gif'}
    VECTOR_FORMATS = {'.svg'}
    IMAGE_FORMATS = RASTER_FORMATS | VECTOR_FORMATS
    
    def __init__(self):
        self.file_paths = []
        self.file_type = None
    
    def check_connection(self):
        return True
    
    def set_files(self, paths):
        """设置文件列表"""
        if isinstance(paths, str):
            paths = [paths]
        
        self.file_paths = []
        for p in paths:
            if os.path.exists(p):
                ext = os.path.splitext(p)[1].lower()
                if ext == '.pdf' or ext in self.IMAGE_FORMATS:
                    self.file_paths.append(p)
        
        if self.file_paths:
            ext = os.path.splitext(self.file_paths[0])[1].lower()
            self.file_type = "pdf" if ext == '.pdf' else "svg" if ext == '.svg' else "image"
    
    def get_file_count(self):
        return len(self.file_paths)
    
    def get_pdf_page_count(self, path=None):
        """获取 PDF 页数"""
        pymupdf = _lazy_import_pymupdf()
        if not pymupdf:
            return 0
        path = path or (self.file_paths[0] if self.file_paths else None)
        if not path:
            return 0
        doc = pymupdf.open(path)
        try:
            count = len(doc)
        finally:
            doc.close()
        return count

    def render_document_page(self, path, page_index=0, dpi=300):
        """渲染 PDF/SVG 页面"""
        from PIL import Image
        pymupdf = _lazy_import_pymupdf()
        if not pymupdf:
            raise ImportError("未安装 pymupdf，无法处理 PDF/SVG")
        doc = pymupdf.open(path)
        try:
            page = doc[page_index]
            zoom = dpi / POINTS_PER_INCH
            pix = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), alpha=False)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        finally:
            doc.close()
        return img

    def render_pdf_page(self, path, page_index, dpi=300):
        """渲染 PDF 页面"""
        return self.render_document_page(path, page_index, dpi)

    def load_image(self, path):
        """加载图片"""
        if self.is_svg(path):
            return self.render_document_page(path, 0)
        from PIL import Image, ImageOps

        with Image.open(path) as image:
            result = ImageOps.exif_transpose(image)
            result.load()
        return result
    
    @classmethod
    def is_image(cls, path):
        ext = os.path.splitext(path)[1].lower()
        return ext in cls.IMAGE_FORMATS
    
    @classmethod
    def is_pdf(cls, path):
        return os.path.splitext(path)[1].lower() == '.pdf'

    @classmethod
    def is_svg(cls, path):
        return os.path.splitext(path)[1].lower() == '.svg'


def get_ppt_controller():
    """获取 PPT 控制器"""
    if CURRENT_OS == "Windows":
        return WindowsPPTController()
    else:
        return MacPPTController()
