# -*- coding: utf-8 -*-
"""
数据源控制器模块
支持 PPT (Windows/Mac)、PDF 文件、图片文件
"""

import os
import platform
import subprocess

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
    
    def get_info(self):
        if not self.app or self.app.Windows.Count == 0:
            raise Exception("未检测到活动的 PPT 窗口")
        window = self.app.ActiveWindow
        presentation = window.Presentation
        try:
            index = window.View.Slide.SlideIndex
        except:
            try:
                index = window.Selection.SlideRange(1).SlideIndex
            except:
                index = 1
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
                presentation.ExportAsFixedFormat(
                    Path=save_path, FixedFormatType=2, Intent=2,
                    PrintRange=None, RangeType=3
                )
        except Exception as e:
            if self.app_kind != "WPS":
                raise Exception(f"PowerPoint 高质量导出 PDF 失败: {e}") from e
            try:
                presentation.SaveAs(save_path, 32)
            except Exception as e2:
                raise Exception(f"导出PDF失败: {e} | {e2}")

    def export_source_copy(self, save_path):
        """保存临时 PPTX 副本，供恢复 PDF 中的原始位图。"""
        self.app.ActivePresentation.SaveCopyAs(save_path, 24)


class MacPPTController(BaseController):
    """Mac PPT 控制器"""
    
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
                set slideIndex to slide index of slide range of selection of active window
            on error
                set slideIndex to 1
            end try
            return slideIndex & "|" & pptName & "|" & pptPath
        end tell
        '''
        res = self._run_applescript(script)
        parts = res.split("|")
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
    
    def export_single_image(self, save_path, width_px, index=1):
        script = f'''
        tell application "Microsoft PowerPoint"
            set outPath to POSIX file "{save_path}" as string
            export slide {index} of active presentation to outPath as PNG with properties {{width:{int(width_px)}}}
        end tell
        '''
        self._run_applescript(script)
    
    def export_temp_pdf(self, save_path, scope="CURRENT", index=1):
        script = f'''
        tell application "Microsoft PowerPoint"
            set outPath to POSIX file "{save_path}" as string
            save active presentation in outPath as save as PDF
        end tell
        '''
        self._run_applescript(script)


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
