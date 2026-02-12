# -*- coding: utf-8 -*-
"""
数据源控制器模块
支持 PPT (Windows/Mac)、PDF 文件、图片文件
"""

import os
import platform
import subprocess

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


class WindowsPPTController(BaseController):
    """Windows PPT 控制器"""
    
    def __init__(self):
        self.app = None
    
    def _connect(self):
        try:
            import win32com.client
        except ImportError:
            return False
        try:
            self.app = win32com.client.GetActiveObject("PowerPoint.Application")
            return True
        except:
            try:
                self.app = win32com.client.GetActiveObject("Kwpp.Application")
                return True
            except:
                return False
    
    def check_connection(self):
        return self._connect()
    
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
        slide.Export(save_path, "PNG", int(width_px))
    
    def export_temp_pdf(self, save_path, scope="CURRENT", index=1):
        presentation = self.app.ActivePresentation
        try:
            if scope == "ALL":
                presentation.ExportAsFixedFormat(Path=save_path, FixedFormatType=2)
            else:
                presentation.PrintOptions.Ranges.ClearAll()
                presentation.PrintOptions.Ranges.Add(Start=index, End=index)
                presentation.ExportAsFixedFormat(
                    Path=save_path, FixedFormatType=2, RangeType=4,
                    PrintRange=presentation.PrintOptions.Ranges.Item(1)
                )
        except Exception as e:
            try:
                presentation.SaveAs(save_path, 32)
            except Exception as e2:
                raise Exception(f"导出PDF失败: {e} | {e2}")


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
    
    IMAGE_FORMATS = {'.png', '.jpg', '.jpeg', '.bmp', '.tiff', '.tif', '.webp', '.gif'}
    
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
            self.file_type = "pdf" if ext == '.pdf' else "image"
    
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
        count = len(doc)
        doc.close()
        return count
    
    def render_pdf_page(self, path, page_index, dpi=300):
        """渲染 PDF 页面"""
        from PIL import Image
        pymupdf = _lazy_import_pymupdf()
        doc = pymupdf.open(path)
        page = doc[page_index]
        zoom = dpi / 72.0
        pix = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), alpha=False)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        doc.close()
        return img
    
    def load_image(self, path):
        """加载图片"""
        from PIL import Image
        return Image.open(path)
    
    @classmethod
    def is_image(cls, path):
        ext = os.path.splitext(path)[1].lower()
        return ext in cls.IMAGE_FORMATS
    
    @classmethod
    def is_pdf(cls, path):
        return os.path.splitext(path)[1].lower() == '.pdf'


def get_ppt_controller():
    """获取 PPT 控制器"""
    if CURRENT_OS == "Windows":
        return WindowsPPTController()
    else:
        return MacPPTController()
