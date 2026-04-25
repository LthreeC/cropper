# -*- coding: utf-8 -*-
"""
白边检测算法模块 v0.2
提供三种检测模式：smart(智能), simple(简单), edge(边缘敏感)
"""


class WhiteBorderDetector:
    """白边检测器"""
    
    # 缓存预处理结果
    _cache = {}
    
    def __init__(self, threshold=250, sensitivity=15, mode="smart"):
        self.threshold = threshold
        self.sensitivity = sensitivity
        self.mode = mode
    
    def detect(self, pil_img):
        """检测内容边界框"""
        import numpy as np
        from PIL import Image
        
        # 大图缩放检测
        scale = 1.0
        max_size = 1200
        
        if pil_img.width > max_size or pil_img.height > max_size:
            scale = max_size / max(pil_img.width, pil_img.height)
            small_img = pil_img.resize(
                (int(pil_img.width * scale), int(pil_img.height * scale)),
                Image.NEAREST  # 最快的插值
            )
        else:
            small_img = pil_img
        
        img = self._preprocess(small_img)
        
        # 使用 uint8 更快
        img_array = np.asarray(img, dtype=np.uint8)
        
        if self.mode == "simple":
            mask = self._detect_simple(img_array)
        elif self.mode == "edge":
            mask = self._detect_edge(img_array)
        else:
            mask = self._detect_smart(img_array)
        
        bbox = self._get_bbox_from_mask(mask)
        
        # 还原坐标
        if bbox and scale != 1.0:
            x0, y0, x1, y1 = bbox
            bbox = (
                int(x0 / scale),
                int(y0 / scale),
                min(pil_img.width, int(x1 / scale) + 1),
                min(pil_img.height, int(y1 / scale) + 1)
            )
        
        return bbox
    
    def _preprocess(self, pil_img):
        """预处理：统一转 RGB"""
        from PIL import Image
        
        if pil_img.mode == 'RGB':
            return pil_img
        
        if pil_img.mode in ('RGBA', 'LA') or \
           (pil_img.mode == 'P' and 'transparency' in pil_img.info):
            bg = Image.new("RGB", pil_img.size, (255, 255, 255))
            if pil_img.mode != 'RGBA':
                pil_img = pil_img.convert('RGBA')
            bg.paste(pil_img, mask=pil_img.split()[3])
            return bg
        return pil_img.convert("RGB")
    
    def _detect_simple(self, img_array):
        """简单模式：仅亮度检测"""
        # 使用整数运算更快
        gray = img_array.mean(axis=2)
        return gray < self.threshold
    
    def _detect_edge(self, img_array):
        """边缘敏感模式"""
        import numpy as np
        
        gray = img_array.mean(axis=2).astype(np.int16)
        
        # 梯度检测 (使用切片避免创建新数组)
        grad_x = np.abs(gray[:, 1:] - gray[:, :-1])
        grad_y = np.abs(gray[1:, :] - gray[:-1, :])
        
        # 填充回原尺寸
        gx = np.zeros_like(gray)
        gy = np.zeros_like(gray)
        gx[:, 1:] = grad_x
        gy[1:, :] = grad_y
        
        gradient = np.maximum(gx, gy)
        edge_threshold = max(3, int(self.sensitivity * 0.3))
        
        edge_mask = gradient > edge_threshold
        content_mask = gray < self.threshold
        combined = edge_mask | content_mask
        
        try:
            from scipy import ndimage
            combined = ndimage.binary_dilation(combined, iterations=1)
        except ImportError:
            pass
        
        return combined
    
    def _detect_smart(self, img_array):
        """智能模式：综合检测"""
        import numpy as np
        
        # 一次性计算灰度
        gray = img_array.mean(axis=2).astype(np.int16)
        
        # 1. 亮度检测
        brightness_mask = gray < self.threshold
        
        # 2. 颜色变化检测
        r = img_array[:, :, 0].astype(np.int16)
        g = img_array[:, :, 1].astype(np.int16)
        b = img_array[:, :, 2].astype(np.int16)
        
        # 使用 abs 差值代替多次 maximum
        rg = np.abs(r - g)
        gb = np.abs(g - b)
        color_var = np.maximum(rg, gb)
        color_threshold = max(3, int(self.sensitivity * 0.5))
        color_mask = color_var > color_threshold
        
        # 3. 边缘检测（简化）
        grad_x = np.abs(gray[:, 1:] - gray[:, :-1])
        grad_y = np.abs(gray[1:, :] - gray[:-1, :])
        
        gx = np.zeros_like(gray)
        gy = np.zeros_like(gray)
        gx[:, 1:] = grad_x
        gy[1:, :] = grad_y
        
        gradient = np.maximum(gx, gy)
        edge_threshold = max(3, int(self.sensitivity * 0.4))
        edge_mask = gradient > edge_threshold
        
        # 合并
        combined = brightness_mask | color_mask | edge_mask
        
        # 形态学处理
        try:
            from scipy import ndimage
            combined = ndimage.binary_dilation(combined, iterations=1)
        except ImportError:
            pass
        
        return combined
    
    def _get_bbox_from_mask(self, mask):
        """从 mask 获取边界框"""
        import numpy as np
        
        rows = np.any(mask, axis=1)
        cols = np.any(mask, axis=0)
        
        if not rows.any() or not cols.any():
            return None
        
        y_indices = np.where(rows)[0]
        x_indices = np.where(cols)[0]
        
        x0, x1 = int(x_indices[0]), int(x_indices[-1] + 1)
        y0, y1 = int(y_indices[0]), int(y_indices[-1] + 1)
        
        # 内容占比检查
        content_ratio = mask.sum() / mask.size
        if content_ratio > 0.95:
            return None
        
        return (x0, y0, x1, y1)


def get_bbox(pil_img, threshold=250, sensitivity=15, mode="smart"):
    """便捷函数"""
    detector = WhiteBorderDetector(threshold, sensitivity, mode)
    return detector.detect(pil_img)
