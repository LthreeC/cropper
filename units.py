# -*- coding: utf-8 -*-
"""PDF、PostScript 与 PowerPoint 共用的尺寸单位换算。"""


# 规范定义：1 point = 1/72 inch。它与显示器 DPI 和系统缩放无关。
POINTS_PER_INCH = 72.0


def points_to_pixels(points, dpi):
    """按目标 DPI 将 point 长度换算为像素长度。"""
    return float(points) * float(dpi) / POINTS_PER_INCH


def pixels_to_points(pixels, dpi):
    """按目标 DPI 将像素长度换算为 point 长度。"""
    dpi = float(dpi)
    if dpi <= 0:
        raise ValueError("DPI 必须大于 0")
    return float(pixels) * POINTS_PER_INCH / dpi
