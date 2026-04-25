# -*- coding: utf-8 -*-
"""
白边裁剪工具 v0.3
主入口 - 直接运行此文件

Usage:
    python main.py
"""

import platform
import tkinter as tk

VERSION = "0.3"
CURRENT_OS = platform.system()


def main():
    from ui import CropperApp

    # Windows DPI 适配（必须在创建窗口前）
    if CURRENT_OS == "Windows":
        try:
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(1)
        except:
            pass
    
    root = tk.Tk()
    root.title(f"白边裁剪工具 v{VERSION}")
    
    # 窗口大小和位置
    width, height = 780, 1300
    screen_w = root.winfo_screenwidth()
    screen_h = root.winfo_screenheight()
    x = (screen_w - width) // 2
    y = (screen_h - height) // 2 - 30
    root.geometry(f"{width}x{height}+{x}+{y}")
    root.minsize(700, 800)
    
    # 创建应用
    app = CropperApp(root)
    app.pack(fill="both", expand=True)
    
    root.mainloop()


if __name__ == "__main__":
    main()
