# -*- coding: utf-8 -*-
"""
白边裁剪工具 v0.4
主入口 - 直接运行此文件

Usage:
    python main.py
"""

import platform
import tkinter as tk

VERSION = "0.4"
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
    
    # 根据屏幕工作区选择紧凑尺寸，避免小屏幕下内容被挡住
    screen_w = root.winfo_screenwidth()
    screen_h = root.winfo_screenheight()
    width = max(760, min(900, screen_w - 80))
    height = max(700, min(900, screen_h - 100))
    x = max(0, (screen_w - width) // 2)
    y = max(0, (screen_h - height) // 2 - 20)
    root.geometry(f"{width}x{height}+{x}+{y}")
    root.minsize(760, 700)
    
    # 创建应用
    app = CropperApp(root)
    app.pack(fill="both", expand=True)
    
    root.mainloop()


if __name__ == "__main__":
    main()
