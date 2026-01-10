#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
icon.py - 应用程序图标
圣斯德望十字架 (Cross of Saint Stephen) 图标
"""

import os
import sys
import tkinter as tk
from typing import Optional

# 全局缓存，避免重复加载
_icon_photo: Optional[tk.PhotoImage] = None
_icon_path: Optional[str] = None


def _get_base_dir() -> str:
    """获取项目根目录"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def get_icon_path() -> Optional[str]:
    """获取图标文件路径"""
    global _icon_path
    
    if _icon_path is not None:
        return _icon_path
    
    base_dir = _get_base_dir()
    
    # 只使用 cross_of_saint_stephen.png
    icon_candidates = [
        'cross_of_saint_stephen.png',
        'Cross_of_saint_stephen.png',
    ]
    
    for icon_name in icon_candidates:
        icon_path = os.path.join(base_dir, icon_name)
        if os.path.exists(icon_path):
            _icon_path = icon_path
            return _icon_path
    
    return None


def set_window_icon(root: tk.Tk) -> None:
    """
    设置主窗口图标为圣斯德望十字架
    """
    global _icon_photo
    
    try:
        icon_path = get_icon_path()
        if icon_path:
            # 确保窗口已经映射
            root.update_idletasks()
            # 创建新的PhotoImage并保存引用
            _icon_photo = tk.PhotoImage(file=icon_path)
            root.iconphoto(True, _icon_photo)
            # 同时保存到root对象防止被回收
            root._icon_photo = _icon_photo
        else:
            print("警告: 未找到图标文件 cross_of_saint_stephen.png")
    except Exception as e:
        print(f"设置窗口图标失败: {e}")


def set_toplevel_icon(toplevel: tk.Toplevel) -> None:
    """
    设置Toplevel窗口图标为圣斯德望十字架
    """
    try:
        icon_path = get_icon_path()
        if icon_path:
            # 为每个Toplevel创建新的PhotoImage
            # 注意：必须保存引用以防止被垃圾回收
            photo = tk.PhotoImage(file=icon_path)
            toplevel.iconphoto(False, photo)
            # 将photo保存到窗口对象上以防止被回收
            toplevel._icon_photo = photo
    except Exception as e:
        print(f"设置Toplevel图标失败: {e}")
