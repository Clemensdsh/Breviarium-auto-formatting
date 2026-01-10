#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
main.py - 程序入口点
支持命令行参数和调试模式
"""

import argparse
import tkinter as tk
import sys
import sys
import os

from psalter_generator import PsalterApp, init_logging, AppConfig

def resource_path(relative_path):
    """ 获取资源的绝对路径，适配开发环境和 PyInstaller 打包后的环境 """
    # 尝试获取 _MEIPASS，如果获取不到（开发环境），就使用当前目录
    base_path = getattr(sys, '_MEIPASS', os.path.abspath("."))
    
    return os.path.join(base_path, relative_path)
def parse_args() -> argparse.Namespace:
    """解析命令行参数"""
    parser = argparse.ArgumentParser(
        description="Psalter LaTeX Generator - 礼仪书LaTeX文档生成工具"
    )
    parser.add_argument(
        '--debug', '-d',
        action='store_true',
        help='启用调试模式（详细日志）'
    )
    parser.add_argument(
        '--log-file', '-l',
        type=str,
        default=None,
        help='日志文件路径'
    )
    parser.add_argument(
        '--config', '-c',
        type=str,
        default=None,
        help='配置文件路径（JSON格式）'
    )
    return parser.parse_args()


def setup_dpi_awareness() -> None:
    """设置Windows DPI感知"""
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass


def setup_window_icon(root: tk.Tk) -> None:
    """设置窗口图标"""
    try:
        from psalter_generator.icon import set_window_icon
        set_window_icon(root)
    except Exception:
        pass


def main() -> int:
    """主函数"""
    args = parse_args()
    
    # 初始化日志
    init_logging(debug=args.debug, log_file=args.log_file)
    
    # 加载配置
    config = None
    if args.config:
        config = AppConfig.load_from_json(args.config)
    
    # Windows DPI适配
    setup_dpi_awareness()
    
    # 创建并运行应用
    try:
        root = tk.Tk()
        # 图标设置移到 PsalterApp 内部，确保窗口完全初始化后再设置
        app = PsalterApp(root, config=config)
        root.mainloop()
        return 0
    except Exception as e:
        print(f"启动失败: {e}", file=sys.stderr)
        return 1


if __name__ == '__main__':
    sys.exit(main())
