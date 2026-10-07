#!/usr/bin/env python3
"""程序入口"""

import argparse
import sys
import tkinter as tk

from psalter_generator import PsalterApp, init_logging, AppConfig


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Psalter LaTeX Generator")
    parser.add_argument('--debug', '-d', action='store_true', help='启用调试日志')
    parser.add_argument('--log-file', '-l', help='日志文件路径')
    parser.add_argument('--config', '-c', help='配置文件路径（JSON）')
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    init_logging(debug=args.debug, log_file=args.log_file)
    config = AppConfig.load_from_json(args.config) if args.config else None

    # Windows DPI 适配
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    try:
        root = tk.Tk()
        PsalterApp(root, config=config)
        root.mainloop()
        return 0
    except Exception as e:
        print(f"启动失败: {e}", file=sys.stderr)
        return 1


if __name__ == '__main__':
    sys.exit(main())
