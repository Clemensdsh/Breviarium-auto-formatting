"""窗口图标（圣斯德望十字架）"""

import logging
import os
import tkinter as tk

from .paths import ICON_PATH

logger = logging.getLogger(__name__)


def set_icon(window: tk.Wm, default: bool = False) -> None:
    """为窗口设置图标；default=True 时同时作为后续所有窗口的默认图标"""
    if not os.path.exists(ICON_PATH):
        return
    try:
        photo = tk.PhotoImage(file=ICON_PATH)
        window.iconphoto(default, photo)
        window._icon_photo = photo  # 保留引用，防止被回收
    except tk.TclError as e:
        logger.warning(f"设置窗口图标失败: {e}")
