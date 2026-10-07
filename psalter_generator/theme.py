#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
theme.py - UI主题与样式配置
以红白黑礼仪色为基调，参考Magnificat风格设计
"""

from __future__ import annotations
from dataclasses import dataclass
from typing import Tuple


@dataclass(frozen=True)
class Theme:
    """UI主题配置类（不可变）- Magnificat礼仪风格"""
    
    # === 核心礼仪色 ===
    # 背景色 - 奶白色系
    BG_DARK: str = "#F5F2EE"       # 主背景 - 温暖的奶白
    BG_LIGHT: str = "#FFFFFF"      # 内容区背景 - 纯白
    BG_HOVER: str = "#E8E4DF"      # 悬停背景
    BG_PANEL: str = "#FAF8F5"      # 面板背景 - 更浅的奶白
    
    # 礼仪红 - 核心强调色（统一色号）
    RUBRIC_RED: str = "#B22222"    # 主红色（火砖红 - 经典礼仪红）
    RUBRIC_DARK: str = "#8B0000"   # 深红（悬停/激活）
    
    # 文字色 - 黑色系
    TEXT: str = "#1A1A1A"          # 主文字 - 深黑
    TEXT_SEC: str = "#666666"      # 次要文字 - 灰色

    # 状态色
    SUCCESS: str = "#2E7D32"       # 成功绿
    WARNING: str = "#E65100"       # 警告橙（更深一点）
    DANGER: str = "#B22222"        # 危险红（统一礼仪红）
    
    # 边框与分隔线
    BORDER: str = "#D0C8C0"        # 主边框色
    SCROLL_FG: str = "#C0B8B0"     # 滚动条
    
    # 按钮色
    COMPILE_BTN: str = "#B22222"   # 编译按钮

    # 字体配置
    FONT_FAMILY: str = "Segoe UI"
    FONT_MONO: str = "Consolas"
    FONT_SIZE_NORMAL: int = 10
    FONT_SIZE_TITLE: int = 12
    FONT_SIZE_SMALL: int = 9
    FONT_SIZE_LARGE: int = 14

    def darken(self, color: str, amount: int = 20) -> str:
        """将颜色变暗"""
        c = color.lstrip('#')
        r = max(0, int(c[0:2], 16) - amount)
        g = max(0, int(c[2:4], 16) - amount)
        b = max(0, int(c[4:6], 16) - amount)
        return f'#{r:02x}{g:02x}{b:02x}'
    
    def get_font(self, size: str = "normal", bold: bool = False) -> Tuple[str, int, str]:
        """获取字体元组"""
        sizes = {
            "small": self.FONT_SIZE_SMALL,
            "normal": self.FONT_SIZE_NORMAL,
            "title": self.FONT_SIZE_TITLE,
            "large": self.FONT_SIZE_LARGE,
        }
        weight = "bold" if bold else "normal"
        return (self.FONT_FAMILY, sizes.get(size, self.FONT_SIZE_NORMAL), weight)
    
    def get_mono_font(self, size: str = "normal") -> Tuple[str, int]:
        """获取等宽字体元组"""
        sizes = {
            "small": self.FONT_SIZE_SMALL,
            "normal": self.FONT_SIZE_NORMAL,
            "title": self.FONT_SIZE_TITLE,
        }
        return (self.FONT_MONO, sizes.get(size, self.FONT_SIZE_NORMAL))


DEFAULT_THEME = Theme()
