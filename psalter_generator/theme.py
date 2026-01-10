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
    RUBRIC_LIGHT: str = "#CD5C5C"  # 浅红（次要强调）
    
    # 文字色 - 黑色系
    TEXT: str = "#1A1A1A"          # 主文字 - 深黑
    TEXT_SEC: str = "#666666"      # 次要文字 - 灰色
    TEXT_LIGHT: str = "#999999"    # 浅色文字
    
    # 强调色（兼容旧代码）- 统一使用礼仪红
    ACCENT: str = "#B22222"        # 同礼仪红
    ACCENT_LIGHT: str = "#B22222"  # 同礼仪红
    
    # 状态色
    SUCCESS: str = "#2E7D32"       # 成功绿
    WARNING: str = "#E65100"       # 警告橙（更深一点）
    DANGER: str = "#B22222"        # 危险红（统一礼仪红）
    
    # 边框与分隔线
    BORDER: str = "#D0C8C0"        # 主边框色
    BORDER_LIGHT: str = "#E0D8D0"  # 浅边框
    SCROLL_FG: str = "#C0B8B0"     # 滚动条
    
    # 按钮色 - 统一使用礼仪红系
    PURPLE: str = "#B22222"        # 统一为礼仪红（原紫色按钮）
    COMPILE_BTN: str = "#B22222"   # 编译按钮
    COMPILE_BTN_HOVER: str = "#8B0000"  # 编译按钮悬停
    
    # 圆角配置
    CORNER_RADIUS: int = 6         # 默认圆角半径
    CORNER_RADIUS_SMALL: int = 4   # 小圆角
    CORNER_RADIUS_LARGE: int = 10  # 大圆角
    
    # 字体配置
    FONT_FAMILY: str = "Segoe UI"
    FONT_FAMILY_CN: str = "Microsoft YaHei"
    FONT_MONO: str = "Consolas"
    FONT_SIZE_NORMAL: int = 10
    FONT_SIZE_TITLE: int = 12
    FONT_SIZE_SMALL: int = 9
    FONT_SIZE_LARGE: int = 14

    def lighten(self, color: str, amount: int = 20) -> str:
        """将颜色变亮"""
        c = color.lstrip('#')
        r = min(255, int(c[0:2], 16) + amount)
        g = min(255, int(c[2:4], 16) + amount)
        b = min(255, int(c[4:6], 16) + amount)
        return f'#{r:02x}{g:02x}{b:02x}'
    
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


# 默认礼仪风格主题
DEFAULT_THEME = Theme()

# 保留深色主题选项（可通过设置切换）
DARK_THEME = Theme(
    BG_DARK="#17212b",
    BG_LIGHT="#242f3d",
    BG_HOVER="#2b5278",
    BG_PANEL="#1e2a36",
    RUBRIC_RED="#B22222",
    RUBRIC_DARK="#8B0000",
    RUBRIC_LIGHT="#CD5C5C",
    TEXT="#f5f5f5",
    TEXT_SEC="#8b9ba5",
    TEXT_LIGHT="#6b7b85",
    ACCENT="#B22222",
    ACCENT_LIGHT="#CD5C5C",
    SUCCESS="#50a550",
    WARNING="#d4a535",
    DANGER="#B22222",
    BORDER="#3d4d5c",
    BORDER_LIGHT="#4d5d6c",
    SCROLL_FG="#4a5d6e",
    PURPLE="#B22222",
    COMPILE_BTN="#B22222",
    COMPILE_BTN_HOVER="#8B0000",
)

# 别名 - 向后兼容
LIGHT_THEME = DEFAULT_THEME
