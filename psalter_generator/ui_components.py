#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ui_components.py - 自定义UI组件
包含自定义滚动条、分隔面板等可复用组件
Magnificat礼仪风格
"""

from __future__ import annotations
import tkinter as tk
from tkinter import ttk
from typing import Callable, Optional, Any

from .theme import Theme, DEFAULT_THEME


class TelegramScrollbar(tk.Canvas):
    """简洁风格的自定义滚动条"""
    
    def __init__(
        self, 
        parent: tk.Widget, 
        command: Optional[Callable[..., Any]] = None, 
        theme: Theme = DEFAULT_THEME, 
        **kwargs: Any
    ):
        super().__init__(parent, width=8, highlightthickness=0, bg=theme.BG_LIGHT, **kwargs)
        
        self.theme = theme
        self.command = command
        self.thumb_pos: float = 0.0
        self.thumb_size: float = 0.3
        self.dragging: bool = False
        self.drag_start: float = 0.0
        self.hover: bool = False
        
        self.bind('<Button-1>', self._on_click)
        self.bind('<B1-Motion>', self._on_drag)
        self.bind('<ButtonRelease-1>', lambda e: setattr(self, 'dragging', False))
        self.bind('<Configure>', self._draw)
        self.bind('<Enter>', self._on_enter)
        self.bind('<Leave>', self._on_leave)
    
    def set(self, first: str, last: str) -> None:
        self.thumb_pos = float(first)
        self.thumb_size = float(last) - float(first)
        self._draw()
    
    def _draw(self, event: Optional[tk.Event] = None) -> None:
        self.delete('all')
        h = self.winfo_height()
        w = self.winfo_width()
        
        if h <= 1:
            return
        
        self.create_rectangle(0, 0, w, h, fill=self.theme.BG_LIGHT, outline='')
        
        thumb_height = max(30, h * self.thumb_size)
        thumb_top = self.thumb_pos * (h - thumb_height) / (1 - self.thumb_size) if self.thumb_size < 1 else 0
        
        # 礼仪红色滚动条
        color = self.theme.RUBRIC_RED if self.hover else self.theme.SCROLL_FG
        self.create_rectangle(2, thumb_top + 2, w - 2, thumb_top + thumb_height - 2, 
                            fill=color, outline='', width=0)
    
    def _on_click(self, event: tk.Event) -> None:
        h = self.winfo_height()
        thumb_height = max(30, h * self.thumb_size)
        thumb_top = self.thumb_pos * (h - thumb_height) / (1 - self.thumb_size) if self.thumb_size < 1 else 0
        
        if thumb_top <= event.y <= thumb_top + thumb_height:
            self.dragging = True
            self.drag_start = event.y - thumb_top
        elif self.command:
            self.command('moveto', str(event.y / h))
    
    def _on_drag(self, event: tk.Event) -> None:
        if not self.dragging or not self.command:
            return
        
        h = self.winfo_height()
        thumb_height = max(30, h * self.thumb_size)
        new_top = event.y - self.drag_start
        new_pos = new_top / (h - thumb_height) * (1 - self.thumb_size)
        new_pos = max(0, min(1 - self.thumb_size, new_pos))
        self.command('moveto', str(new_pos))
    
    def _on_enter(self, event: tk.Event) -> None:
        self.hover = True
        self._draw()
    
    def _on_leave(self, event: tk.Event) -> None:
        self.hover = False
        self._draw()


class SmoothPanedWindow(tk.Frame):
    """
    平滑可调整大小的分隔面板 - 修复闪动问题
    使用place布局而非pack，支持任意位置拖动
    """
    
    def __init__(self, parent: tk.Widget, theme: Theme = DEFAULT_THEME, 
                 initial_ratio: float = 0.35, **kwargs: Any):
        super().__init__(parent, bg=theme.BG_DARK, **kwargs)
        
        self.theme = theme
        self.ratio = initial_ratio  # 左侧面板占比
        self.min_left = 180
        self.min_right = 200
        self.sash_width = 6
        
        # 创建三个区域
        self.left_frame = tk.Frame(self, bg=theme.BG_LIGHT)
        self.sash = tk.Frame(self, bg=theme.BORDER, width=self.sash_width, cursor='sb_h_double_arrow')
        self.right_frame = tk.Frame(self, bg=theme.BG_LIGHT)
        
        # 使用place布局
        self.bind('<Configure>', self._on_configure)
        
        # 分隔条事件
        self.sash.bind('<Button-1>', self._start_drag)
        self.sash.bind('<B1-Motion>', self._do_drag)
        self.sash.bind('<ButtonRelease-1>', self._end_drag)
        self.sash.bind('<Enter>', lambda e: self.sash.config(bg=theme.RUBRIC_RED))
        self.sash.bind('<Leave>', self._on_sash_leave)
        
        self._dragging = False
        self._drag_start_x = 0
        self._initial_layout_done = False
    
    def _on_configure(self, event: tk.Event) -> None:
        """窗口大小改变时重新布局"""
        if event.widget == self:
            self._layout()
    
    def _layout(self) -> None:
        """执行布局"""
        total_width = self.winfo_width()
        total_height = self.winfo_height()
        
        if total_width < 10 or total_height < 10:
            return
        
        # 计算左侧宽度
        left_width = int(total_width * self.ratio)
        left_width = max(self.min_left, min(total_width - self.min_right - self.sash_width, left_width))
        
        sash_x = left_width
        right_x = left_width + self.sash_width
        right_width = total_width - right_x
        
        # 使用place精确定位，避免闪动
        self.left_frame.place(x=0, y=0, width=left_width, height=total_height)
        self.sash.place(x=sash_x, y=0, width=self.sash_width, height=total_height)
        self.right_frame.place(x=right_x, y=0, width=right_width, height=total_height)
        
        self._initial_layout_done = True
    
    def _start_drag(self, event: tk.Event) -> None:
        self._dragging = True
        self._drag_start_x = event.x_root
        self._initial_ratio = self.ratio
    
    def _do_drag(self, event: tk.Event) -> None:
        if not self._dragging:
            return
        
        total_width = self.winfo_width()
        if total_width < 10:
            return
        
        delta_x = event.x_root - self._drag_start_x
        delta_ratio = delta_x / total_width
        
        new_ratio = self._initial_ratio + delta_ratio
        
        # 限制范围
        min_ratio = self.min_left / total_width
        max_ratio = (total_width - self.min_right - self.sash_width) / total_width
        new_ratio = max(min_ratio, min(max_ratio, new_ratio))
        
        self.ratio = new_ratio
        self._layout()
    
    def _end_drag(self, event: tk.Event) -> None:
        self._dragging = False
    
    def _on_sash_leave(self, event: tk.Event) -> None:
        if not self._dragging:
            self.sash.config(bg=self.theme.BORDER)


# 保持向后兼容
PanedWindow = SmoothPanedWindow


class StyledButton(tk.Label):
    """样式化按钮 - Magnificat风格"""
    
    def __init__(
        self, 
        parent: tk.Widget, 
        text: str, 
        command: Callable[[], None],
        bg_color: str, 
        theme: Theme = DEFAULT_THEME,
        width: Optional[int] = None,
        font_size: str = "normal",
        **kwargs: Any
    ):
        # 确定前景色：深色背景用白字，浅色背景用深字
        fg_color = "#FFFFFF" if self._is_dark_color(bg_color) else theme.TEXT
        
        super().__init__(
            parent, text=text, bg=bg_color, fg=fg_color,
            font=theme.get_font(font_size), cursor='hand2', padx=12, pady=6, **kwargs
        )
        
        self.theme = theme
        self.bg_color = bg_color
        self.fg_color = fg_color
        self._command = command
        
        if width:
            self.config(width=width)
        
        self.bind('<Enter>', self._on_enter)
        self.bind('<Leave>', self._on_leave)
        self.bind('<Button-1>', self._on_click)
    
    def _is_dark_color(self, color: str) -> bool:
        """判断是否为深色"""
        c = color.lstrip('#')
        if len(c) != 6:
            return False
        r, g, b = int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16)
        luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255
        return luminance < 0.5
    
    def _on_enter(self, event: tk.Event) -> None:
        self.config(bg=self.theme.darken(self.bg_color, 15))
    
    def _on_leave(self, event: tk.Event) -> None:
        self.config(bg=self.bg_color)
    
    def _on_click(self, event: tk.Event) -> None:
        if self._command:
            self._command()


class ButtonFactory:
    """按钮工厂 - Magnificat风格"""
    
    STYLE_COLORS = {
        "default": "BG_HOVER",
        "accent": "RUBRIC_RED",      # 使用礼仪红
        "primary": "RUBRIC_RED",     # 主要按钮
        "success": "SUCCESS",
        "warning": "WARNING",
        "danger": "RUBRIC_RED",      # 危险按钮也用礼仪红
        "purple": "RUBRIC_RED",      # 统一为礼仪红
        "compile": "COMPILE_BTN",
        "secondary": "BG_LIGHT",     # 次要按钮 - 白色
    }
    
    def __init__(self, theme: Theme = DEFAULT_THEME):
        self.theme = theme
    
    def create(
        self, 
        parent: tk.Widget, 
        text: str, 
        command: Callable[[], None],
        style: str = "default", 
        width: Optional[int] = None,
        font_size: str = "normal"
    ) -> StyledButton:
        color_attr = self.STYLE_COLORS.get(style, "BG_HOVER")
        bg_color = getattr(self.theme, color_attr)
        return StyledButton(parent, text, command, bg_color, self.theme, width, font_size)


class ScrollableListbox(tk.Frame):
    """带自定义滚动条的列表框"""
    
    def __init__(self, parent: tk.Widget, theme: Theme = DEFAULT_THEME, **kwargs: Any):
        super().__init__(parent, bg=theme.BG_LIGHT)
        
        self.theme = theme
        self.listbox = tk.Listbox(
            self, bg=theme.BG_LIGHT, fg=theme.TEXT,
            selectbackground=theme.RUBRIC_RED,  # 选中时用礼仪红
            selectforeground="#FFFFFF",
            font=theme.get_font(), borderwidth=0, highlightthickness=0,
            activestyle='none', **kwargs
        )
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.scrollbar = TelegramScrollbar(self, command=self.listbox.yview, theme=theme)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox.config(yscrollcommand=self.scrollbar.set)
    
    def insert(self, index: Any, *elements: str) -> None:
        self.listbox.insert(index, *elements)
    
    def delete(self, first: Any, last: Optional[Any] = None) -> None:
        self.listbox.delete(first, last)
    
    def curselection(self) -> tuple:
        return self.listbox.curselection()
    
    def selection_set(self, first: Any, last: Optional[Any] = None) -> None:
        self.listbox.selection_set(first, last)
    
    def bind(self, sequence: Optional[str] = None, func: Optional[Callable] = None, add: Optional[str] = None) -> str:
        return self.listbox.bind(sequence, func, add)


class ScrollableTreeview(tk.Frame):
    """带自定义滚动条的树形视图"""
    
    def __init__(self, parent: tk.Widget, theme: Theme = DEFAULT_THEME, **kwargs: Any):
        super().__init__(parent, bg=theme.BG_LIGHT)
        
        self.theme = theme
        
        # 配置Treeview样式 - Magnificat风格
        style = ttk.Style()
        style.configure(
            'Treeview',
            background=theme.BG_LIGHT,
            foreground=theme.TEXT,
            fieldbackground=theme.BG_LIGHT,
            font=theme.get_font(),
            rowheight=26
        )
        style.map(
            'Treeview',
            background=[('selected', theme.RUBRIC_RED)],  # 选中时用礼仪红
            foreground=[('selected', '#FFFFFF')]
        )
        
        self.tree = ttk.Treeview(self, show='tree', selectmode='browse', **kwargs)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.scrollbar = TelegramScrollbar(self, command=self.tree.yview, theme=theme)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.config(yscrollcommand=self.scrollbar.set)
    
    def insert(self, parent: str, index: Any, **kwargs: Any) -> str:
        return self.tree.insert(parent, index, **kwargs)
    
    def delete(self, *items: str) -> None:
        self.tree.delete(*items)
    
    def get_children(self, item: Optional[str] = None) -> tuple:
        return self.tree.get_children(item)
    
    def item(self, item: str, **kwargs: Any) -> dict:
        return self.tree.item(item, **kwargs)
    
    def selection(self) -> tuple:
        return self.tree.selection()


class LargeRadioButton(tk.Frame):
    """大型单选按钮 - 用于乐谱来源选择等"""
    
    def __init__(
        self, 
        parent: tk.Widget, 
        text: str, 
        variable: tk.StringVar, 
        value: str,
        theme: Theme = DEFAULT_THEME,
        description: str = "",
        **kwargs: Any
    ):
        super().__init__(parent, bg=theme.BG_LIGHT, cursor='hand2', **kwargs)
        
        self.theme = theme
        self.variable = variable
        self.value = value
        
        # 外框
        self.config(highlightthickness=2, highlightbackground=theme.BORDER)
        
        # 内容
        inner = tk.Frame(self, bg=theme.BG_LIGHT, padx=15, pady=12)
        inner.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        self.title_label = tk.Label(
            inner, text=text, bg=theme.BG_LIGHT, fg=theme.TEXT,
            font=theme.get_font("title", bold=True)
        )
        self.title_label.pack(anchor='w')
        
        # 描述
        if description:
            self.desc_label = tk.Label(
                inner, text=description, bg=theme.BG_LIGHT, fg=theme.TEXT_SEC,
                font=theme.get_font("small"), wraplength=200, justify=tk.LEFT
            )
            self.desc_label.pack(anchor='w', pady=(5, 0))
        
        # 绑定点击事件
        self.bind('<Button-1>', self._on_click)
        inner.bind('<Button-1>', self._on_click)
        self.title_label.bind('<Button-1>', self._on_click)
        if description:
            self.desc_label.bind('<Button-1>', self._on_click)
        
        # 监听变量变化
        variable.trace_add('write', self._on_var_change)
        self._update_style()
    
    def _on_click(self, event: tk.Event) -> None:
        self.variable.set(self.value)
    
    def _on_var_change(self, *args: Any) -> None:
        self._update_style()
    
    def _update_style(self) -> None:
        """更新选中/未选中样式"""
        if self.variable.get() == self.value:
            self.config(highlightbackground=self.theme.RUBRIC_RED, highlightthickness=3)
            self.title_label.config(fg=self.theme.RUBRIC_RED)
        else:
            self.config(highlightbackground=self.theme.BORDER, highlightthickness=2)
            self.title_label.config(fg=self.theme.TEXT)
