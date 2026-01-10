#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
latex_generator.py - LaTeX内容生成器
负责将内容项转换为LaTeX代码
分离状态机与模板渲染逻辑
"""

from __future__ import annotations
import logging
from typing import List, Optional, Protocol
from abc import ABC, abstractmethod
from dataclasses import dataclass
from enum import Enum, auto

from .models import ContentItem, MultiLineContentItem, TitlePageData
from .config import TexMappingConfig
from .exceptions import TemplateError

logger = logging.getLogger(__name__)


# ==========================================
# 状态机
# ==========================================

class ColumnMode(Enum):
    """分栏模式"""
    DOUBLE = auto()
    SINGLE = auto()


@dataclass
class GeneratorState:
    """生成器状态"""
    mode: ColumnMode = ColumnMode.DOUBLE
    in_paracol: bool = True
    
    @property
    def is_single_col(self) -> bool:
        return self.mode == ColumnMode.SINGLE
    
    def toggle_mode(self) -> ColumnMode:
        """切换模式，返回新模式"""
        if self.mode == ColumnMode.DOUBLE:
            self.mode = ColumnMode.SINGLE
            self.in_paracol = False
        else:
            self.mode = ColumnMode.DOUBLE
            self.in_paracol = True
        return self.mode
    
    def set_double(self) -> None:
        self.mode = ColumnMode.DOUBLE
        self.in_paracol = True
    
    def set_single(self) -> None:
        self.mode = ColumnMode.SINGLE
        self.in_paracol = False
    
    def reset(self) -> None:
        self.mode = ColumnMode.DOUBLE
        self.in_paracol = True


# ==========================================
# 内容项渲染器（策略模式）
# ==========================================

class ItemRenderer(ABC):
    """内容项渲染器基类"""
    
    @abstractmethod
    def can_render(self, item_type: str) -> bool:
        """是否能渲染此类型"""
        pass
    
    @abstractmethod
    def render(self, item: ContentItem, state: GeneratorState) -> List[str]:
        """渲染内容项，返回LaTeX行列表"""
        pass


class MappedItemRenderer(ItemRenderer):
    """基于映射表的渲染器"""
    
    def __init__(self, mapping: TexMappingConfig):
        self.mapping = mapping
    
    def can_render(self, item_type: str) -> bool:
        return self.mapping.has(item_type)
    
    def render(self, item: ContentItem, state: GeneratorState) -> List[str]:
        cmd = self.mapping.get(item.item_type)
        if not cmd:
            return []
        
        formatted = cmd.format(
            state.is_single_col,
            l=item.latin,
            c=item.chinese,
            a=item.arg
        )
        return [formatted]


class ControlItemRenderer(ItemRenderer):
    """控制类型渲染器（tocstart, singlecol, pagebreak）"""
    
    CONTROL_TYPES = {'tocstart', 'singlecol', 'pagebreak'}
    
    def can_render(self, item_type: str) -> bool:
        return item_type in self.CONTROL_TYPES
    
    def render(self, item: ContentItem, state: GeneratorState) -> List[str]:
        if item.item_type == 'tocstart':
            return self._render_tocstart(state)
        elif item.item_type == 'singlecol':
            return self._render_singlecol(state)
        elif item.item_type == 'pagebreak':
            return self._render_pagebreak(state)
        return []
    
    def _render_tocstart(self, state: GeneratorState) -> List[str]:
        """渲染目录起始"""
        lines: List[str] = []
        
        if state.in_paracol:
            lines.append(r"\end{paracol}")
        
        lines.extend([
            r"\psPrintToc",
            # psPrintToc 内部已经有 clearpage，不需要额外添加
            r"\pagenumbering{arabic}",
            r"\pagestyle{fancy}",
            r"\begin{paracol}{2}",
        ])
        
        state.set_double()
        return lines
    
    def _render_singlecol(self, state: GeneratorState) -> List[str]:
        """渲染单栏/双栏切换"""
        if state.is_single_col:
            state.set_double()
            return [r"\psExitSingleCol"]
        else:
            state.set_single()
            return [r"\psEnterSingleCol"]
    
    def _render_pagebreak(self, state: GeneratorState) -> List[str]:
        """渲染分页"""
        if state.is_single_col:
            return [r"\psSinglePageBreak"]
        else:
            return [r"\psPageBreak"]


class SpecialItemRenderer(ItemRenderer):
    """特殊类型渲染器（antiphonnum, image）"""
    
    SPECIAL_TYPES = {'antiphonnum', 'image'}
    
    def can_render(self, item_type: str) -> bool:
        return item_type in self.SPECIAL_TYPES
    
    def render(self, item: ContentItem, state: GeneratorState) -> List[str]:
        if item.item_type == 'antiphonnum':
            return self._render_antiphonnum(item, state)
        elif item.item_type == 'image':
            return self._render_image(item, state)
        return []
    
    def _render_antiphonnum(self, item: ContentItem, state: GeneratorState) -> List[str]:
        """渲染带编号的对经"""
        if state.is_single_col:
            return [rf"\psSingleAntiphonNum{{{item.arg}}}{{{item.chinese}}}"]
        else:
            return [rf"\psAntiphonNum{{{item.arg}}}{{{item.latin}}}{{{item.chinese}}}"]
    
    def _render_image(self, item: ContentItem, state: GeneratorState) -> List[str]:
        """渲染图片"""
        if state.is_single_col:
            return [rf"\psSingleImage{{{item.latin}}}"]
        else:
            return [rf"\psImageFullWidth{{{item.latin}}}"]


class ScoreItemRenderer(ItemRenderer):
    """
    乐谱渲染器 - 使用GregorioTeX插入乐谱
    
    自动处理单栏/双栏切换：
    1. 如果当前是双栏，先切换到单栏
    2. 插入乐谱
    3. 切换回原来的模式
    
    ContentItem字段用法:
    - latin: .gabc文件路径（不带扩展名）或内联GABC代码
    - chinese: 乐谱标题/注释（可选）
    - arg: 选项，如 "inline" 表示内联GABC，否则为文件路径
    """
    
    def can_render(self, item_type: str) -> bool:
        return item_type == 'score'
    
    def render(self, item: ContentItem, state: GeneratorState) -> List[str]:
        lines: List[str] = []
        was_double_col = not state.is_single_col
        
        # 1. 如果当前是双栏，切换到单栏
        if was_double_col:
            lines.append(r"% === 乐谱开始：切换到单栏 ===")
            lines.append(r"\psEnterSingleCol")
            state.set_single()
        
        # 2. 添加乐谱标题（如果有）
        if item.chinese:
            lines.append(rf"\psSingleRubric{{{item.chinese}}}")
        
        # 3. 插入乐谱
        if item.arg == "inline":
            # 内联GABC代码
            lines.append(r"\gresetlyriccentering{firstletter}")
            lines.append(rf"\gabcsnippet{{{item.latin}}}")
        else:
            # 外部.gabc文件
            # 去掉.gabc扩展名（如果有）
            score_path = item.latin
            if score_path.endswith('.gabc'):
                score_path = score_path[:-5]
            lines.append(rf"\gregorioscore{{{score_path}}}")
        
        # 4. 添加乐谱后的间距
        lines.append(r"\vspace{0.5em}")
        
        # 5. 如果之前是双栏，切换回双栏
        if was_double_col:
            lines.append(r"\psExitSingleCol")
            lines.append(r"% === 乐谱结束：恢复双栏 ===")
            state.set_double()
        
        return lines


# ==========================================
# 主生成器
# ==========================================

class TexGenerator:
    """LaTeX内容生成器"""
    
    def __init__(self, mapping_config: Optional[TexMappingConfig] = None):
        self.mapping = mapping_config or TexMappingConfig()
        self.state = GeneratorState()
        
        # 注册渲染器（按优先级顺序）
        self.renderers: List[ItemRenderer] = [
            ControlItemRenderer(),
            ScoreItemRenderer(),
            SpecialItemRenderer(),
            MappedItemRenderer(self.mapping),
        ]
    
    def reset(self) -> None:
        """重置生成器状态"""
        self.state.reset()
    
    def generate(self, items: List[ContentItem]) -> str:
        """从内容项列表生成LaTeX代码"""
        self.reset()
        
        # 展开多行内容项
        flat_items = self._flatten_items(items)
        
        lines: List[str] = [r"\begin{paracol}{2}"]
        
        for item in flat_items:
            item_lines = self._render_item(item)
            lines.extend(item_lines)
        
        # 确保正确关闭paracol
        if self.state.in_paracol:
            lines.append(r"\end{paracol}")
        
        logger.debug(f"生成了 {len(lines)} 行LaTeX代码")
        return "\n".join(lines)
    
    def _flatten_items(self, items: List[ContentItem]) -> List[ContentItem]:
        """展开多行内容项"""
        flat: List[ContentItem] = []
        for item in items:
            if isinstance(item, MultiLineContentItem):
                flat.extend(item.get_flat_items())
            else:
                flat.append(item)
        return flat
    
    def _render_item(self, item: ContentItem) -> List[str]:
        """渲染单个内容项"""
        for renderer in self.renderers:
            if renderer.can_render(item.item_type):
                return renderer.render(item, self.state)
        
        # 未知类型
        logger.warning(f"未知内容类型: {item.item_type}")
        return [f"% 未知类型: {item.item_type} | {item.latin} | {item.chinese}"]


# ==========================================
# 模板生成器
# ==========================================

class MainTexGenerator:
    """主文档生成器 - 处理main.tex模板"""
    
    PLACEHOLDERS = {
        "%TITLE_ZH%": "title_zh",
        "%TITLE_LAT%": "title_lat",
        "%EDITION_INFO%": "edition",
        "%FOOTER_TEXT%": "footer",
    }
    
    def __init__(self, template_path: str):
        self.template_path = template_path
        self._template_content: Optional[str] = None
    
    def load_template(self) -> str:
        """加载模板内容"""
        if self._template_content is None:
            try:
                with open(self.template_path, 'r', encoding='utf-8') as f:
                    self._template_content = f.read()
                logger.debug(f"已加载模板: {self.template_path}")
            except IOError as e:
                raise TemplateError(f"无法加载模板", self.template_path, str(e))
        return self._template_content
    
    def generate(self, title_data: TitlePageData) -> str:
        """生成带有标题信息的main.tex内容"""
        content = self.load_template()
        data_dict = title_data.to_dict()
        
        for placeholder, attr in self.PLACEHOLDERS.items():
            value = data_dict.get(attr, "")
            content = content.replace(placeholder, value)
        
        return content
    
    def reload_template(self) -> None:
        """强制重新加载模板"""
        self._template_content = None


class BodyTexGenerator:
    """body.tex生成器"""
    
    HEADER = "% Generated by Psalter LaTeX Generator\n"
    
    def __init__(self, mapping_config: Optional[TexMappingConfig] = None):
        self.tex_generator = TexGenerator(mapping_config)
    
    def generate(self, items: List[ContentItem]) -> str:
        """生成body.tex内容"""
        content = self.tex_generator.generate(items)
        return self.HEADER + content
    
    def save(self, items: List[ContentItem], filepath: str) -> None:
        """生成并保存到文件"""
        content = self.generate(items)
        try:
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(content)
            logger.info(f"已保存body.tex: {filepath}")
        except IOError as e:
            raise TemplateError(f"保存文件失败", filepath, str(e))
