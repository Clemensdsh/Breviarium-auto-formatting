#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
models.py - 数据模型定义
包含所有内容项的数据结构，使用dataclass和完整类型注解
"""

from __future__ import annotations
from dataclasses import dataclass, field
from typing import List, Optional, Dict, Tuple
from enum import Enum
import os


class ContentType(Enum):
    """内容类型枚举"""
    H1 = "h1"
    H1CAP = "h1cap"
    H1LOWERCASE = "h1lowercase"
    H2 = "h2"
    H3 = "h3"
    PSALMTITLE = "psalmtitle"
    CANTICLETITLE = "canticletitle"
    HYMNTITLE = "hymntitle"
    HYMNHEADER = "hymnheader"
    ANTIPHON = "antiphon"
    ANTIPHONNUM = "antiphonnum"
    DROPCAP = "dropcap"
    VERSE = "verse"
    GLORIA = "gloria"
    RUBRIC = "rubric"
    V = "V"
    R = "R"
    HYMN = "hymn"
    CAPIT = "capit"
    CAPITHEADER = "capitheader"
    SCRIPTUREREF = "scriptureref"
    COLLECT = "collect"
    LESSON = "lesson"
    TEXT = "text"
    RULE = "rule"
    THICKRULE = "thickrule"
    PAGEBREAK = "pagebreak"
    TOCSTART = "tocstart"
    SINGLECOL = "singlecol"
    IMAGE = "image"


# 格式类型显示名称映射
FORMAT_TYPE_NAMES: Dict[str, str] = {
    "h1": "大标题",
    "h1cap": "目录大标题",
    "h1lowercase": "目录小标题",
    "h2": "副标题",
    "h3": "节次标题",
    "psalmtitle": "圣咏标题",
    "canticletitle": "圣歌标题",
    "hymntitle": "赞美诗标题",
    "hymnheader": "赞美诗加粗标题",
    "antiphon": "对经",
    "antiphonnum": "对经(带编号)",
    "dropcap": "首字下沉文本",
    "verse": "诗节",
    "gloria": "圣三光荣颂",
    "rubric": "礼仪指示",
    "V": "启(V)",
    "R": "应(R)",
    "hymn": "赞美诗节",
    "capit": "短读经",
    "capitheader": "短读经标题",
    "scriptureref": "圣经引用",
    "collect": "集祷经",
    "lesson": "读经标题",
    "text": "普通文本",
    "rule": "分隔线",
    "thickrule": "粗分隔线",
    "pagebreak": "分页",
    "tocstart": "目录起始",
    "singlecol": "单栏/双栏切换",
    "image": "图片",
    "score": "乐谱(Gregorio)",
    # 新增类型
    "imprimatur": "Imprimatur签名块",
    "nihilobstat": "Nihil Obstat签名块",
    "redinline": "行内红字",
    "scriptureinline": "行内圣经引用",
}

# 模块类型显示名称
MODULE_TYPE_NAMES: Dict[str, str] = {
    "matins": "夜课 (Matutinum)",
    "lauds": "赞美经 (Laudes)",
    "prime": "一时经 (Prima)",
    "terce": "三时经 (Tertia)",
    "sext": "六时经 (Sexta)",
    "none": "九时经 (Nona)",
    "vespers": "晚祷 (Vesperae)",
    "compline": "夜祷 (Completorium)",
}

# 特殊显示类型（无需拉丁/中文内容）
SPECIAL_DISPLAY_TYPES: Dict[str, str] = {
    "rule": "[分隔线]",
    "thickrule": "[粗分隔线]",
    "pagebreak": "[分页]",
    "tocstart": "[目录起始]",
    "singlecol": "[单栏/双栏切换]",
    "score": "[乐谱]",
}


def get_format_types_list() -> List[Tuple[str, str]]:
    """返回格式类型列表，用于UI下拉框"""
    return [(k, v) for k, v in FORMAT_TYPE_NAMES.items()]


@dataclass
class ContentItem:
    """单个内容项"""
    item_type: str
    latin: str = ""
    chinese: str = ""
    arg: str = ""
    source_file: str = ""
    is_multiline: bool = False
    line_count: int = 1

    def to_csv_row(self) -> List[str]:
        """转换为CSV行"""
        return [self.item_type, self.latin, self.chinese, self.arg]

    def get_display_text(self) -> str:
        """获取用于列表显示的文本"""
        t = self.item_type
        
        if self.is_multiline:
            lat_preview = self._truncate(self.latin, 20)
            chi_preview = self._truncate(self.chinese, 10)
            return f"[{t}] {lat_preview} | {chi_preview} (+{self.line_count - 1}行)"
        
        # 特殊类型显示
        if t in SPECIAL_DISPLAY_TYPES:
            if t == "score":
                # 乐谱显示文件名
                return f"[乐谱] {os.path.basename(self.latin)}" if self.latin else "[乐谱]"
            return SPECIAL_DISPLAY_TYPES[t]
        
        if t == "image":
            return f"[图片] {os.path.basename(self.latin)}"
        
        # 普通类型：截断显示
        lat_preview = self._truncate(self.latin, 20)
        chi_preview = self._truncate(self.chinese, 10)
        return f"[{t}] {lat_preview} | {chi_preview}"
    
    @staticmethod
    def _truncate(text: str, max_len: int) -> str:
        """截断文本"""
        return text[:max_len] + "..." if len(text) > max_len else text
    
    def validate(self) -> Tuple[bool, str]:
        """验证内容项是否有效"""
        if not self.item_type:
            return False, "类型不能为空"
        if self.item_type not in FORMAT_TYPE_NAMES:
            return False, f"未知类型: {self.item_type}"
        # 非特殊类型需要内容
        if self.item_type not in SPECIAL_DISPLAY_TYPES and self.item_type != "image":
            if not self.latin and not self.chinese:
                return False, "拉丁文和中文不能同时为空"
        return True, ""


@dataclass
class MultiLineContentItem(ContentItem):
    """多行内容项（从文件加载的内容块）"""
    items: List[ContentItem] = field(default_factory=list)

    def __post_init__(self) -> None:
        if self.items:
            first = self.items[0]
            self.item_type = first.item_type
            self.latin = first.latin
            self.chinese = first.chinese
            self.arg = first.arg
            self.is_multiline = True
            self.line_count = len(self.items)

    def to_csv_rows(self) -> List[List[str]]:
        """转换为多行CSV"""
        return [item.to_csv_row() for item in self.items]

    def get_flat_items(self) -> List[ContentItem]:
        """获取扁平化的内容项列表"""
        return self.items


@dataclass
class TitlePageData:
    """封面页数据"""
    title_zh: str = "羅馬大日課\\\\[0.5em]耶穌聖誕瞻禮"
    title_lat: str = "Breviárium Románum\\\\[0.5em]In Nativitáte Dómini"
    edition: str = "中拉對照\\\\[0.5em]Edítio Sínico-Latína"
    footer: str = "Pro Manuscripto"

    def to_dict(self) -> Dict[str, str]:
        return {
            "title_zh": self.title_zh,
            "title_lat": self.title_lat,
            "edition": self.edition,
            "footer": self.footer,
        }

    @classmethod
    def from_dict(cls, data: Dict[str, str]) -> TitlePageData:
        return cls(
            title_zh=data.get("title_zh", ""),
            title_lat=data.get("title_lat", ""),
            edition=data.get("edition", ""),
            footer=data.get("footer", ""),
        )
    
    def validate(self) -> Tuple[bool, str]:
        """验证封面数据"""
        if not self.title_zh and not self.title_lat:
            return False, "至少需要一个标题"
        return True, ""


@dataclass
class ModuleItem:
    """
    模块项 - 包装时辰模块供UI使用
    
    Attributes:
        module_type: 模块类型（matins, lauds, etc.）
        module_id: 模块唯一标识
        display_name: 显示名称
        is_expanded: 是否已展开显示
        config_data: 模块配置数据（字典形式）
        expanded_items: 展开后的ContentItem列表（缓存）
    """
    module_type: str
    module_id: str = ""
    display_name: str = ""
    is_expanded: bool = False
    config_data: Dict = field(default_factory=dict)
    expanded_items: List[ContentItem] = field(default_factory=list)
    
    def __post_init__(self) -> None:
        if not self.module_id:
            import uuid
            self.module_id = f"{self.module_type}_{uuid.uuid4().hex[:8]}"
        if not self.display_name:
            self.display_name = MODULE_TYPE_NAMES.get(self.module_type, self.module_type)
    
    def get_display_text(self) -> str:
        """获取列表显示文本"""
        status = "▼" if self.is_expanded else "▶"
        item_count = len(self.expanded_items) if self.expanded_items else "..."
        return f"{status} [{self.display_name}] ({item_count}项)"
    
    def to_dict(self) -> Dict:
        """序列化为字典"""
        return {
            "module_type": self.module_type,
            "module_id": self.module_id,
            "display_name": self.display_name,
            "config_data": self.config_data,
        }
    
    @classmethod
    def from_dict(cls, data: Dict) -> ModuleItem:
        """从字典反序列化"""
        return cls(
            module_type=data.get("module_type", ""),
            module_id=data.get("module_id", ""),
            display_name=data.get("display_name", ""),
            config_data=data.get("config_data", {}),
        )
