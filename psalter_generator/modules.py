#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
modules.py - 时辰模块定义
封装礼仪时辰的结构，支持模块化编辑和参数化生成
支持1962年旧礼（Extraordinary Form）和新礼（Ordinary Form / LOTH）
"""

from __future__ import annotations
from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from typing import List, Optional, Dict, Any, Callable
from enum import Enum
import copy

from .models import ContentItem


class HourType(Enum):
    """时辰类型枚举"""
    # === 1962年旧礼 (Extraordinary Form) ===
    MATINS_1962 = "matins_1962"           # 旧礼夜课 (Matutinum)
    LAUDS_1962 = "lauds_1962"             # 旧礼赞美经 (Laudes)
    PRIME_1962 = "prime_1962"             # 旧礼一时经 (Prima)
    TERCE_1962 = "terce_1962"             # 旧礼三时经 (Tertia)
    SEXT_1962 = "sext_1962"               # 旧礼六时经 (Sexta)
    NONE_1962 = "none_1962"               # 旧礼九时经 (Nona)
    VESPERS_1962 = "vespers_1962"         # 旧礼晚祷 (Vesperae)
    COMPLINE_1962 = "compline_1962"       # 旧礼夜祷 (Completorium)
    
    # === 新礼 (Ordinary Form / LOTH) ===
    OFFICE_OF_READINGS = "office_readings"  # 新礼诵读 (Officium lectionis)
    LAUDS_LOTH = "lauds_loth"               # 新礼晨祷 (Laudes)
    TERCE_LOTH = "terce_loth"               # 新礼日间祈祷-午前 (Tertia)
    SEXT_LOTH = "sext_loth"                 # 新礼日间祈祷-午时 (Sexta)
    NONE_LOTH = "none_loth"                 # 新礼日间祈祷-午后 (Nona)
    VESPERS_LOTH = "vespers_loth"           # 新礼晚祷 (Vesperae)
    COMPLINE_LOTH = "compline_loth"         # 新礼夜祷 (Completorium)


class ComponentType(Enum):
    """组件类型枚举"""
    NOCTURN = "nocturn"                     # 单个夜课单元
    PSALM_WITH_ANTIPHON = "psalm_antiphon"  # 带对经的圣咏
    LESSON_WITH_RESPONSORY = "lesson_resp"  # 读经+答唱
    PRECES = "preces"                       # 祷词组
    HYMN = "hymn"                           # 赞美诗
    CANTICLE = "canticle"                   # 圣歌
    CHAPTER = "chapter"                     # 短读经
    COLLECT = "collect"                     # 集祷经
    VERSICLE = "versicle"                   # 启应对
    INTERCESSIONS = "intercessions"         # 信友祷词 (新礼)


# 时辰类型显示名称
HOUR_TYPE_NAMES: Dict[str, str] = {
    # 旧礼 1962
    "matins_1962": "旧礼夜课 (Matutinum_1962)",
    "lauds_1962": "旧礼赞美经 (Laudes_1962)",
    "prime_1962": "旧礼一时经 (Prima_1962)",
    "terce_1962": "旧礼三时经 (Tertia_1962)",
    "sext_1962": "旧礼六时经 (Sexta_1962)",
    "none_1962": "旧礼九时经 (Nona_1962)",
    "vespers_1962": "旧礼晚祷 (Vesperae_1962)",
    "compline_1962": "旧礼夜祷 (Completorium_1962)",
    # 新礼 LOTH
    "office_readings": "新礼诵读 (Officium lectionis)",
    "lauds_loth": "新礼晨祷 (Laudes)",
    "terce_loth": "新礼日间祈祷-午前 (Tertia)",
    "sext_loth": "新礼日间祈祷-午时 (Sexta)",
    "none_loth": "新礼日间祈祷-午后 (Nona)",
    "vespers_loth": "新礼晚祷 (Vesperae)",
    "compline_loth": "新礼夜祷 (Completorium)",
}

# 组件类型显示名称
COMPONENT_TYPE_NAMES: Dict[str, str] = {
    "nocturn": "夜课单元",
    "psalm_antiphon": "圣咏（带对经）",
    "lesson_resp": "读经+答唱",
    "preces": "祷词组",
    "hymn": "赞美诗",
    "canticle": "圣歌",
    "chapter": "短读经",
    "collect": "集祷经",
    "versicle": "启应对",
    "intercessions": "信友祷词",
}


@dataclass
class ModuleSlot:
    """模块插槽 - 定义模块中可填充的位置"""
    slot_id: str
    slot_type: str
    label_lat: str
    label_zh: str
    required: bool = True
    default_content: Optional[ContentItem] = None
    content: Optional[ContentItem] = None
    
    def is_filled(self) -> bool:
        return self.content is not None or self.default_content is not None
    
    def get_content(self) -> Optional[ContentItem]:
        return self.content if self.content is not None else self.default_content
    
    def clear(self) -> None:
        self.content = None


@dataclass
class ModuleConfig:
    """模块配置参数"""
    title_lat: str = ""
    title_zh: str = ""
    nocturn_count: int = 3
    lessons_per_nocturn: int = 3
    psalms_per_nocturn: int = 3
    psalm_count: int = 3
    show_gloria: bool = True
    show_antiphon_repeat: bool = True
    
    def to_dict(self) -> Dict[str, Any]:
        return {
            "title_lat": self.title_lat,
            "title_zh": self.title_zh,
            "nocturn_count": self.nocturn_count,
            "lessons_per_nocturn": self.lessons_per_nocturn,
            "psalms_per_nocturn": self.psalms_per_nocturn,
            "psalm_count": self.psalm_count,
            "show_gloria": self.show_gloria,
            "show_antiphon_repeat": self.show_antiphon_repeat,
        }
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> ModuleConfig:
        return cls(**{k: v for k, v in data.items() if hasattr(cls, k)})


class BaseComponent(ABC):
    """组件基类"""
    
    def __init__(self, component_id: str, name_lat: str, name_zh: str):
        self.component_id = component_id
        self.name_lat = name_lat
        self.name_zh = name_zh
        self.slots: List[ModuleSlot] = []
        self._custom_items: List[ContentItem] = []
    
    @abstractmethod
    def _create_slots(self) -> List[ModuleSlot]:
        pass
    
    def initialize(self) -> None:
        self.slots = self._create_slots()
    
    def get_slot(self, slot_id: str) -> Optional[ModuleSlot]:
        for slot in self.slots:
            if slot.slot_id == slot_id:
                return slot
        return None
    
    def set_slot_content(self, slot_id: str, content: ContentItem) -> bool:
        slot = self.get_slot(slot_id)
        if slot:
            slot.content = content
            return True
        return False
    
    def add_custom_item(self, item: ContentItem, position: int = -1) -> None:
        if position < 0 or position >= len(self._custom_items):
            self._custom_items.append(item)
        else:
            self._custom_items.insert(position, item)
    
    def remove_custom_item(self, index: int) -> bool:
        if 0 <= index < len(self._custom_items):
            self._custom_items.pop(index)
            return True
        return False
    
    @abstractmethod
    def expand(self) -> List[ContentItem]:
        pass
    
    def validate(self) -> tuple[bool, List[str]]:
        errors = []
        for slot in self.slots:
            if slot.required and not slot.is_filled():
                errors.append(f"缺少必填项: {slot.label_zh}")
        return len(errors) == 0, errors


class PsalmWithAntiphonComponent(BaseComponent):
    """带对经的圣咏组件"""
    
    def __init__(self, psalm_num: int = 1, component_id: str = ""):
        super().__init__(
            component_id or f"psalm_ant_{psalm_num}",
            f"Psalmus {psalm_num}",
            f"圣咏 {psalm_num}"
        )
        self.psalm_num = psalm_num
        self.show_gloria = True
        self.repeat_antiphon = True
        self.initialize()
    
    def _create_slots(self) -> List[ModuleSlot]:
        slots = [
            ModuleSlot("antiphon", "antiphon", "Antiphona", "对经", True),
            ModuleSlot("psalm_title", "psalm_title", f"Psalmus {self.psalm_num}", f"圣咏 {self.psalm_num}", True),
        ]
        # 默认7个诗节插槽
        for i in range(1, 8):
            slots.append(ModuleSlot(
                f"verse_{i}", "verse", f"Versus {i}", f"诗节 {i}", 
                required=(i == 1)  # 只有第一个是必需的
            ))
        slots.append(ModuleSlot("gloria", "gloria", "Gloria Patri", "圣三光荣颂", required=False,
                  default_content=ContentItem(
                      item_type="gloria",
                      latin="Glória Patri, et Fílio, * et Spirítui Sancto. Sicut erat in princípio, et nunc, et semper, * et in sǽcula sæculórum. Amen.",
                      chinese="愿光荣归于父、及子、及圣神。起初如何，今日亦然，直到永远。阿门。"
                  )))
        return slots
    
    def expand(self) -> List[ContentItem]:
        items: List[ContentItem] = []
        
        ant_slot = self.get_slot("antiphon")
        if ant_slot and ant_slot.is_filled():
            ant = ant_slot.get_content()
            items.append(ContentItem(item_type="antiphon", latin=ant.latin, chinese=ant.chinese))
        
        title_slot = self.get_slot("psalm_title")
        if title_slot and title_slot.is_filled():
            title = title_slot.get_content()
            items.append(ContentItem(item_type="psalmtitle", latin=title.latin, chinese=title.chinese))
        
        # 处理所有诗节插槽
        for slot in self.slots:
            if slot.slot_id.startswith("verse_") and slot.is_filled():
                verse = slot.get_content()
                items.append(ContentItem(item_type="verse", latin=verse.latin, chinese=verse.chinese))
        
        if self.show_gloria:
            gloria_slot = self.get_slot("gloria")
            if gloria_slot and gloria_slot.is_filled():
                gloria = gloria_slot.get_content()
                items.append(ContentItem(item_type="gloria", latin=gloria.latin, chinese=gloria.chinese))
        
        if self.repeat_antiphon and ant_slot and ant_slot.is_filled():
            ant = ant_slot.get_content()
            items.append(ContentItem(item_type="antiphonrepeat", latin=ant.latin, chinese=ant.chinese))
        
        items.extend(self._custom_items)
        return items


class LessonWithResponsoryComponent(BaseComponent):
    """读经+答唱组件"""
    
    def __init__(self, lesson_num: int = 1, component_id: str = ""):
        super().__init__(
            component_id or f"lesson_resp_{lesson_num}",
            f"Lectio {lesson_num}",
            f"读经 {lesson_num}"
        )
        self.lesson_num = lesson_num
        self.initialize()
    
    def _create_slots(self) -> List[ModuleSlot]:
        return [
            ModuleSlot("lesson_title", "lesson_title", f"Lectio {self.lesson_num}", f"读经 {self.lesson_num}", True),
            ModuleSlot("lesson_content", "lesson", "Textus", "正文", True),
            ModuleSlot("responsory", "responsory", "Responsorium", "答唱咏", required=False),
        ]
    
    def expand(self) -> List[ContentItem]:
        items: List[ContentItem] = []
        
        title_slot = self.get_slot("lesson_title")
        if title_slot and title_slot.is_filled():
            title = title_slot.get_content()
            items.append(ContentItem(item_type="lesson", latin=title.latin, chinese=title.chinese))
        
        content_slot = self.get_slot("lesson_content")
        if content_slot and content_slot.is_filled():
            content = content_slot.get_content()
            items.append(ContentItem(item_type="text", latin=content.latin, chinese=content.chinese))
        
        resp_slot = self.get_slot("responsory")
        if resp_slot and resp_slot.is_filled():
            resp = resp_slot.get_content()
            items.append(ContentItem(item_type="V", latin=resp.latin, chinese=resp.chinese))
        
        items.extend(self._custom_items)
        return items


class HourModule(ABC):
    """时辰模块基类"""
    
    def __init__(self, hour_type: HourType, config: Optional[ModuleConfig] = None):
        self.hour_type = hour_type
        self.config = config or ModuleConfig()
        self.components: List[BaseComponent] = []
        self._header_items: List[ContentItem] = []
        self._footer_items: List[ContentItem] = []
    
    @property
    def name_lat(self) -> str:
        return self._get_hour_name_lat()
    
    @property
    def name_zh(self) -> str:
        return self._get_hour_name_zh()
    
    @abstractmethod
    def _get_hour_name_lat(self) -> str:
        pass
    
    @abstractmethod
    def _get_hour_name_zh(self) -> str:
        pass
    
    @abstractmethod
    def _build_structure(self) -> None:
        pass
    
    def initialize(self) -> None:
        self.components.clear()
        self._header_items.clear()
        self._footer_items.clear()
        self._build_structure()
    
    def expand(self) -> List[ContentItem]:
        items: List[ContentItem] = []
        
        # 标题
        items.append(ContentItem(
            item_type="h1cap",
            latin=self._get_hour_name_lat(),
            chinese=self._get_hour_name_zh()
        ))
        items.append(ContentItem(item_type="rule"))
        
        # 头部固定内容
        items.extend(self._header_items)
        
        # 各组件内容
        for component in self.components:
            items.extend(component.expand())
        
        # 尾部固定内容
        items.extend(self._footer_items)
        
        return items
    
    def get_component(self, component_id: str) -> Optional[BaseComponent]:
        for comp in self.components:
            if comp.component_id == component_id:
                return comp
        return None


# ========================================
# 1962年旧礼模块
# ========================================

class MatinsModule1962(HourModule):
    """旧礼夜课模块（Matutinum 1962）"""
    
    def __init__(self, config: Optional[ModuleConfig] = None):
        super().__init__(HourType.MATINS_1962, config)
        self.initialize()
    
    def _get_hour_name_lat(self) -> str:
        return "Ad Matutinum"
    
    def _get_hour_name_zh(self) -> str:
        return "旧礼夜课"
    
    def _build_structure(self) -> None:
        self._header_items = [
            ContentItem(item_type="V", latin="Dómine, lábia mea apéries.", chinese="主啊，求祢开启我的口。"),
            ContentItem(item_type="R", latin="Et os meum annuntiábit laudem tuam.", chinese="我的口要宣扬祢的光荣。"),
            ContentItem(item_type="rubric", latin="Invitatorium", chinese="邀请曲"),
        ]
        
        nocturn_count = self.config.nocturn_count
        psalms_per = self.config.psalms_per_nocturn
        lessons_per = self.config.lessons_per_nocturn
        
        for n in range(1, nocturn_count + 1):
            for p in range(1, psalms_per + 1):
                psalm_num = (n - 1) * psalms_per + p
                psalm_comp = PsalmWithAntiphonComponent(psalm_num, f"nocturn{n}_psalm{p}")
                self.components.append(psalm_comp)
            
            for l in range(1, lessons_per + 1):
                lesson_num = (n - 1) * lessons_per + l
                lesson_comp = LessonWithResponsoryComponent(lesson_num, f"nocturn{n}_lesson{l}")
                self.components.append(lesson_comp)
        
        self._footer_items = [
            ContentItem(item_type="rule"),
            ContentItem(item_type="rubric", latin="Te Deum laudámus.", chinese="天主，我们赞美祢。"),
        ]


class LaudsModule1962(HourModule):
    """旧礼赞美经模块（Laudes 1962）"""
    
    def __init__(self, config: Optional[ModuleConfig] = None):
        super().__init__(HourType.LAUDS_1962, config)
        self.initialize()
    
    def _get_hour_name_lat(self) -> str:
        return "Ad Laudes"
    
    def _get_hour_name_zh(self) -> str:
        return "旧礼赞美经"
    
    def _build_structure(self) -> None:
        self._header_items = [
            ContentItem(item_type="V", latin="Deus, in adjutórium meum inténde.", chinese="天主，求祢快来拯救我。"),
            ContentItem(item_type="R", latin="Dómine, ad adjuvándum me festína.", chinese="主啊，求祢速来扶助我。"),
            ContentItem(item_type="gloria",
                       latin="Glória Patri, et Fílio, * et Spirítui Sancto. Sicut erat in princípio, et nunc, et semper, * et in sǽcula sæculórum. Amen. Allelúia.",
                       chinese="愿光荣归于父、及子、及圣神。起初如何，今日亦然，直到永远。阿门。阿肋路亚。"),
        ]
        
        for i in range(1, 6):
            psalm_comp = PsalmWithAntiphonComponent(i, f"lauds_psalm{i}")
            self.components.append(psalm_comp)
        
        chapter_comp = LessonWithResponsoryComponent(1, "lauds_chapter")
        chapter_comp.name_lat = "Capitulum"
        chapter_comp.name_zh = "短读经"
        self.components.append(chapter_comp)
        
        self._footer_items = [
            ContentItem(item_type="rule"),
            ContentItem(item_type="rubric", latin="Benedictus", chinese="赞主曲"),
        ]


class SmallHourModule1962(HourModule):
    """旧礼小时课模块（Prima, Tertia, Sexta, Nona 1962）"""
    
    def __init__(self, hour_type: HourType, config: Optional[ModuleConfig] = None):
        valid_types = [HourType.PRIME_1962, HourType.TERCE_1962, HourType.SEXT_1962, HourType.NONE_1962]
        if hour_type not in valid_types:
            raise ValueError(f"SmallHourModule1962 不支持 {hour_type}")
        super().__init__(hour_type, config)
        self.initialize()
    
    def _get_hour_name_lat(self) -> str:
        names = {
            HourType.PRIME_1962: "Ad Primam",
            HourType.TERCE_1962: "Ad Tertiam",
            HourType.SEXT_1962: "Ad Sextam",
            HourType.NONE_1962: "Ad Nonam",
        }
        return names.get(self.hour_type, "")
    
    def _get_hour_name_zh(self) -> str:
        names = {
            HourType.PRIME_1962: "旧礼一时经",
            HourType.TERCE_1962: "旧礼三时经",
            HourType.SEXT_1962: "旧礼六时经",
            HourType.NONE_1962: "旧礼九时经",
        }
        return names.get(self.hour_type, "")
    
    def _build_structure(self) -> None:
        self._header_items = [
            ContentItem(item_type="V", latin="Deus, in adjutórium meum inténde.", chinese="天主，求祢快来拯救我。"),
            ContentItem(item_type="R", latin="Dómine, ad adjuvándum me festína.", chinese="主啊，求祢速来扶助我。"),
            ContentItem(item_type="gloria",
                       latin="Glória Patri, et Fílio, * et Spirítui Sancto. Sicut erat in princípio, et nunc, et semper, * et in sǽcula sæculórum. Amen. Allelúia.",
                       chinese="愿光荣归于父、及子、及圣神。起初如何，今日亦然，直到永远。阿门。阿肋路亚。"),
        ]
        
        psalm_count = self.config.psalm_count
        for i in range(1, psalm_count + 1):
            psalm_comp = PsalmWithAntiphonComponent(i, f"{self.hour_type.value}_psalm{i}")
            self.components.append(psalm_comp)
        
        chapter_comp = LessonWithResponsoryComponent(1, f"{self.hour_type.value}_chapter")
        chapter_comp.name_lat = "Capitulum"
        chapter_comp.name_zh = "短读经"
        chapter_comp.slots = [s for s in chapter_comp.slots if s.slot_id != "responsory"]
        self.components.append(chapter_comp)
        
        self._footer_items = [
            ContentItem(item_type="V", latin="Dómine, exáudi oratiónem meam.", chinese="主啊，求祢俯听我的祈祷。"),
            ContentItem(item_type="R", latin="Et clamor meus ad te véniat.", chinese="愿我的呼声上达于祢。"),
        ]


class VespersModule1962(HourModule):
    """旧礼晚祷模块（Vesperae 1962）"""
    
    def __init__(self, config: Optional[ModuleConfig] = None):
        super().__init__(HourType.VESPERS_1962, config)
        self.initialize()
    
    def _get_hour_name_lat(self) -> str:
        return "Ad Vesperas"
    
    def _get_hour_name_zh(self) -> str:
        return "旧礼晚祷"
    
    def _build_structure(self) -> None:
        self._header_items = [
            ContentItem(item_type="V", latin="Deus, in adjutórium meum inténde.", chinese="天主，求祢快来拯救我。"),
            ContentItem(item_type="R", latin="Dómine, ad adjuvándum me festína.", chinese="主啊，求祢速来扶助我。"),
            ContentItem(item_type="gloria",
                       latin="Glória Patri, et Fílio, * et Spirítui Sancto. Sicut erat in princípio, et nunc, et semper, * et in sǽcula sæculórum. Amen. Allelúia.",
                       chinese="愿光荣归于父、及子、及圣神。起初如何，今日亦然，直到永远。阿门。阿肋路亚。"),
        ]
        
        for i in range(1, 6):
            psalm_comp = PsalmWithAntiphonComponent(i, f"vespers_psalm{i}")
            self.components.append(psalm_comp)
        
        chapter_comp = LessonWithResponsoryComponent(1, "vespers_chapter")
        chapter_comp.name_lat = "Capitulum"
        chapter_comp.name_zh = "短读经"
        self.components.append(chapter_comp)
        
        self._footer_items = [
            ContentItem(item_type="rule"),
            ContentItem(item_type="rubric", latin="Magnificat", chinese="圣母赞主曲"),
        ]


class ComplineModule1962(HourModule):
    """旧礼夜祷模块（Completorium 1962）"""
    
    def __init__(self, config: Optional[ModuleConfig] = None):
        super().__init__(HourType.COMPLINE_1962, config)
        self.initialize()
    
    def _get_hour_name_lat(self) -> str:
        return "Ad Completorium"
    
    def _get_hour_name_zh(self) -> str:
        return "旧礼夜祷"
    
    def _build_structure(self) -> None:
        self._header_items = [
            ContentItem(item_type="V", latin="Jube, domne, benedícere.", chinese="父啊，请祝福。"),
            ContentItem(item_type="rubric", 
                       latin="Noctem quiétam et finem perféctum concédat nobis Dóminus omnípotens.",
                       chinese="愿全能的天主赐我们平安的夜晚和善终。"),
            ContentItem(item_type="R", latin="Amen.", chinese="阿门。"),
        ]
        
        for i in range(1, 4):
            psalm_comp = PsalmWithAntiphonComponent(i, f"compline_psalm{i}")
            self.components.append(psalm_comp)
        
        chapter_comp = LessonWithResponsoryComponent(1, "compline_chapter")
        chapter_comp.name_lat = "Capitulum"
        chapter_comp.name_zh = "短读经"
        chapter_comp.slots = [s for s in chapter_comp.slots if s.slot_id != "responsory"]
        self.components.append(chapter_comp)
        
        self._footer_items = [
            ContentItem(item_type="rule"),
            ContentItem(item_type="rubric", latin="Nunc dimittis", chinese="西默盎赞主曲"),
            ContentItem(item_type="rule"),
            ContentItem(item_type="rubric", latin="Salve Regina", chinese="又圣母经"),
        ]


# ========================================
# 新礼模块 (LOTH / Ordinary Form)
# ========================================

class OfficeOfReadingsModule(HourModule):
    """新礼诵读模块（Officium lectionis）"""
    
    def __init__(self, config: Optional[ModuleConfig] = None):
        super().__init__(HourType.OFFICE_OF_READINGS, config)
        self.initialize()
    
    def _get_hour_name_lat(self) -> str:
        return "Officium lectionis"
    
    def _get_hour_name_zh(self) -> str:
        return "新礼诵读"
    
    def _build_structure(self) -> None:
        self._header_items = [
            ContentItem(item_type="V", latin="Deus, in adjutórium meum inténde.", chinese="天主，求祢快来拯救我。"),
            ContentItem(item_type="R", latin="Dómine, ad adjuvándum me festína.", chinese="主啊，求祢速来扶助我。"),
            ContentItem(item_type="gloria",
                       latin="Glória Patri, et Fílio, et Spirítui Sancto. Sicut erat in princípio, et nunc, et semper, et in sǽcula sæculórum. Amen.",
                       chinese="愿光荣归于父、及子、及圣神。起初如何，今日亦然，直到永远。阿门。"),
            ContentItem(item_type="rubric", latin="Hymnus", chinese="赞美诗"),
        ]
        
        # 3首圣咏
        for i in range(1, 4):
            psalm_comp = PsalmWithAntiphonComponent(i, f"readings_psalm{i}")
            self.components.append(psalm_comp)
        
        # 两篇读经
        reading1 = LessonWithResponsoryComponent(1, "first_reading")
        reading1.name_lat = "Lectio prima"
        reading1.name_zh = "第一篇读经"
        self.components.append(reading1)
        
        reading2 = LessonWithResponsoryComponent(2, "second_reading")
        reading2.name_lat = "Lectio altera"
        reading2.name_zh = "第二篇读经"
        self.components.append(reading2)
        
        self._footer_items = [
            ContentItem(item_type="rule"),
            ContentItem(item_type="rubric", latin="Oratio", chinese="结束祷词"),
        ]


class LaudsModuleLOTH(HourModule):
    """新礼晨祷模块（Laudes LOTH）"""
    
    def __init__(self, config: Optional[ModuleConfig] = None):
        super().__init__(HourType.LAUDS_LOTH, config)
        self.initialize()
    
    def _get_hour_name_lat(self) -> str:
        return "Laudes matutinae"
    
    def _get_hour_name_zh(self) -> str:
        return "新礼晨祷"
    
    def _build_structure(self) -> None:
        self._header_items = [
            ContentItem(item_type="V", latin="Deus, in adjutórium meum inténde.", chinese="天主，求祢快来拯救我。"),
            ContentItem(item_type="R", latin="Dómine, ad adjuvándum me festína.", chinese="主啊，求祢速来扶助我。"),
            ContentItem(item_type="gloria",
                       latin="Glória Patri, et Fílio, et Spirítui Sancto. Sicut erat in princípio, et nunc, et semper, et in sǽcula sæculórum. Amen.",
                       chinese="愿光荣归于父、及子、及圣神。起初如何，今日亦然，直到永远。阿门。"),
            ContentItem(item_type="rubric", latin="Hymnus", chinese="赞美诗"),
        ]
        
        # 晨祷圣咏组（圣咏、旧约圣歌、赞美圣咏）
        psalm1 = PsalmWithAntiphonComponent(1, "lauds_psalm1")
        psalm1.name_lat = "Psalmus"
        psalm1.name_zh = "圣咏"
        self.components.append(psalm1)
        
        canticle = PsalmWithAntiphonComponent(2, "lauds_canticle")
        canticle.name_lat = "Canticum"
        canticle.name_zh = "旧约圣歌"
        self.components.append(canticle)
        
        psalm2 = PsalmWithAntiphonComponent(3, "lauds_psalm2")
        psalm2.name_lat = "Psalmus"
        psalm2.name_zh = "赞美圣咏"
        self.components.append(psalm2)
        
        # 短读经
        chapter = LessonWithResponsoryComponent(1, "lauds_chapter")
        chapter.name_lat = "Lectio brevis"
        chapter.name_zh = "短读经"
        chapter.slots = [s for s in chapter.slots if s.slot_id != "responsory"]
        self.components.append(chapter)
        
        self._footer_items = [
            ContentItem(item_type="rubric", latin="Responsorium breve", chinese="短对答咏"),
            ContentItem(item_type="rule"),
            ContentItem(item_type="rubric", latin="Canticum Zachariae (Benedictus)", chinese="匝加利亚赞主曲"),
            ContentItem(item_type="rule"),
            ContentItem(item_type="rubric", latin="Preces", chinese="祷词"),
            ContentItem(item_type="rubric", latin="Pater noster", chinese="天主经"),
            ContentItem(item_type="rubric", latin="Oratio", chinese="结束祷词"),
        ]


class DaytimePrayerModuleLOTH(HourModule):
    """新礼日间祈祷模块（Tertia, Sexta, Nona LOTH）"""
    
    def __init__(self, hour_type: HourType, config: Optional[ModuleConfig] = None):
        valid_types = [HourType.TERCE_LOTH, HourType.SEXT_LOTH, HourType.NONE_LOTH]
        if hour_type not in valid_types:
            raise ValueError(f"DaytimePrayerModuleLOTH 不支持 {hour_type}")
        super().__init__(hour_type, config)
        self.initialize()
    
    def _get_hour_name_lat(self) -> str:
        names = {
            HourType.TERCE_LOTH: "Hora tertia",
            HourType.SEXT_LOTH: "Hora sexta",
            HourType.NONE_LOTH: "Hora nona",
        }
        return names.get(self.hour_type, "")
    
    def _get_hour_name_zh(self) -> str:
        names = {
            HourType.TERCE_LOTH: "新礼日间祈祷-午前",
            HourType.SEXT_LOTH: "新礼日间祈祷-午时",
            HourType.NONE_LOTH: "新礼日间祈祷-午后",
        }
        return names.get(self.hour_type, "")
    
    def _build_structure(self) -> None:
        self._header_items = [
            ContentItem(item_type="V", latin="Deus, in adjutórium meum inténde.", chinese="天主，求祢快来拯救我。"),
            ContentItem(item_type="R", latin="Dómine, ad adjuvándum me festína.", chinese="主啊，求祢速来扶助我。"),
            ContentItem(item_type="gloria",
                       latin="Glória Patri, et Fílio, et Spirítui Sancto. Sicut erat in princípio, et nunc, et semper, et in sǽcula sæculórum. Amen.",
                       chinese="愿光荣归于父、及子、及圣神。起初如何，今日亦然，直到永远。阿门。"),
            ContentItem(item_type="rubric", latin="Hymnus", chinese="赞美诗"),
        ]
        
        # 3首圣咏（通常118篇分段）
        for i in range(1, 4):
            psalm_comp = PsalmWithAntiphonComponent(i, f"{self.hour_type.value}_psalm{i}")
            self.components.append(psalm_comp)
        
        # 短读经
        chapter = LessonWithResponsoryComponent(1, f"{self.hour_type.value}_chapter")
        chapter.name_lat = "Lectio brevis"
        chapter.name_zh = "短读经"
        chapter.slots = [s for s in chapter.slots if s.slot_id != "responsory"]
        self.components.append(chapter)
        
        self._footer_items = [
            ContentItem(item_type="V", latin="Dómine, exáudi oratiónem meam.", chinese="主啊，求祢俯听我的祈祷。"),
            ContentItem(item_type="R", latin="Et clamor meus ad te véniat.", chinese="愿我的呼声上达于祢。"),
            ContentItem(item_type="rubric", latin="Oratio", chinese="结束祷词"),
        ]


class VespersModuleLOTH(HourModule):
    """新礼晚祷模块（Vesperae LOTH）"""
    
    def __init__(self, config: Optional[ModuleConfig] = None):
        super().__init__(HourType.VESPERS_LOTH, config)
        self.initialize()
    
    def _get_hour_name_lat(self) -> str:
        return "Vesperae"
    
    def _get_hour_name_zh(self) -> str:
        return "新礼晚祷"
    
    def _build_structure(self) -> None:
        self._header_items = [
            ContentItem(item_type="V", latin="Deus, in adjutórium meum inténde.", chinese="天主，求祢快来拯救我。"),
            ContentItem(item_type="R", latin="Dómine, ad adjuvándum me festína.", chinese="主啊，求祢速来扶助我。"),
            ContentItem(item_type="gloria",
                       latin="Glória Patri, et Fílio, et Spirítui Sancto. Sicut erat in princípio, et nunc, et semper, et in sǽcula sæculórum. Amen.",
                       chinese="愿光荣归于父、及子、及圣神。起初如何，今日亦然，直到永远。阿门。"),
            ContentItem(item_type="rubric", latin="Hymnus", chinese="赞美诗"),
        ]
        
        # 晚祷圣咏组（2首圣咏 + 新约圣歌）
        psalm1 = PsalmWithAntiphonComponent(1, "vespers_psalm1")
        self.components.append(psalm1)
        
        psalm2 = PsalmWithAntiphonComponent(2, "vespers_psalm2")
        self.components.append(psalm2)
        
        canticle = PsalmWithAntiphonComponent(3, "vespers_canticle")
        canticle.name_lat = "Canticum"
        canticle.name_zh = "新约圣歌"
        self.components.append(canticle)
        
        # 短读经
        chapter = LessonWithResponsoryComponent(1, "vespers_chapter")
        chapter.name_lat = "Lectio brevis"
        chapter.name_zh = "短读经"
        chapter.slots = [s for s in chapter.slots if s.slot_id != "responsory"]
        self.components.append(chapter)
        
        self._footer_items = [
            ContentItem(item_type="rubric", latin="Responsorium breve", chinese="短对答咏"),
            ContentItem(item_type="rule"),
            ContentItem(item_type="rubric", latin="Canticum B. Mariae Virginis (Magnificat)", chinese="圣母赞主曲"),
            ContentItem(item_type="rule"),
            ContentItem(item_type="rubric", latin="Preces", chinese="祷词"),
            ContentItem(item_type="rubric", latin="Pater noster", chinese="天主经"),
            ContentItem(item_type="rubric", latin="Oratio", chinese="结束祷词"),
        ]


class ComplineModuleLOTH(HourModule):
    """新礼夜祷模块（Completorium LOTH）"""
    
    def __init__(self, config: Optional[ModuleConfig] = None):
        super().__init__(HourType.COMPLINE_LOTH, config)
        self.initialize()
    
    def _get_hour_name_lat(self) -> str:
        return "Completorium"
    
    def _get_hour_name_zh(self) -> str:
        return "新礼夜祷"
    
    def _build_structure(self) -> None:
        self._header_items = [
            ContentItem(item_type="V", latin="Deus, in adjutórium meum inténde.", chinese="天主，求祢快来拯救我。"),
            ContentItem(item_type="R", latin="Dómine, ad adjuvándum me festína.", chinese="主啊，求祢速来扶助我。"),
            ContentItem(item_type="gloria",
                       latin="Glória Patri, et Fílio, et Spirítui Sancto. Sicut erat in princípio, et nunc, et semper, et in sǽcula sæculórum. Amen.",
                       chinese="愿光荣归于父、及子、及圣神。起初如何，今日亦然，直到永远。阿门。"),
            ContentItem(item_type="rubric", latin="Examen conscientiae", chinese="省察"),
            ContentItem(item_type="rubric", latin="Hymnus", chinese="赞美诗"),
        ]
        
        # 夜祷圣咏（1首或2首）
        psalm_comp = PsalmWithAntiphonComponent(1, "compline_psalm")
        self.components.append(psalm_comp)
        
        # 短读经
        chapter = LessonWithResponsoryComponent(1, "compline_chapter")
        chapter.name_lat = "Lectio brevis"
        chapter.name_zh = "短读经"
        chapter.slots = [s for s in chapter.slots if s.slot_id != "responsory"]
        self.components.append(chapter)
        
        self._footer_items = [
            ContentItem(item_type="rubric", latin="Responsorium breve", chinese="短对答咏"),
            ContentItem(item_type="rule"),
            ContentItem(item_type="rubric", latin="Canticum Simeonis (Nunc dimittis)", chinese="西默盎赞主曲"),
            ContentItem(item_type="rubric", latin="Oratio", chinese="结束祷词"),
            ContentItem(item_type="rule"),
            ContentItem(item_type="rubric", latin="Antiphona finalis B. Mariae Virginis", chinese="圣母对经"),
        ]


# 模块工厂
def create_hour_module(hour_type: HourType, config: Optional[ModuleConfig] = None) -> HourModule:
    """创建时辰模块的工厂函数"""
    module_classes = {
        # 旧礼 1962
        HourType.MATINS_1962: MatinsModule1962,
        HourType.LAUDS_1962: LaudsModule1962,
        HourType.PRIME_1962: lambda c: SmallHourModule1962(HourType.PRIME_1962, c),
        HourType.TERCE_1962: lambda c: SmallHourModule1962(HourType.TERCE_1962, c),
        HourType.SEXT_1962: lambda c: SmallHourModule1962(HourType.SEXT_1962, c),
        HourType.NONE_1962: lambda c: SmallHourModule1962(HourType.NONE_1962, c),
        HourType.VESPERS_1962: VespersModule1962,
        HourType.COMPLINE_1962: ComplineModule1962,
        # 新礼 LOTH
        HourType.OFFICE_OF_READINGS: OfficeOfReadingsModule,
        HourType.LAUDS_LOTH: LaudsModuleLOTH,
        HourType.TERCE_LOTH: lambda c: DaytimePrayerModuleLOTH(HourType.TERCE_LOTH, c),
        HourType.SEXT_LOTH: lambda c: DaytimePrayerModuleLOTH(HourType.SEXT_LOTH, c),
        HourType.NONE_LOTH: lambda c: DaytimePrayerModuleLOTH(HourType.NONE_LOTH, c),
        HourType.VESPERS_LOTH: VespersModuleLOTH,
        HourType.COMPLINE_LOTH: ComplineModuleLOTH,
    }
    
    creator = module_classes.get(hour_type)
    if creator is None:
        raise ValueError(f"未知的时辰类型: {hour_type}")
    
    if callable(creator) and not isinstance(creator, type):
        return creator(config)
    else:
        return creator(config)
