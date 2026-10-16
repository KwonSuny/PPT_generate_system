"""Lucide 本地图标库加载器。

assets/icons/ 下收录了常用 Lucide 图标的官方 SVG（24x24 stroke 风格）。
本模块负责：名称归一化（kebab-case、旧版别名）、磁盘读取、内存缓存。
渲染层拿到的返回值是 <svg> 的 innerXML（不含外层 <svg> 标签），
由 renderer 统一补壳、上色、控制线宽。

安全说明：只读 assets/icons/ 目录下的 .svg 文件；icon_name 会被严格
归一化，任何含路径分隔符 / .. 的输入都会被拒绝，杜绝路径穿越。
"""

import os
import re
from functools import lru_cache
from typing import Optional
from xml.etree import ElementTree

_ICONS_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "assets", "icons")

# 旧版 Lucide 名称 -> 新版名称（Writer/Planner 可能输出旧名）
_ALIASES = {
    "check-circle": "circle-check",
    "check_circle": "circle-check",
    "alert-circle": "circle-alert",
    "bar-chart": "chart-column",
    "bar-chart-3": "chart-column",
    "bar_chart_3": "chart-column",
    "user": "user",
}

_NAME_RE = re.compile(r"^[a-z0-9][a-z0-9-]*$")


def normalize_icon_name(name: Optional[str]) -> Optional[str]:
    """icon_name 归一化：小写、下划线转连字符、应用别名表。"""
    if not name or not isinstance(name, str):
        return None
    candidate = name.strip().lower().replace("_", "-").replace(" ", "-")
    if not _NAME_RE.match(candidate):
        return None  # 拒绝路径分隔符、.. 等危险输入
    return _ALIASES.get(candidate, candidate)


@lru_cache(maxsize=1)
def _available_set() -> frozenset:
    return frozenset(available_icons())


@lru_cache(maxsize=256)
def load_lucide_svg(icon_name: Optional[str]) -> Optional[str]:
    """读取图标 SVG，返回 innerXML 字符串；不存在时返回 None。"""
    normalized = normalize_icon_name(icon_name)
    if normalized is None:
        return None
    path = os.path.join(_ICONS_DIR, f"{normalized}.svg")
    if not os.path.isfile(path):
        return None
    try:
        tree = ElementTree.parse(path)
        root = tree.getroot()
        inner_parts = [ElementTree.tostring(child, encoding="unicode") for child in root]
        inner = "".join(inner_parts).strip()
        return inner or None
    except ElementTree.ParseError:
        return None


def resolve_icon_name(name: Optional[str]) -> Optional[str]:
    """归一化 + 模糊匹配：白名单外名称按 token 重叠度映射到最近似图标。

    匹配优先级：别名/精确命中 > 子串包含 > token 重叠（长度 >= 2 才计分）。
    无法可靠匹配时返回 None，由渲染层回退占位圆点。
    """
    normalized = normalize_icon_name(name)
    if normalized is None:
        return None
    if normalized in _available_set():
        return normalized
    available = _available_set()
    if not available:
        return None

    tokens = {t for t in normalized.split("-") if len(t) >= 2}
    best, best_score = None, 0
    for cand in available:
        cand_tokens = {t for t in cand.split("-") if len(t) >= 2}
        if normalized in cand or cand in normalized:
            # 子串包含：如 mountain-peak -> mountain
            score = 10 + len(tokens & cand_tokens)
        else:
            score = len(tokens & cand_tokens)
        if score > best_score:
            best, best_score = cand, score
    return best if best_score >= 1 else None


def available_icons() -> list:
    """列出图标库中所有可用图标名（供测试与 prompt 生成用）。"""
    if not os.path.isdir(_ICONS_DIR):
        return []
    return sorted(f[:-4] for f in os.listdir(_ICONS_DIR) if f.endswith(".svg"))
