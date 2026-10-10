"""模板参数：公司名、名单、部门、比例这类可变量集中在这里管理。

参数有三个来源，优先级从高到低：命令行显式指定 > 配置文件 > 默认值。
"""

from __future__ import annotations

import json
from dataclasses import dataclass, fields
from datetime import date
from pathlib import Path
from typing import Any, Iterable

DEFAULT_STAFF: tuple[str, ...] = ("张三", "李四", "王五", "赵六", "钱七")
DEFAULT_DEPARTMENTS: tuple[str, ...] = ("销售部", "技术部", "财务部")


def split_list(raw: str | Iterable[str] | None) -> tuple[str, ...] | None:
    """把 "a,b" / "a，b" 拆成元组。空输入返回 None（表示"没指定"）。"""
    if raw is None:
        return None
    if not isinstance(raw, str):
        items = tuple(str(x).strip() for x in raw if str(x).strip())
        return items or None
    parts = (p.strip() for p in raw.replace("，", ",").split(","))
    items = tuple(p for p in parts if p)
    return items or None


def read_names(path: str | Path) -> tuple[str, ...]:
    """从文本文件读名单：一行一个，忽略空行与 # 开头的行。"""
    try:
        text = Path(path).read_text(encoding="utf-8")
    except OSError as exc:
        raise ValueError(f"读取名单文件失败：{path}（{exc}）") from exc
    names = tuple(ln.strip() for ln in text.splitlines() if ln.strip() and not ln.lstrip().startswith("#"))
    if not names:
        raise ValueError(f"名单文件里没有有效姓名：{path}")
    return names


def with_company(company: str, text: str) -> str:
    """给标题加公司名前缀；没填公司名就原样返回。"""
    return f"{company} · {text}" if company else text


@dataclass(frozen=True)
class Options:
    """生成模板时用到的全部可变参数。"""

    company: str = ""
    staff: tuple[str, ...] = DEFAULT_STAFF
    departments: tuple[str, ...] = DEFAULT_DEPARTMENTS
    year: int = 2026
    month: int = 10
    social_rate: float = 0.105
    fund_rate: float = 0.07
    start_date: date = date(2026, 10, 12)

    def __post_init__(self) -> None:
        if not 1 <= self.month <= 12:
            raise ValueError(f"月份必须在 1-12 之间，收到 {self.month}")
        if not 0 <= self.social_rate <= 1:
            raise ValueError(f"社保比例必须在 0-1 之间，收到 {self.social_rate}")
        if not 0 <= self.fund_rate <= 1:
            raise ValueError(f"公积金比例必须在 0-1 之间，收到 {self.fund_rate}")
        if not self.staff:
            raise ValueError("员工名单不能为空")
        if not self.departments:
            raise ValueError("部门列表不能为空")

    @property
    def month_tag(self) -> str:
        """形如 2026-10 的月份标记，用于记账本的汇总月份。"""
        return f"{self.year}-{self.month:02d}"

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> Options:
        raw = dict(data)
        for key in ("staff", "departments"):
            if key in raw and not isinstance(raw[key], (list, tuple)):
                raw[key] = split_list(str(raw[key]))
        if "start_date" in raw and isinstance(raw["start_date"], str):
            try:
                raw["start_date"] = date.fromisoformat(raw["start_date"])
            except ValueError as exc:
                raise ValueError(f"start_date 需要 YYYY-MM-DD 格式：{raw['start_date']}") from exc
        for key in ("social_rate", "fund_rate"):
            if key in raw and isinstance(raw[key], str):
                raw[key] = float(raw[key])

        known = {f.name for f in fields(cls)}
        unknown = set(raw) - known
        if unknown:
            raise ValueError(f"未知参数：{', '.join(sorted(unknown))}")

        return cls(**{k: (tuple(v) if isinstance(v, list) else v) for k, v in raw.items()})

    @classmethod
    def from_json(cls, path: str | Path) -> Options:
        try:
            data = json.loads(Path(path).read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as exc:
            raise ValueError(f"读取配置文件失败：{path}（{exc}）") from exc
        if not isinstance(data, dict):
            raise ValueError(f"配置文件内容需要是一个 JSON 对象：{path}")
        return cls.from_dict(data)
