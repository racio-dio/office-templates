"""声明式模板定义：用 YAML / JSON 描述一张表，不用写 Python。

最小例子：

    filename: 我的台账.xlsx
    title: 我的台账
    sheets:
      - name: 明细
        columns:
          - {label: 日期, type: date, format: yyyy-mm-dd}
          - {label: 金额, type: number, format: "#,##0.00", total: true}
        totals: true
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field, fields
from pathlib import Path
from typing import Any

from .i18n import t

try:
    import yaml
except ImportError:  # PyYAML 是可选依赖，用到 YAML 时才需要
    yaml = None

ALLOWED_TYPES = ("text", "number", "date")


def _known(cls: type) -> set[str]:
    return {f.name for f in fields(cls)}


def _reject_unknown(cls: type, data: dict[str, Any], where_key: str) -> None:
    unknown = set(data) - _known(cls)
    if unknown:
        raise ValueError(t("err.unknown_field", where=t(where_key), names=", ".join(sorted(unknown))))


@dataclass
class ConditionalSpec:
    """条件格式：range / when 支持 {r1} {last} {total} 占位。"""

    range: str
    when: str
    font_color: str = ""
    fill: str = ""
    bold: bool = False

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> ConditionalSpec:
        _reject_unknown(cls, data, "where.conditional")
        if "range" not in data or "when" not in data:
            raise ValueError(t("err.conditional_fields"))
        return cls(**data)


@dataclass
class ColumnSpec:
    """一列的定义。填了 formula 就按公式生成，否则有 sample 就填示例数据。"""

    label: str
    key: str = ""
    width: float | None = None
    type: str = "text"
    format: str = ""
    formula: str = ""
    choices: list[str] | None = None
    sample: list[Any] = field(default_factory=list)
    fill: str = ""
    total: bool = False

    def __post_init__(self) -> None:
        if self.type not in ALLOWED_TYPES:
            raise ValueError(
                t("err.column_type", label=self.label, types="/".join(ALLOWED_TYPES), value=self.type)
            )

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> ColumnSpec:
        _reject_unknown(cls, data, "where.column")
        if not data.get("label"):
            raise ValueError(t("err.column_label"))
        return cls(**data)


@dataclass
class SheetSpec:
    """一张工作表的定义。"""

    name: str
    columns: list[ColumnSpec]
    title: str = ""
    subtitle: str = ""
    header_row: int = 3
    rows: int = 30
    freeze: str = ""
    totals: bool = False
    conditionals: list[ConditionalSpec] = field(default_factory=list)

    def __post_init__(self) -> None:
        if not self.columns:
            raise ValueError(t("err.sheet_columns", name=self.name))
        if self.rows < 1:
            raise ValueError(t("err.sheet_rows", name=self.name))

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> SheetSpec:
        _reject_unknown(cls, data, "where.sheet")
        if not data.get("name"):
            raise ValueError(t("err.sheet_name"))
        raw = dict(data)
        raw["columns"] = [ColumnSpec.from_dict(c) for c in raw.get("columns", [])]
        raw["conditionals"] = [ConditionalSpec.from_dict(c) for c in raw.get("conditionals", [])]
        return cls(**raw)


@dataclass
class Spec:
    """一整个模板（一个 .xlsx 文件）的定义。"""

    filename: str
    sheets: list[SheetSpec]
    title: str = ""
    subtitle: str = ""

    def __post_init__(self) -> None:
        if not self.filename:
            raise ValueError(t("err.filename_required"))
        if not self.filename.endswith(".xlsx"):
            raise ValueError(t("err.filename_suffix", value=self.filename))
        if not self.sheets:
            raise ValueError(t("err.sheets_required", filename=self.filename))

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> Spec:
        _reject_unknown(cls, data, "where.spec")
        raw = dict(data)
        raw["sheets"] = [SheetSpec.from_dict(s) for s in raw.get("sheets", [])]
        return cls(**raw)


def load_spec(path: str | Path) -> Spec:
    """从 YAML / JSON 文件读取模板定义。"""
    p = Path(path)
    try:
        text = p.read_text(encoding="utf-8")
    except OSError as exc:
        raise ValueError(t("err.spec_read", path=path, error=exc)) from exc

    if p.suffix.lower() == ".json":
        try:
            data = json.loads(text)
        except json.JSONDecodeError as exc:
            raise ValueError(t("err.spec_json", path=path, error=exc)) from exc
    else:
        if yaml is None:
            raise ValueError(t("err.spec_needs_yaml", path=path))
        try:
            data = yaml.safe_load(text)
        except yaml.YAMLError as exc:
            raise ValueError(t("err.spec_yaml", path=path, error=exc)) from exc

    if not isinstance(data, dict):
        raise ValueError(t("err.spec_object", path=path))
    return Spec.from_dict(data)
