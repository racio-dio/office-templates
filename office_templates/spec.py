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
import string
from dataclasses import dataclass, field, fields
from pathlib import Path
from typing import Any

from .i18n import t

try:
    import yaml
except ImportError:  # PyYAML 是可选依赖，用到 YAML 时才需要
    yaml = None

ALLOWED_TYPES = ("text", "number", "date")

# 占位符：公式里可以用当前行 {r}；条件格式作用于整片区域，没有"当前行"
ROW_PLACEHOLDERS = ("r", "r1", "last", "total")
RANGE_PLACEHOLDERS = ("r1", "last", "total")

SHEET_NAME_MAX = 31
SHEET_NAME_BAD = set('\\/?*[]:')
DV_LIST_MAX = 255  # Excel 对下拉列表字面量的长度限制


def check_placeholders(text: str, allowed: tuple[str, ...]) -> None:
    """提前发现拼错的占位符，免得渲染时才抛 KeyError。"""
    try:
        names = [name for _, name, _, _ in string.Formatter().parse(text) if name is not None]
    except ValueError as exc:  # 不成对的花括号
        raise ValueError(t("err.placeholder", text=text, name="?", allowed=", ".join(allowed))) from exc
    for name in names:
        if name not in allowed:
            raise ValueError(t("err.placeholder", text=text, name=name, allowed=", ".join(allowed)))


def _as_int(name: str, value: Any) -> int:
    if isinstance(value, bool):
        raise ValueError(t("err.not_int", field=name, value=value))
    if isinstance(value, int):
        return value
    try:
        return int(str(value).strip())
    except ValueError:
        raise ValueError(t("err.not_int", field=name, value=value)) from None


def _as_list(name: str, value: Any) -> list:
    if value is None:
        return []
    if not isinstance(value, list):
        raise ValueError(t("err.not_list", field=name))
    return value


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

    def __post_init__(self) -> None:
        self.range = str(self.range)
        self.when = str(self.when)
        check_placeholders(self.range, RANGE_PLACEHOLDERS)
        check_placeholders(self.when, RANGE_PLACEHOLDERS)


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
        self.label = str(self.label)
        if self.formula:
            check_placeholders(self.formula, ROW_PLACEHOLDERS)
        if self.choices:
            self.choices = [str(c) for c in self.choices]
            for c in self.choices:
                if "," in c:
                    raise ValueError(t("err.choice_comma", label=self.label, value=c))
            if len(",".join(self.choices)) > DV_LIST_MAX:
                raise ValueError(t("err.choices_too_long", label=self.label))

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> ColumnSpec:
        _reject_unknown(cls, data, "where.column")
        if not data.get("label"):
            raise ValueError(t("err.column_label"))
        raw = dict(data)
        if "choices" in raw and raw["choices"] is not None:
            raw["choices"] = _as_list("choices", raw["choices"])
        if "sample" in raw:
            raw["sample"] = _as_list("sample", raw["sample"])
        return cls(**raw)


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
    total_label: str = ""
    conditionals: list[ConditionalSpec] = field(default_factory=list)

    def __post_init__(self) -> None:
        self.name = str(self.name)
        if len(self.name) > SHEET_NAME_MAX:
            raise ValueError(t("err.sheet_name_length", name=self.name))
        if SHEET_NAME_BAD & set(self.name):
            raise ValueError(t("err.sheet_name_chars", name=self.name))
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
        for key in ("rows", "header_row"):
            if key in raw:
                raw[key] = _as_int(key, raw[key])
        raw["columns"] = [ColumnSpec.from_dict(c) for c in _as_list("columns", raw.get("columns"))]
        raw["conditionals"] = [ConditionalSpec.from_dict(c) for c in _as_list("conditionals", raw.get("conditionals"))]
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
        if "\\" in self.filename or self.filename != Path(self.filename).name or self.filename.startswith("."):
            raise ValueError(t("err.filename_plain", value=self.filename))
        if not self.sheets:
            raise ValueError(t("err.sheets_required", filename=self.filename))
        seen: set[str] = set()
        for sheet in self.sheets:
            low = sheet.name.lower()  # Excel 的表名不区分大小写
            if low in seen:
                raise ValueError(t("err.sheet_name_dup", name=sheet.name))
            seen.add(low)
            minimum = 3 if (sheet.subtitle or self.subtitle) else 2
            if sheet.header_row < minimum:
                raise ValueError(t("err.header_row", name=sheet.name, min=minimum))

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> Spec:
        _reject_unknown(cls, data, "where.spec")
        raw = dict(data)
        raw["sheets"] = [SheetSpec.from_dict(s) for s in _as_list("sheets", raw.get("sheets"))]
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
