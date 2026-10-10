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

try:
    import yaml
except ImportError:  # PyYAML 是可选依赖，用到 YAML 时才需要
    yaml = None

ALLOWED_TYPES = ("text", "number", "date")


def _known(cls: type) -> set[str]:
    return {f.name for f in fields(cls)}


def _reject_unknown(cls: type, data: dict[str, Any], where: str) -> None:
    unknown = set(data) - _known(cls)
    if unknown:
        raise ValueError(f"{where}里有未知字段：{', '.join(sorted(unknown))}")


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
        _reject_unknown(cls, data, "条件格式")
        if "range" not in data or "when" not in data:
            raise ValueError("条件格式必须同时有 range 和 when")
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
            raise ValueError(f"列「{self.label}」的 type 只能是 {'/'.join(ALLOWED_TYPES)}，收到 {self.type}")

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> ColumnSpec:
        _reject_unknown(cls, data, "列定义")
        if not data.get("label"):
            raise ValueError("每一列必须有 label")
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
            raise ValueError(f"工作表「{self.name}」至少要有一列")
        if self.rows < 1:
            raise ValueError(f"工作表「{self.name}」的 rows 至少为 1")

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> SheetSpec:
        _reject_unknown(cls, data, "工作表定义")
        if not data.get("name"):
            raise ValueError("每个工作表必须有 name")
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
            raise ValueError("模板必须有 filename")
        if not self.filename.endswith(".xlsx"):
            raise ValueError(f"filename 必须以 .xlsx 结尾：{self.filename}")
        if not self.sheets:
            raise ValueError(f"模板「{self.filename}」至少要有一个工作表")

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> Spec:
        _reject_unknown(cls, data, "模板定义")
        raw = dict(data)
        raw["sheets"] = [SheetSpec.from_dict(s) for s in raw.get("sheets", [])]
        return cls(**raw)


def load_spec(path: str | Path) -> Spec:
    """从 YAML / JSON 文件读取模板定义。"""
    p = Path(path)
    try:
        text = p.read_text(encoding="utf-8")
    except OSError as exc:
        raise ValueError(f"读取模板定义失败：{path}（{exc}）") from exc

    if p.suffix.lower() == ".json":
        try:
            data = json.loads(text)
        except json.JSONDecodeError as exc:
            raise ValueError(f"模板定义不是合法 JSON：{path}（{exc}）") from exc
    else:
        if yaml is None:
            raise ValueError(f"读取 YAML 模板定义需要 PyYAML：pip install pyyaml（{path}）")
        try:
            data = yaml.safe_load(text)
        except yaml.YAMLError as exc:
            raise ValueError(f"模板定义不是合法 YAML：{path}（{exc}）") from exc

    if not isinstance(data, dict):
        raise ValueError(f"模板定义内容需要是一个映射（键值对）：{path}")
    return Spec.from_dict(data)
