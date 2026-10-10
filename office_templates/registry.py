"""模板注册表：把模板模块收集起来，并负责解析选择与落盘。"""

from __future__ import annotations

import logging
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Iterable, Sequence

from openpyxl import Workbook

from .i18n import t
from .options import Options
from .templates import MODULES

logger = logging.getLogger(__name__)


@dataclass(frozen=True)
class Template:
    """一个可生成的模板。"""

    key: str
    filename: str
    desc: str
    build: Callable[[Options], Workbook]


TEMPLATES: tuple[Template, ...] = tuple(
    Template(key=m.KEY, filename=m.FILENAME, desc=m.DESC, build=m.build) for m in MODULES
)

_BY_KEY = {tpl.key: tpl for tpl in TEMPLATES}
_BY_INDEX = {str(i): tpl for i, tpl in enumerate(TEMPLATES, 1)}


def all_templates() -> tuple[Template, ...]:
    return TEMPLATES


def resolve(specs: Iterable[str] | None) -> list[Template]:
    """把选择项解析成模板列表。

    支持编号（1、01）与名称（kaoqin），每项内部还可用逗号分隔。
    不选则返回全部。重复项会自动去重，并保持注册表原有顺序。
    """
    if not specs:
        return list(TEMPLATES)

    picked: list[Template] = []
    seen: set[str] = set()
    for raw in specs:
        for part in str(raw).split(","):
            token = part.strip()
            if not token:
                continue
            if token.isdigit():
                tpl = _BY_INDEX.get(str(int(token)))
            else:
                tpl = _BY_KEY.get(token.lower())
            if tpl is None:
                raise ValueError(t("err.unknown_template", token=token))
            if tpl.key not in seen:
                seen.add(tpl.key)
                picked.append(tpl)

    if not picked:
        raise ValueError(t("err.no_template"))
    return picked


def build_templates(
    templates: Sequence[Template],
    out_dir: str | Path,
    options: Options | None = None,
) -> list[Path]:
    """生成指定模板到 out_dir，返回生成的文件路径。

    options 省略时用默认参数（示例公司、示例名单）。
    """
    opts = options or Options()
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)

    paths: list[Path] = []
    for tpl in templates:
        wb = tpl.build(opts)
        path = out / tpl.filename
        wb.save(path)
        paths.append(path)
        logger.info(t("msg.generated", path=path))
    return paths
