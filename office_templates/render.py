"""把声明式 Spec 渲染成 Workbook。

公式与条件格式里可以用这些占位符：
    {r}      当前数据行
    {r1}     数据区首行
    {last}   数据区末行
    {total}  合计行（没开 totals 时等于 {last}）
"""

from __future__ import annotations

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation

from .i18n import t
from .options import Options, with_company
from .spec import ColumnSpec, SheetSpec, Spec
from .styles import F, body, header, title

DEFAULT_WIDTH = 12


def render_spec(spec: Spec, options: Options | None = None) -> Workbook:
    """按模板定义生成一个 Workbook（不落盘）。"""
    opts = options or Options()
    wb = Workbook()
    wb.remove(wb.active)  # 去掉默认空表，完全按定义来

    for sheet in spec.sheets:
        ws = wb.create_sheet(sheet.name)
        _render_sheet(ws, sheet, spec, opts)
    return wb


def _render_sheet(ws, sheet: SheetSpec, spec: Spec, options: Options) -> None:
    ncols = len(sheet.columns)
    r1 = sheet.header_row + 1
    last = sheet.header_row + sheet.rows
    total_row = (last + 1) if sheet.totals else last
    fmt = {"r1": r1, "last": last, "total": total_row}

    heading = sheet.title or spec.title or sheet.name
    subtitle = sheet.subtitle or spec.subtitle
    title(ws, with_company(options.company, heading), ncols, subtitle or None)

    header(
        ws,
        sheet.header_row,
        [c.label for c in sheet.columns],
        [c.width or DEFAULT_WIDTH for c in sheet.columns],
    )

    for k in range(sheet.rows):
        r = r1 + k
        for j, col in enumerate(sheet.columns, 1):
            cell = ws.cell(r, j)
            if col.formula:
                cell.value = col.formula.format(r=r, **fmt)
            elif k < len(col.sample):
                cell.value = col.sample[k]
            if col.format:
                cell.number_format = col.format

    body(ws, r1, last, 1, ncols)

    # 填写区底色放在 body 之后，避免被隔行底色覆盖
    for j, col in enumerate(sheet.columns, 1):
        if not col.fill:
            continue
        fill = PatternFill("solid", fgColor=col.fill)
        for r in range(r1, last + 1):
            ws.cell(r, j).fill = fill

    for j, col in enumerate(sheet.columns, 1):
        if not col.choices:
            continue
        dv = DataValidation(type="list", formula1='"' + ",".join(col.choices) + '"', allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(f"{L(j)}{r1}:{L(j)}{last}")

    if sheet.totals:
        ws.cell(total_row, 1, t("render.total")).font = Font(name=F, bold=True)
        for j, col in enumerate(sheet.columns, 1):
            if not col.total:
                continue
            cell = ws.cell(total_row, j, f"=SUM({L(j)}{r1}:{L(j)}{last})")
            if col.format:
                cell.number_format = col.format

    for cond in sheet.conditionals:
        font = Font(color=cond.font_color, bold=cond.bold) if (cond.font_color or cond.bold) else None
        fill = PatternFill("solid", fgColor=cond.fill) if cond.fill else None
        ws.conditional_formatting.add(
            cond.range.format(**fmt),
            FormulaRule(formula=[cond.when.format(**fmt)], font=font, fill=fill),
        )

    if sheet.freeze:
        ws.freeze_panes = sheet.freeze
