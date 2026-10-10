"""02 工资计算与工资条。"""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

from ..options import Options, with_company
from ..styles import B, F, HEADER, INPUT, INPUT_SOFT, L, body, fill_cells, header, title

KEY = "gongzi"
FILENAME = "02_工资计算与工资条.xlsx"
DESC = "社保公积金按比例自动扣，个税按七级累进简化计算，一键生成可打印工资条"

MIN_ROWS = 20

# 示例薪资：工号会按名单重新生成，这里只提供「基本工资 / 绩效 / 补贴 / 缺勤扣款 / 专项附加扣除」
SAMPLE_PAY = [
    (8000, 2500, 500, 0, 2000),
    (15000, 3000, 800, 200, 3000),
    (9000, 1000, 500, 0, 1000),
]


def build(options: Options) -> Workbook:
    wb = Workbook()
    ws = wb.active
    ws.title = "工资明细"
    title(
        ws,
        with_company(options.company, "工资计算表（自动算个税）"),
        13,
        "填写黄色区域；社保公积金按比例自动算，个税按 2026 年综合所得月度预扣简化计算（起征点 5000），仅供参考",
    )
    ws["A3"] = "社保个人比例"
    ws["B3"] = options.social_rate
    ws["C3"] = "公积金比例"
    ws["D3"] = options.fund_rate
    for a in ("B3", "D3"):
        ws[a].number_format = "0.0%"
        ws[a].fill = PatternFill("solid", fgColor=INPUT)

    cols = ["工号", "姓名", "部门", "基本工资", "绩效奖金", "补贴", "缺勤扣款", "应发工资", "社保", "公积金", "专项附加扣除", "个税", "实发工资"]
    header(ws, 5, cols, [8, 10, 10, 11, 11, 9, 10, 12, 10, 10, 12, 10, 12])

    rows = max(MIN_ROWS, len(options.staff))
    last = 5 + rows
    total_row = 6 + rows
    dept_count = len(options.departments)

    for k in range(rows):
        r = 6 + k
        ws.cell(r, 1, f"A{k + 1:03d}")
        ws.cell(r, 2, options.staff[k] if k < len(options.staff) else None)
        ws.cell(r, 3, options.departments[k % dept_count])
        if k < len(SAMPLE_PAY):
            base, bonus, allow, deduct, special = SAMPLE_PAY[k]
            ws.cell(r, 4, base)
            ws.cell(r, 5, bonus)
            ws.cell(r, 6, allow)
            ws.cell(r, 7, deduct)
            ws.cell(r, 11, special)
        ws.cell(r, 8, f'=IF(B{r}="","",D{r}+E{r}+F{r}-G{r})')
        ws.cell(r, 9, f'=IF(B{r}="","",ROUND(D{r}*$B$3,2))')
        ws.cell(r, 10, f'=IF(B{r}="","",ROUND(D{r}*$D$3,2))')
        taxable = f"MAX(H{r}-I{r}-J{r}-K{r}-5000,0)"
        ws.cell(r, 12, f'=IF(B{r}="","",ROUND(MAX({taxable}*{{0.03,0.1,0.2,0.25,0.3,0.35,0.45}}-{{0,210,1410,2660,4410,7160,15160}}),2))')
        ws.cell(r, 13, f'=IF(B{r}="","",H{r}-I{r}-J{r}-L{r})')
        for c in range(4, 14):
            ws.cell(r, c).number_format = "#,##0.00"
    body(ws, 6, last, 1, 13)
    fill_cells(ws, 6, last, (4, 5, 6, 7, 11), INPUT_SOFT)

    ws.cell(total_row, 2, "合计").font = Font(name=F, bold=True)
    for c in range(8, 14):
        ws.cell(total_row, c, f"=SUM({L(c)}6:{L(c)}{last})").number_format = "#,##0.00"
    ws.freeze_panes = "C6"

    # 工资条（打印页）
    s = wb.create_sheet("工资条（打印）")
    title(s, with_company(options.company, "工资条 — 在 A3 输入第几位员工，或整页打印自动生成全部"), 13)
    for i in range(rows):
        hr = 4 + i * 3
        for j, h in enumerate(cols, 1):
            c = s.cell(hr, j, h)
            c.font = Font(name=F, bold=True, size=9, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor=HEADER)
            c.border = B
            c.alignment = Alignment(horizontal="center")
            d = s.cell(hr + 1, j, f'=IF(工资明细!$B${6 + i}="","",工资明细!{L(j)}{6 + i})')
            d.border = B
            d.font = Font(name=F, size=9)
            d.alignment = Alignment(horizontal="center")
            if j >= 4:
                d.number_format = "#,##0.00"
    for j in range(1, 14):
        s.column_dimensions[L(j)].width = 11
    return wb
