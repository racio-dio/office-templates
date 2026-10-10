"""命令行与错误提示的双语文案。

语言优先级：--lang > 环境变量 OFFICE_TEMPLATES_LANG > LANG/LC_ALL/LC_MESSAGES > 默认中文。
作为库调用时，模块加载会按环境自动选一次语言，可用 set_lang() 覆盖。
"""

from __future__ import annotations

import os

DEFAULT_LANG = "zh"
SUPPORTED_LANGS = ("zh", "en")

MESSAGES: dict[str, dict[str, str]] = {
    # 命令行
    "cli.description": {
        "zh": "生成 5 套「填数即自动计算」的中文办公 Excel 模板",
        "en": "Generate 5 Excel templates that calculate as soon as you fill them in",
    },
    "cli.out_dir": {"zh": "输出目录，默认 dist/", "en": "output directory (default: dist/)"},
    "cli.only": {
        "zh": "只生成指定内置模板，编号或名称均可，空格或逗号分隔",
        "en": "build only these built-in templates, by number or key, space or comma separated",
    },
    "cli.list": {"zh": "列出所有可用模板后退出", "en": "list available templates and exit"},
    "cli.quiet": {"zh": "只输出错误信息", "en": "only report errors"},
    "cli.lang": {"zh": "界面语言：zh 中文 / en English", "en": "interface language: zh Chinese / en English"},
    "cli.group.params": {
        "zh": "模板参数（影响生成的模板内容）",
        "en": "template options (change what goes into the files)",
    },
    "cli.company": {"zh": "公司名，会加在每套模板的标题前", "en": "company name, prefixed to every template title"},
    "cli.staff": {"zh": "员工名单，逗号分隔", "en": "staff names, comma separated"},
    "cli.staff_file": {
        "zh": "员工名单文件，一行一个姓名，# 开头为注释",
        "en": "file of staff names, one per line, # starts a comment",
    },
    "cli.departments": {"zh": "部门列表，逗号分隔", "en": "departments, comma separated"},
    "cli.year": {"zh": "年份，默认 2026", "en": "year (default: 2026)"},
    "cli.month": {"zh": "月份 1-12，默认 10", "en": "month 1-12 (default: 10)"},
    "cli.social_rate": {"zh": "社保个人比例，如 0.105", "en": "social insurance rate, e.g. 0.105"},
    "cli.fund_rate": {"zh": "公积金比例，如 0.07", "en": "housing fund rate, e.g. 0.07"},
    "cli.start_date": {"zh": "项目起始日 YYYY-MM-DD，用于甘特图", "en": "project start date YYYY-MM-DD (Gantt)"},
    "cli.config": {"zh": "JSON 配置文件，批量指定以上参数（命令行参数优先）", "en": "JSON config file for the options above (CLI flags win)"},
    "cli.spec": {
        "zh": "按 YAML/JSON 模板定义生成自定义模板（见 examples/specs/）；不指定 -t 时只生成这些",
        "en": "build custom templates from YAML/JSON specs (see examples/specs/); without -t, only these are built",
    },
    "cli.metavar.template": {"zh": "模板", "en": "TEMPLATE"},
    "cli.metavar.file": {"zh": "文件", "en": "FILE"},
    # 模板描述（--list 用）
    "desc.kaoqin": {
        "zh": "下拉标记 √/迟/假/旷/休，自动统计出勤率，周末自动标红",
        "en": "Mark √ / late / leave / absent per day; attendance rate computed, weekends red",
    },
    "desc.gongzi": {
        "zh": "社保公积金按比例自动扣，个税按七级累进简化计算，一键生成可打印工资条",
        "en": "Social insurance and housing fund by rate, progressive income tax, printable payslips",
    },
    "desc.gantt": {
        "zh": "改开始日期和工期，色块自动生成；进度条、今日红线、延期自动提示",
        "en": "Set start and duration; bars, today marker and overdue flag appear on their own",
    },
    "desc.jizhang": {
        "zh": "分类下拉记账，月度汇总、储蓄率、支出构成饼图自动出",
        "en": "Category dropdowns, monthly totals, savings rate, spending pie chart",
    },
    "desc.kucun": {
        "zh": "登记出入库，当前库存自动汇总，低于安全库存自动标红",
        "en": "Log stock in and out, running balance, low-stock rows turn red",
    },
    # 生成结果里的固定文案
    "render.total": {"zh": "合计", "en": "Total"},
    # 运行提示
    "list.header": {"zh": "可用模板：", "en": "Available templates:"},
    "msg.generated": {"zh": "已生成 {path}", "en": "Generated {path}"},
    "msg.done": {"zh": "完成：{count} 个文件 → {dir}", "en": "Done: {count} file(s) → {dir}"},
    "msg.write_failed": {"zh": "写入失败：{error}", "en": "Write failed: {error}"},
    # 参数与注册表
    "err.unknown_template": {
        "zh": "未知模板：{token}（用 --list 查看可用模板）",
        "en": "Unknown template: {token} (run --list to see valid ones)",
    },
    "err.no_template": {"zh": "没有选中任何模板", "en": "No template selected"},
    "err.month": {"zh": "月份必须在 1-12 之间，收到 {value}", "en": "Month must be 1-12, got {value}"},
    "err.social_rate": {
        "zh": "社保比例必须在 0-1 之间，收到 {value}",
        "en": "Social insurance rate must be 0-1, got {value}",
    },
    "err.fund_rate": {"zh": "公积金比例必须在 0-1 之间，收到 {value}", "en": "Housing fund rate must be 0-1, got {value}"},
    "err.empty_staff": {"zh": "员工名单不能为空", "en": "Staff list cannot be empty"},
    "err.empty_departments": {"zh": "部门列表不能为空", "en": "Department list cannot be empty"},
    "err.start_date_format": {
        "zh": "start_date 需要 YYYY-MM-DD 格式：{value}",
        "en": "start_date must be YYYY-MM-DD: {value}",
    },
    "err.unknown_option": {"zh": "未知参数：{names}", "en": "Unknown option(s): {names}"},
    "err.config_read": {"zh": "读取配置文件失败：{path}（{error}）", "en": "Cannot read config: {path} ({error})"},
    "err.config_json": {"zh": "配置文件不是合法 JSON：{path}（{error}）", "en": "Invalid JSON in config: {path} ({error})"},
    "err.config_object": {"zh": "配置文件内容需要是一个 JSON 对象：{path}", "en": "Config must be a JSON object: {path}"},
    "err.staff_file_read": {"zh": "读取名单文件失败：{path}（{error}）", "en": "Cannot read staff file: {path} ({error})"},
    "err.staff_file_empty": {"zh": "名单文件里没有有效姓名：{path}", "en": "No valid names in staff file: {path}"},
    # 模板定义
    "err.spec_read": {"zh": "读取模板定义失败：{path}（{error}）", "en": "Cannot read template spec: {path} ({error})"},
    "err.spec_json": {"zh": "模板定义不是合法 JSON：{path}（{error}）", "en": "Invalid JSON in template spec: {path} ({error})"},
    "err.spec_yaml": {"zh": "模板定义不是合法 YAML：{path}（{error}）", "en": "Invalid YAML in template spec: {path} ({error})"},
    "err.spec_needs_yaml": {
        "zh": "读取 YAML 模板定义需要 PyYAML：pip install pyyaml（{path}）",
        "en": "PyYAML is required to read YAML specs: pip install pyyaml ({path})",
    },
    "err.spec_object": {
        "zh": "模板定义内容需要是一个映射（键值对）：{path}",
        "en": "Template spec must be a key-value mapping: {path}",
    },
    "err.unknown_field": {"zh": "{where}里有未知字段：{names}", "en": "Unknown field(s) in {where}: {names}"},
    "where.spec": {"zh": "模板定义", "en": "template spec"},
    "where.sheet": {"zh": "工作表定义", "en": "sheet"},
    "where.column": {"zh": "列定义", "en": "column"},
    "where.conditional": {"zh": "条件格式", "en": "conditional rule"},
    "err.column_type": {
        "zh": "列「{label}」的 type 只能是 {types}，收到 {value}",
        "en": "Column '{label}': type must be one of {types}, got {value}",
    },
    "err.column_label": {"zh": "每一列必须有 label", "en": "Every column needs a label"},
    "err.sheet_name": {"zh": "每个工作表必须有 name", "en": "Every sheet needs a name"},
    "err.sheet_columns": {"zh": "工作表「{name}」至少要有一列", "en": "Sheet '{name}' needs at least one column"},
    "err.sheet_rows": {"zh": "工作表「{name}」的 rows 至少为 1", "en": "Sheet '{name}': rows must be at least 1"},
    "err.conditional_fields": {
        "zh": "条件格式必须同时有 range 和 when",
        "en": "A conditional rule needs both range and when",
    },
    "err.filename_required": {"zh": "模板必须有 filename", "en": "Template spec needs a filename"},
    "err.filename_suffix": {"zh": "filename 必须以 .xlsx 结尾：{value}", "en": "filename must end with .xlsx: {value}"},
    "err.sheets_required": {
        "zh": "模板「{filename}」至少要有一个工作表",
        "en": "Template '{filename}' needs at least one sheet",
    },
}

_EPILOG = {
    "zh": (
        "示例：\n"
        "  python build.py                  生成全部内置模板到 dist/\n"
        "  python build.py -o ./我的模板     指定输出目录\n"
        "  python build.py -t 1,3            只生成第 1、3 套\n"
        "  python build.py --list            查看可用模板\n"
        "\n"
        "带上自己的公司和名单：\n"
        "  python build.py -c 某某科技 --staff 张三,李四,王五 --departments 销售部,技术部\n"
        "  python build.py --staff-file names.txt --year 2026 --month 11\n"
        "  python build.py --config examples/params.json\n"
        "\n"
        "用 YAML 定义自己的模板：\n"
        "  python build.py --spec examples/specs/差旅报销单.yaml\n"
        "  python build.py --spec my.yaml -c 某某科技 -o ./out\n"
        "\n"
        "英文界面：\n"
        "  python build.py --lang en --help\n"
    ),
    "en": (
        "Examples:\n"
        "  python build.py                      build every built-in template into dist/\n"
        "  python build.py -o ./my-templates    choose the output directory\n"
        "  python build.py -t 1,3               build only #1 and #3\n"
        "  python build.py --list               show available templates\n"
        "\n"
        "With your own company and staff:\n"
        "  python build.py -c Acme --staff Alice,Bob,Carol --departments Sales,Engineering\n"
        "  python build.py --staff-file names.txt --year 2026 --month 11\n"
        "  python build.py --config examples/params.json\n"
        "\n"
        "Define your own template in YAML:\n"
        "  python build.py --spec examples/specs/expense-report.yaml\n"
        "  python build.py --spec my.yaml -c Acme -o ./out\n"
        "\n"
        "Chinese interface:\n"
        "  python build.py --lang zh --help\n"
    ),
}


def detect_lang() -> str:
    """按环境变量猜一次语言，猜不出就用默认中文。"""
    for value in (os.environ.get("OFFICE_TEMPLATES_LANG", ""), *(os.environ.get(v, "") for v in ("LANG", "LC_ALL", "LC_MESSAGES"))):
        low = value.lower()
        if low.startswith("zh"):
            return "zh"
        if low.startswith("en"):
            return "en"
    return DEFAULT_LANG


_LANG = detect_lang()


def set_lang(lang: str) -> None:
    """设置当前语言（zh / en）。"""
    global _LANG
    if lang in SUPPORTED_LANGS:
        _LANG = lang


def get_lang() -> str:
    return _LANG


def t(key: str, **kwargs) -> str:
    """取一条文案并按 kwargs 格式化，缺 key 时原样返回 key。"""
    entry = MESSAGES.get(key)
    if entry is None:
        return key
    text = entry.get(_LANG, entry.get(DEFAULT_LANG, key))
    return text.format(**kwargs) if kwargs else text


def epilog(lang: str | None = None) -> str:
    return _EPILOG.get(lang or _LANG, _EPILOG[DEFAULT_LANG])
