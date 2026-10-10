"""命令行入口：python build.py [选项]。"""

from __future__ import annotations

import argparse
import json
import logging
import sys
from pathlib import Path
from typing import Any

from . import __version__
from .options import Options, read_names, split_list
from .registry import TEMPLATES, build_templates, resolve
from .render import render_spec
from .spec import load_spec

logger = logging.getLogger("office_templates")

DEFAULT_OUT_DIR = "dist"

# 命令行参数 → Options 字段
OPTION_FIELDS = ("company", "year", "month", "social_rate", "fund_rate", "start_date")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="build",
        description="生成 5 套「填数即自动计算」的中文办公 Excel 模板",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
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
        ),
    )
    parser.add_argument("-o", "--out-dir", default=DEFAULT_OUT_DIR, help=f"输出目录，默认 {DEFAULT_OUT_DIR}/")
    parser.add_argument(
        "-t",
        "--only",
        nargs="+",
        metavar="模板",
        help="只生成指定内置模板，编号或名称均可，空格或逗号分隔",
    )
    parser.add_argument("-l", "--list", action="store_true", help="列出所有可用模板后退出")
    parser.add_argument("-q", "--quiet", action="store_true", help="只输出错误信息")
    parser.add_argument("-V", "--version", action="version", version=f"%(prog)s {__version__}")

    params = parser.add_argument_group("模板参数（影响生成的模板内容）")
    params.add_argument("-c", "--company", help="公司名，会加在每套模板的标题前")
    params.add_argument("--staff", help="员工名单，逗号分隔，如：张三,李四,王五")
    params.add_argument("--staff-file", help="员工名单文件，一行一个姓名，# 开头为注释")
    params.add_argument("--departments", help="部门列表，逗号分隔，如：销售部,技术部,财务部")
    params.add_argument("--year", type=int, help="年份，默认 2026")
    params.add_argument("--month", type=int, help="月份 1-12，默认 10")
    params.add_argument("--social-rate", type=float, help="社保个人比例，如 0.105")
    params.add_argument("--fund-rate", type=float, help="公积金比例，如 0.07")
    params.add_argument("--start-date", help="项目起始日 YYYY-MM-DD，用于甘特图")
    params.add_argument("--config", help="JSON 配置文件，批量指定以上参数（命令行参数优先）")
    params.add_argument(
        "--spec",
        nargs="+",
        metavar="文件",
        help="按 YAML/JSON 模板定义生成自定义模板（见 examples/specs/）；不指定 -t 时只生成这些",
    )
    return parser


def collect_options(args: argparse.Namespace) -> Options:
    """按「命令行 > 配置文件 > 默认值」的优先级组装参数。"""
    data: dict[str, Any] = {}

    if args.config:
        path = Path(args.config)
        try:
            loaded = json.loads(path.read_text(encoding="utf-8"))
        except OSError as exc:
            raise ValueError(f"读取配置文件失败：{args.config}（{exc}）") from exc
        except json.JSONDecodeError as exc:
            raise ValueError(f"配置文件不是合法 JSON：{args.config}（{exc}）") from exc
        if not isinstance(loaded, dict):
            raise ValueError(f"配置文件内容需要是一个 JSON 对象：{args.config}")
        data.update(loaded)

    for field in OPTION_FIELDS:
        value = getattr(args, field)
        if value is not None:
            data[field] = value

    names: list[str] = []
    if args.staff:
        names.extend(split_list(args.staff) or ())
    if args.staff_file:
        names.extend(read_names(args.staff_file))
    if names:
        data["staff"] = names

    if args.departments:
        depts = split_list(args.departments)
        if depts:
            data["departments"] = depts

    return Options.from_dict(data)


def print_list() -> None:
    width = max(len(t.key) for t in TEMPLATES)
    print("可用模板：")
    for i, t in enumerate(TEMPLATES, 1):
        print(f"  {i}. {t.key:<{width}}  {t.filename}  — {t.desc}")


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    logging.basicConfig(
        level=logging.WARNING if args.quiet else logging.INFO,
        format="%(message)s",
        stream=sys.stdout,
    )

    if args.list:
        print_list()
        return 0

    try:
        # 给了 --spec 又没指定 -t 时，只生成自定义模板
        selected = resolve(args.only) if (args.only or not args.spec) else []
        options = collect_options(args)
        specs = [load_spec(p) for p in (args.spec or [])]
    except ValueError as exc:
        logger.error("%s", exc)
        return 2

    out_dir = Path(args.out_dir)
    try:
        paths = build_templates(selected, out_dir, options)
    except OSError as exc:
        logger.error("写入失败：%s", exc)
        return 1

    for spec in specs:
        try:
            wb = render_spec(spec, options)
            path = out_dir / spec.filename
            path.parent.mkdir(parents=True, exist_ok=True)
            wb.save(path)
        except OSError as exc:
            logger.error("写入失败：%s", exc)
            return 1
        paths.append(path)
        logger.info("已生成 %s", path)

    logger.info("完成：%d 个文件 → %s", len(paths), out_dir.resolve())
    return 0


if __name__ == "__main__":
    sys.exit(main())
