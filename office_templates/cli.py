"""命令行入口：python build.py [选项]。"""

from __future__ import annotations

import argparse
import json
import logging
import sys
from pathlib import Path
from typing import Any

from . import __version__
from .i18n import SUPPORTED_LANGS, detect_lang, epilog, set_lang, t
from .options import Options, read_names, split_list
from .registry import TEMPLATES, build_templates, resolve
from .render import render_spec
from .spec import load_spec

logger = logging.getLogger("office_templates")

DEFAULT_OUT_DIR = "dist"

# 命令行参数 → Options 字段
OPTION_FIELDS = ("company", "year", "month", "social_rate", "fund_rate", "start_date")


def build_parser(lang: str) -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="build",
        description=t("cli.description"),
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=epilog(lang),
    )
    parser.add_argument("-o", "--out-dir", default=DEFAULT_OUT_DIR, help=t("cli.out_dir"))
    parser.add_argument(
        "-t",
        "--only",
        nargs="+",
        metavar=t("cli.metavar.template"),
        help=t("cli.only"),
    )
    parser.add_argument("-l", "--list", action="store_true", help=t("cli.list"))
    parser.add_argument("-q", "--quiet", action="store_true", help=t("cli.quiet"))
    parser.add_argument("-V", "--version", action="version", version=f"%(prog)s {__version__}")
    parser.add_argument("--lang", choices=SUPPORTED_LANGS, help=t("cli.lang"))

    params = parser.add_argument_group(t("cli.group.params"))
    params.add_argument("-c", "--company", help=t("cli.company"))
    params.add_argument("--staff", help=t("cli.staff"))
    params.add_argument("--staff-file", help=t("cli.staff_file"))
    params.add_argument("--departments", help=t("cli.departments"))
    params.add_argument("--year", type=int, help=t("cli.year"))
    params.add_argument("--month", type=int, help=t("cli.month"))
    params.add_argument("--social-rate", type=float, help=t("cli.social_rate"))
    params.add_argument("--fund-rate", type=float, help=t("cli.fund_rate"))
    params.add_argument("--start-date", help=t("cli.start_date"))
    params.add_argument("--config", help=t("cli.config"))
    params.add_argument(
        "--spec",
        nargs="+",
        metavar=t("cli.metavar.file"),
        help=t("cli.spec"),
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
            raise ValueError(t("err.config_read", path=args.config, error=exc)) from exc
        except json.JSONDecodeError as exc:
            raise ValueError(t("err.config_json", path=args.config, error=exc)) from exc
        if not isinstance(loaded, dict):
            raise ValueError(t("err.config_object", path=args.config))
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
    width = max(len(tpl.key) for tpl in TEMPLATES)
    print(t("list.header"))
    for i, tpl in enumerate(TEMPLATES, 1):
        print(f"  {i}. {tpl.key:<{width}}  {tpl.filename}  — {t(f'desc.{tpl.key}')}")


def main(argv: list[str] | None = None) -> int:
    # 先只解析 --lang，好让 --help 也用对应语言；其余参数交给正式解析器
    pre = argparse.ArgumentParser(add_help=False)
    pre.add_argument("--lang", choices=SUPPORTED_LANGS)
    known, _ = pre.parse_known_args(argv)
    lang = known.lang or detect_lang()
    set_lang(lang)

    args = build_parser(lang).parse_args(argv)
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
        logger.error(t("msg.write_failed", error=exc))
        return 1

    for spec in specs:
        try:
            wb = render_spec(spec, options)
            path = out_dir / spec.filename
            path.parent.mkdir(parents=True, exist_ok=True)
            wb.save(path)
        except OSError as exc:
            logger.error(t("msg.write_failed", error=exc))
            return 1
        paths.append(path)
        logger.info(t("msg.generated", path=path))

    logger.info(t("msg.done", count=len(paths), dir=out_dir.resolve()))
    return 0


if __name__ == "__main__":
    sys.exit(main())
