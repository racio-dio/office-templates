#!/usr/bin/env python3
"""生成 5 套「填数即自动计算」的中文办公 Excel 模板。

用法：
    python build.py                  # 全部生成到 dist/
    python build.py --help           # 查看全部选项
"""

from office_templates.cli import main

if __name__ == "__main__":
    raise SystemExit(main())
