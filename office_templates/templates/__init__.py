"""模板实现模块。

每个模板模块约定提供四个名字：KEY、FILENAME、DESC、build()。
新增模板时在 MODULES 里登记即可，CLI 与 --list 会自动识别。
"""

from . import gantt, gongzi, jizhang, kaoqin, kucun

MODULES = (kaoqin, gongzi, gantt, jizhang, kucun)

__all__ = ["MODULES", "kaoqin", "gongzi", "gantt", "jizhang", "kucun"]
