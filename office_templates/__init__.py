"""office-templates：声明式 Excel 模板生成器（declarative Excel template generator）。"""

__version__ = "0.5.0"

from .i18n import detect_lang, get_lang, set_lang
from .options import Options
from .registry import TEMPLATES, Template, all_templates, build_templates, resolve
from .render import render_spec
from .spec import ColumnSpec, ConditionalSpec, SheetSpec, Spec, load_spec

__all__ = [
    "__version__",
    "Options",
    "TEMPLATES",
    "Template",
    "ColumnSpec",
    "ConditionalSpec",
    "SheetSpec",
    "Spec",
    "load_spec",
    "render_spec",
    "all_templates",
    "build_templates",
    "resolve",
    "set_lang",
    "get_lang",
    "detect_lang",
]
