"""office-templates：职场 Excel 自动化模板生成器。"""

__version__ = "0.4.0"

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
]
