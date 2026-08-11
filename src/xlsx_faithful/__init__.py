"""xlsx-faithful:往 Excel 模板里写数据,同时保住模板里的一切。"""

from .core import XlsxFaithfulError, list_parts, write_cells

__all__ = ["write_cells", "list_parts", "XlsxFaithfulError"]
__version__ = "0.1.0"
