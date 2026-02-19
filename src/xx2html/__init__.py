from importlib.metadata import PackageNotFoundError, version

from xx2html.core import apply_openpyxl_patches, create_xlsx_transform

try:
    __version__ = version("xx2html")
except PackageNotFoundError:
    __version__ = "0.0.0"

__all__ = ["__version__", "apply_openpyxl_patches", "create_xlsx_transform"]
