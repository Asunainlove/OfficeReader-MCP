"""
OfficeReader-MCP: Convert Office documents to Markdown via MCP protocol.

Supported formats:
- Word: .docx, .doc
- Excel: .xlsx, .xls
- PowerPoint: .pptx, .ppt
"""

__version__ = "2.0.0"
__author__ = "Asunainlove"

from .converter import DocxConverter, OfficeConverter, SUPPORTED_FORMATS, get_file_type
from .image_optimizer import ImageOptimizer
from .server import main

__all__ = [
    "DocxConverter",
    "OfficeConverter",
    "ImageOptimizer",
    "SUPPORTED_FORMATS",
    "get_file_type",
    "main",
]
