"""
Test script for all supported Office formats.
"""

import sys
sys.path.insert(0, 'src')

from officereader_mcp import OfficeConverter, ImageOptimizer, SUPPORTED_FORMATS

def test_imports():
    """Test that all modules import correctly."""
    print("=" * 50)
    print("Testing imports...")
    print("=" * 50)

    from officereader_mcp.converter import DocxConverter, OfficeConverter, get_file_type
    from officereader_mcp.image_optimizer import ImageOptimizer
    from officereader_mcp.excel_converter import ExcelConverter
    from officereader_mcp.pptx_converter import PowerPointConverter

    print("All imports successful!")
    print(f"Supported formats: {SUPPORTED_FORMATS}")
    print()

def test_image_optimizer():
    """Test image optimizer functionality."""
    print("=" * 50)
    print("Testing ImageOptimizer...")
    print("=" * 50)

    from PIL import Image
    import io

    optimizer = ImageOptimizer(
        max_width=800,
        max_height=600,
        prefer_webp=True
    )

    # Create a test image
    test_img = Image.new('RGB', (1920, 1080), color='red')

    # Optimize it
    optimized_bytes, ext = optimizer.optimize_image(test_img, "test_image")

    print(f"Original size: 1920x1080")
    print(f"Optimized format: {ext}")
    print(f"Optimized size: {len(optimized_bytes)} bytes")

    # Check resized image
    optimized_img = Image.open(io.BytesIO(optimized_bytes))
    print(f"Optimized dimensions: {optimized_img.size}")
    print()

def test_office_converter():
    """Test OfficeConverter initialization."""
    print("=" * 50)
    print("Testing OfficeConverter initialization...")
    print("=" * 50)

    converter = OfficeConverter()

    print(f"Cache directory: {converter.cache_dir}")
    print(f"Output directory: {converter.output_dir}")
    print(f"Image optimizer: {converter.image_optimizer}")
    print(f"Supported formats: {converter.get_supported_formats()}")
    print()

def test_get_file_type():
    """Test file type detection."""
    print("=" * 50)
    print("Testing get_file_type...")
    print("=" * 50)

    from officereader_mcp.converter import get_file_type
    from pathlib import Path

    test_files = [
        "document.docx",
        "document.doc",
        "spreadsheet.xlsx",
        "spreadsheet.xls",
        "presentation.pptx",
        "presentation.ppt",
        "unknown.pdf",
        "unknown.txt",
    ]

    for filename in test_files:
        file_type = get_file_type(Path(filename))
        print(f"{filename} -> {file_type}")
    print()

if __name__ == "__main__":
    print("\n" + "=" * 50)
    print("OfficeReader-MCP v2.0 Test Suite")
    print("=" * 50 + "\n")

    test_imports()
    test_image_optimizer()
    test_office_converter()
    test_get_file_type()

    print("=" * 50)
    print("All tests passed!")
    print("=" * 50)
