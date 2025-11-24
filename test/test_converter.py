"""
Simple test script to verify the converter works.
Run this after installation to test the conversion functionality.
"""

import sys
import os

# Add src to path for testing
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src"))

from officereader_mcp.converter import DocxConverter


def test_converter():
    """Test basic converter functionality."""
    print("Testing OfficeReader-MCP Converter")
    print("=" * 40)

    # Initialize converter with local cache
    cache_dir = os.path.join(os.path.dirname(__file__), "cache")
    converter = DocxConverter(cache_dir=cache_dir)

    print(f"Cache directory: {converter.cache_dir}")
    print(f"Output directory: {converter.output_dir}")
    print()

    # Test cache info
    cache_info = converter.get_cache_info()
    print(f"Existing conversions: {cache_info['total_conversions']}")
    print()

    # Test with a real file if provided
    if len(sys.argv) > 1:
        test_file = sys.argv[1]
        print(f"Converting: {test_file}")
        print("-" * 40)

        try:
            result = converter.convert_file(
                file_path=test_file,
                extract_images=True,
                image_format="file"
            )

            print(f"Output path: {result['output_path']}")
            print(f"Images extracted: {len(result['images'])}")
            print(f"Markdown length: {len(result['markdown'])} characters")
            print()
            print("Preview (first 500 chars):")
            print("-" * 40)
            print(result['markdown'][:500])
            print()
            print("Conversion successful!")

        except Exception as e:
            print(f"Error: {e}")
            return False
    else:
        print("No test file provided.")
        print("Usage: python test_converter.py <path_to_docx_file>")
        print()
        print("Converter initialized successfully!")

    return True


if __name__ == "__main__":
    success = test_converter()
    sys.exit(0 if success else 1)
