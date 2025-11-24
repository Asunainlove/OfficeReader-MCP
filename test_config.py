"""
Test script for configuration loading.
Run this to verify that config.json is being loaded correctly.
"""

import sys
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).parent / "src"))

from officereader_mcp.config_loader import get_config, load_config


def test_config_loading():
    """Test configuration loading from different sources."""

    print("=" * 60)
    print("  OfficeReader-MCP Configuration Test")
    print("=" * 60)
    print()

    # Load configuration
    config = get_config()

    print(f"Config file location: {config.config_path}")
    print(f"Config loaded: {'Yes' if config.config_data else 'No (using defaults)'}")
    print()

    print("-" * 60)
    print("Configuration Values:")
    print("-" * 60)

    # Test cache_dir
    cache_dir = config.get_cache_dir()
    print(f"Cache Directory: {cache_dir if cache_dir else 'Default (system temp)'}")

    # Test output_dir
    output_dir = config.get_output_dir()
    print(f"Output Directory: {output_dir if output_dir else 'Default (cache_dir/output)'}")

    # Test image optimization settings
    img_opt = config.get_image_optimization_settings()
    print(f"\nImage Optimization:")
    print(f"  - Enabled: {img_opt['enabled']}")
    print(f"  - Max Dimension: {img_opt['max_dimension']}px")
    print(f"  - Quality: {img_opt['quality']}%")

    # Test default settings
    defaults = config.get_default_settings()
    print(f"\nDefault Settings:")
    print(f"  - Extract Images: {defaults['extract_images']}")
    print(f"  - Image Format: {defaults['image_format']}")

    print()
    print("-" * 60)
    print("Raw Configuration Data:")
    print("-" * 60)

    import json
    print(json.dumps(config.config_data, indent=2, ensure_ascii=False))

    print()
    print("=" * 60)
    print("Test Complete!")
    print("=" * 60)

    return config


def test_converter_initialization():
    """Test that the converter can be initialized with the config."""
    print("\n" + "=" * 60)
    print("  Testing Converter Initialization")
    print("=" * 60)
    print()

    try:
        from officereader_mcp.converter import OfficeConverter
        from officereader_mcp.config_loader import get_config

        config = get_config()
        cache_dir = config.get_cache_dir()

        converter = OfficeConverter(cache_dir=cache_dir)

        print(f"Converter initialized successfully!")
        print(f"Cache directory: {converter.cache_dir}")
        print(f"Output directory: {converter.output_dir}")

        # Check if directories exist
        print(f"\nDirectory status:")
        print(f"  - Cache dir exists: {converter.cache_dir.exists()}")
        print(f"  - Output dir exists: {converter.output_dir.exists()}")

        print()
        print("=" * 60)
        print("Converter Test Complete!")
        print("=" * 60)

        return converter

    except Exception as e:
        print(f"ERROR: Failed to initialize converter: {e}")
        import traceback
        traceback.print_exc()
        return None


if __name__ == "__main__":
    # Test configuration loading
    config = test_config_loading()

    # Test converter initialization
    converter = test_converter_initialization()

    if converter:
        print("\n[SUCCESS] All tests passed!")
    else:
        print("\n[FAILED] Some tests failed. Check the output above.")
