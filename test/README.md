# OfficeReader-MCP Test Suite

This directory contains test scripts to verify the functionality of OfficeReader-MCP.

## Test Files

### 1. `test_converter.py`
Tests the basic document conversion functionality.

**Usage:**
```bash
# Test converter initialization only
python test/test_converter.py

# Test with an actual document file
python test/test_converter.py path/to/document.docx
```

**What it tests:**
- Converter initialization
- Cache directory setup
- Document conversion (if file provided)
- Image extraction
- Markdown generation

### 2. `test_all_formats.py`
Comprehensive test suite for all components.

**Usage:**
```bash
python test/test_all_formats.py
```

**What it tests:**
- Module imports
- Image optimizer functionality
- OfficeConverter initialization
- File type detection for all supported formats (.docx, .xlsx, .pptx, etc.)

## Running Tests via Console

You can also run tests through the management console:

**Windows:**
```cmd
console.bat
# Select option [10] Run Tests
```

**Linux/macOS:**
```bash
./console.sh
# Select option [10] Run Tests
```

## Test Requirements

All dependencies are installed automatically with the main package:
- `mcp>=1.0.0`
- `python-docx>=1.1.0`
- `mammoth>=1.6.0`
- `Pillow>=10.0.0`
- `markdownify>=0.11.0`
- `openpyxl>=3.1.0`
- `python-pptx>=0.6.21`

## Expected Output

### test_converter.py (without file)
```
Testing OfficeReader-MCP Converter
========================================
Cache directory: /path/to/cache
Output directory: /path/to/cache/output
Existing conversions: 0
No test file provided.
Usage: python test_converter.py <path_to_docx_file>
Converter initialized successfully!
```

### test_all_formats.py
```
==================================================
OfficeReader-MCP v2.0 Test Suite
==================================================

==================================================
Testing imports...
==================================================
All imports successful!
Supported formats: {'word': ['.docx', '.doc'], ...}

==================================================
Testing ImageOptimizer...
==================================================
Original size: 1920x1080
Optimized format: .webp
Optimized size: XXXX bytes
Optimized dimensions: (800, 450)

==================================================
All tests passed!
==================================================
```

## Troubleshooting

### Import Errors
If you get import errors, ensure you're in the project root directory:
```bash
cd /path/to/OfficeReader-MCP
python test/test_converter.py
```

### Module Not Found
Install the package in development mode:
```bash
pip install -e .
```

### Test File Not Found
For `test_converter.py` with a file argument, ensure the file path is correct and the file exists.

## Adding New Tests

To add new tests:
1. Create a new file in this directory: `test_your_feature.py`
2. Follow the existing test structure
3. Use descriptive function names starting with `test_`
4. Add clear print statements for test output
5. Update this README with usage instructions

## Notes

- All test files use English only to avoid encoding issues
- Tests are non-destructive and don't modify source files
- Cache is created in the project directory during tests
- Test files can be run independently or via the console
