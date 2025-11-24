# OfficeReader-MCP

A Model Context Protocol (MCP) server that converts Microsoft Office documents (Word, Excel, PowerPoint) to Markdown format with intelligent image extraction and optimization.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)

## Features

- **Multi-Format Support**: Word (.docx, .doc), Excel (.xlsx, .xls), PowerPoint (.pptx, .ppt)
- **Intelligent Image Processing**: Automatic extraction and optimization with WebP compression
- **Format Preservation**: Maintains document structure including headings, tables, lists, and formatting
- **Metadata Extraction**: Access document properties (author, title, creation date, etc.)
- **Efficient Caching**: Smart caching system for quick reuse of converted documents
- **Cross-Platform**: Works on Windows, macOS, and Linux

## Supported Formats

| Format | Extensions | Features |
|--------|------------|----------|
| **Word** | `.docx`, `.doc` | Text formatting, headings, lists, tables, images |
| **Excel** | `.xlsx`, `.xls` | Multi-sheet support, tables, charts, embedded images |
| **PowerPoint** | `.pptx`, `.ppt` | Slides, text boxes, images, speaker notes, tables |

## Installation

### Prerequisites

- Python 3.10 or higher
- Claude Desktop or Claude Code

### Step 1: Install the Package

```bash
# Clone the repository
git clone https://github.com/Asunainlove/office-reader-mcp.git
cd office-reader-mcp

# Install in editable mode
pip install -e .
```

### Step 2: Configure Claude

#### For Claude Desktop

Add to your Claude Desktop config file:

**Windows**: `%APPDATA%\Claude\claude_desktop_config.json`
**macOS/Linux**: `~/.config/Claude/claude_desktop_config.json`

```json
{
  "mcpServers": {
    "officereader": {
      "command": "python",
      "args": ["-m", "officereader_mcp.server"],
      "env": {
        "OFFICEREADER_CACHE_DIR": "/path/to/cache"
      }
    }
  }
}
```

#### For Claude Code

Add to your Claude Code settings:

**Windows**: `%LOCALAPPDATA%\claude-code\settings.json`
**macOS/Linux**: `~/.config/claude-code/settings.json`

```json
{
  "mcpServers": {
    "officereader": {
      "command": "python",
      "args": ["-m", "officereader_mcp.server"],
      "env": {
        "OFFICEREADER_CACHE_DIR": "/path/to/cache"
      }
    }
  }
}
```

### Step 3: Restart Claude

Restart Claude Desktop or Claude Code to load the MCP server.

## Quick Start

After installation, you can use OfficeReader-MCP directly in your conversations with Claude:

```
Convert my Excel file at D:\Reports\sales_2024.xlsx to markdown
```

```
Extract text and images from D:\Presentations\keynote.pptx
```

```
Get metadata from my document at C:\Documents\report.docx
```

## Available Tools

### 1. `convert_document`

Convert any supported Office document to Markdown format.

**Parameters:**
- `file_path` (required): Absolute path to the document
- `extract_images` (optional, default: true): Extract embedded images
- `image_format` (optional, default: "file"): How to handle images
  - `"file"`: Save images to disk (recommended)
  - `"base64"`: Embed images as base64 in markdown
  - `"both"`: Both save and embed
- `output_name` (optional): Custom name for output files

**Example:**
```
Convert D:\Documents\report.xlsx with images
```

### 2. `read_converted_markdown`

Read the full content of a previously converted markdown file.

**Parameters:**
- `markdown_path` (required): Path to the markdown file

**Example:**
```
Read the markdown at D:\cache\output\report_abc12345\report_abc12345.md
```

### 3. `list_conversions`

List all cached document conversions with details.

**Example:**
```
List all converted documents
```

### 4. `clear_cache`

Clear all cached conversions to free up disk space.

**Example:**
```
Clear the document cache
```

### 5. `get_document_metadata`

Extract metadata from a document without full conversion (faster).

**Parameters:**
- `file_path` (required): Path to the document

**Example:**
```
Get metadata from D:\Documents\presentation.pptx
```

### 6. `get_supported_formats`

Get list of all supported file formats and extensions.

**Example:**
```
What file formats does officereader support?
```

## Output Structure

Converted documents are organized in the cache directory:

```
cache/
└── output/
    └── document_name_abc12345/
        ├── document_name_abc12345.md    # Converted markdown
        └── images/
            ├── image_001.webp           # Optimized images
            ├── slide2_image_002.webp
            └── excel_image_003.webp
```

## Image Optimization

Images are automatically optimized to reduce file size while maintaining quality:

- **Max Dimensions**: 1920×1080 pixels (configurable)
- **Format**: WebP (preferred) or PNG/JPEG fallback
- **Quality**: 80% for photos, 85% for JPEG, lossless PNG for graphics with transparency
- **Typical Compression**: 50-80% size reduction
- **Smart Detection**: Automatically distinguishes between photos and graphics

## Technical Details

### Architecture

```
OfficeReader-MCP/
├── src/officereader_mcp/
│   ├── server.py              # MCP server implementation
│   ├── converter.py           # Word converter (DocxConverter, OfficeConverter)
│   ├── excel_converter.py    # Excel to Markdown converter
│   ├── pptx_converter.py     # PowerPoint to Markdown converter
│   ├── image_optimizer.py    # Image compression utility
│   └── __init__.py           # Package initialization
├── test/
│   ├── test_converter.py     # Basic functionality tests
│   └── test_all_formats.py   # Comprehensive test suite
├── pyproject.toml            # Project configuration
└── README.md                 # Documentation
```

### Dependencies

| Package | Version | Purpose |
|---------|---------|---------|
| `mcp` | >=1.0.0 | Model Context Protocol SDK |
| `python-docx` | >=1.1.0 | DOCX file parsing and manipulation |
| `mammoth` | >=1.6.0 | DOC/DOCX to HTML conversion (fallback) |
| `Pillow` | >=10.0.0 | Image processing and optimization |
| `markdownify` | >=0.11.0 | HTML to Markdown conversion |
| `openpyxl` | >=3.1.0 | Excel file parsing |
| `python-pptx` | >=0.6.21 | PowerPoint file parsing |

All dependencies are automatically installed when you run `pip install -e .`

## Testing

### Run Tests

```bash
# Basic converter test
python test/test_converter.py

# Comprehensive test suite for all formats
python test/test_all_formats.py

# Test with a specific document
python test/test_converter.py path/to/your/document.docx
```

### Test Coverage

The test suite verifies:
- Module imports and initialization
- Converter functionality for all formats
- Image extraction and optimization
- File type detection
- Cache management
- Metadata extraction

## Configuration

OfficeReader-MCP supports multiple configuration methods to customize cache locations and behavior.

### Quick Configuration (Recommended)

1. Copy the example config file:
   ```bash
   cp config.example.json config.json
   ```

2. Edit `config.json` to set your cache directory:
   ```json
   {
     "cache_dir": "D:/MyDocuments/OfficeReaderCache",
     "image_optimization": {
       "enabled": true,
       "max_dimension": 1920,
       "quality": 80
     }
   }
   ```

3. The config file will be automatically loaded on startup.

For detailed configuration options, see [CONFIG.md](CONFIG.md).

### Environment Variables

| Variable | Description | Default |
|----------|-------------|---------|
| `OFFICEREADER_CACHE_DIR` | Directory for cached conversions | System temp directory |

Example usage:
```bash
# Set custom cache directory
export OFFICEREADER_CACHE_DIR=/path/to/custom/cache

# Or in Windows
set OFFICEREADER_CACHE_DIR=C:\path\to\custom\cache
```

**Note**: Environment variables take priority over config file settings.

## Usage Examples

### Converting Excel with Multiple Sheets

```
User: Convert my Excel file at D:\Reports\Q4_sales.xlsx

Claude: I'll convert that Excel file. Each sheet will be converted to a separate
        section in the markdown with properly formatted tables...

[Output includes all sheets as markdown tables with preserved formatting]
```

### Extracting PowerPoint Content

```
User: Extract all text and images from D:\Presentations\product_launch.pptx

Claude: Converting the PowerPoint presentation. I'll extract text from each slide,
        including speaker notes, along with all embedded images...

[Output includes slide-by-slide breakdown with images and notes]
```

### Batch Processing

```
User: Convert all Office documents in D:\Documents\

Claude: I'll convert each document and cache the results for quick access...

[Processes all supported files and provides summary]
```

## Troubleshooting

### "Module not found" Error

```bash
# Reinstall the package
pip install -e .
```

### Configuration Not Loading

1. Verify the config file location is correct
2. Check JSON syntax is valid (use a JSON validator)
3. Restart Claude Desktop or Claude Code completely
4. Check logs for error messages

### Images Not Extracting

Possible causes:
- Document contains linked images (not embedded)
- Insufficient write permissions for cache directory
- Image format not supported by the document library

Solution:
```bash
# Verify cache directory is writable
ls -la /path/to/cache  # Unix/Mac
dir /path/to/cache     # Windows

# Check if images are embedded
# Use convert_document with extract_images=true explicitly
```

### Encoding Issues

The converter uses UTF-8 encoding throughout. If you see garbled text:
- Check the source document encoding
- Ensure your terminal/console supports UTF-8
- Try converting with different system locale settings

## Changelog

### v2.0.0 (2024-11)

**Major Features:**
- Added Excel (.xlsx, .xls) support with multi-sheet conversion
- Added PowerPoint (.pptx, .ppt) support with slide extraction
- Implemented intelligent image optimization with WebP compression
- Added unified OfficeConverter interface for all document types
- Enhanced metadata extraction for all formats

**Improvements:**
- Smart caching system with hash-based file identification
- Lazy-loading of format-specific converters for better performance
- Better error handling and validation
- Comprehensive test suite for all formats

**Tools:**
- Added `get_supported_formats` tool
- Enhanced `get_document_metadata` for all formats
- Improved `list_conversions` with detailed cache information

### v1.0.0 (2024-09)

- Initial release
- Word document (.docx, .doc) conversion
- Basic image extraction
- MCP server implementation

## Contributing

Contributions are welcome! Here's how you can help:

1. **Report Bugs**: Open an issue with details and steps to reproduce
2. **Suggest Features**: Describe your idea and use case
3. **Submit Pull Requests**:
   - Fork the repository
   - Create a feature branch (`git checkout -b feature/amazing-feature`)
   - Commit your changes (`git commit -m 'Add amazing feature'`)
   - Push to your branch (`git push origin feature/amazing-feature`)
   - Open a Pull Request

### Development Setup

```bash
# Clone and install with dev dependencies
git clone https://github.com/Asunainlove/office-reader-mcp.git
cd office-reader-mcp
pip install -e ".[dev]"

# Run tests
python test/test_all_formats.py

# Run linting (if configured)
black src/
ruff check src/
```

## License

MIT License - see [LICENSE](LICENSE) file for details.

## Author

**Asunainlove**

- GitHub: [@Asunainlove](https://github.com/Asunainlove)
- Repository: [office-reader-mcp](https://github.com/Asunainlove/office-reader-mcp)
- Issues: [Report a bug](https://github.com/Asunainlove/office-reader-mcp/issues)

## Acknowledgments

This project uses the following open-source libraries:
- [Model Context Protocol (MCP)](https://modelcontextprotocol.io/) by Anthropic
- [python-docx](https://python-docx.readthedocs.io/) for Word processing
- [openpyxl](https://openpyxl.readthedocs.io/) for Excel processing
- [python-pptx](https://python-pptx.readthedocs.io/) for PowerPoint processing
- [Pillow](https://pillow.readthedocs.io/) for image processing

## Support

If you find this project helpful, please:
- ⭐ Star the repository
- 🐛 Report bugs and issues
- 💡 Suggest new features
- 🔀 Contribute code improvements

---

**Happy converting!** 🚀
