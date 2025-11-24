# OfficeReader-MCP

一个模型上下文协议（MCP）服务器，可将 Microsoft Office 文档（Word、Excel、PowerPoint）转换为 Markdown 格式，并具有智能图像提取和优化功能。

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)

## 功能特点

- **多格式支持**：Word (.docx, .doc)、Excel (.xlsx, .xls)、PowerPoint (.pptx, .ppt)
- **智能图像处理**：自动提取和优化图像，支持 WebP 压缩
- **格式保留**：保持文档结构，包括标题、表格、列表和格式
- **元数据提取**：访问文档属性（作者、标题、创建日期等）
- **高效缓存**：智能缓存系统，快速复用已转换的文档
- **跨平台**：支持 Windows、macOS 和 Linux

## 支持的格式

| 格式 | 扩展名 | 功能特性 |
|------|--------|----------|
| **Word** | `.docx`, `.doc` | 文本格式、标题、列表、表格、图像 |
| **Excel** | `.xlsx`, `.xls` | 多工作表支持、表格、图表、嵌入图像 |
| **PowerPoint** | `.pptx`, `.ppt` | 幻灯片、文本框、图像、演讲者备注、表格 |

## 安装

### 前置要求

- Python 3.10 或更高版本
- Claude Desktop 或 Claude Code

### 步骤 1：安装包

```bash
# 克隆仓库
git clone https://github.com/Asunainlove/office-reader-mcp.git
cd office-reader-mcp

# 以可编辑模式安装
pip install -e .
```

### 步骤 2：配置 Claude

#### 配置 Claude Desktop

在 Claude Desktop 配置文件中添加：

**Windows**：`%APPDATA%\Claude\claude_desktop_config.json`
**macOS/Linux**：`~/.config/Claude/claude_desktop_config.json`

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

#### 配置 Claude Code

在 Claude Code 设置文件中添加：

**Windows**：`%LOCALAPPDATA%\claude-code\settings.json`
**macOS/Linux**：`~/.config/claude-code/settings.json`

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

### 步骤 3：重启 Claude

重启 Claude Desktop 或 Claude Code 以加载 MCP 服务器。

## 快速开始

安装完成后，你可以在与 Claude 的对话中直接使用 OfficeReader-MCP：

```
将 D:\Reports\sales_2024.xlsx 转换为 markdown
```

```
从 D:\Presentations\keynote.pptx 提取文本和图像
```

```
获取 C:\Documents\report.docx 的元数据
```

## 可用工具

### 1. `convert_document`

将任何支持的 Office 文档转换为 Markdown 格式。

**参数：**
- `file_path`（必需）：文档的绝对路径
- `extract_images`（可选，默认：true）：是否提取嵌入图像
- `image_format`（可选，默认："file"）：图像处理方式
  - `"file"`：将图像保存到磁盘（推荐）
  - `"base64"`：将图像作为 base64 嵌入到 markdown 中
  - `"both"`：同时保存和嵌入
- `output_name`（可选）：输出文件的自定义名称

**示例：**
```
转换 D:\Documents\report.xlsx 并提取图像
```

### 2. `read_converted_markdown`

读取先前转换的 markdown 文件的完整内容。

**参数：**
- `markdown_path`（必需）：markdown 文件的路径

**示例：**
```
读取 D:\cache\output\report_abc12345\report_abc12345.md 的内容
```

### 3. `list_conversions`

列出所有缓存的文档转换及其详细信息。

**示例：**
```
列出所有已转换的文档
```

### 4. `clear_cache`

清除所有缓存的转换以释放磁盘空间。

**示例：**
```
清除文档缓存
```

### 5. `get_document_metadata`

从文档中提取元数据，无需完整转换（更快）。

**参数：**
- `file_path`（必需）：文档的路径

**示例：**
```
获取 D:\Documents\presentation.pptx 的元数据
```

### 6. `get_supported_formats`

获取所有支持的文件格式和扩展名列表。

**示例：**
```
officereader 支持哪些文件格式？
```

## 输出结构

转换的文档组织在缓存目录中：

```
cache/
└── output/
    └── document_name_abc12345/
        ├── document_name_abc12345.md    # 转换后的 markdown
        └── images/
            ├── image_001.webp           # 优化后的图像
            ├── slide2_image_002.webp
            └── excel_image_003.webp
```

## 图像优化

图像会自动优化以减小文件大小，同时保持质量：

- **最大尺寸**：1920×1080 像素（可配置）
- **格式**：WebP（首选）或 PNG/JPEG 回退
- **质量**：照片 80%，JPEG 85%，带透明度的图形使用无损 PNG
- **典型压缩率**：文件大小减少 50-80%
- **智能检测**：自动区分照片和图形

## 技术细节

### 架构

```
OfficeReader-MCP/
├── src/officereader_mcp/
│   ├── server.py              # MCP 服务器实现
│   ├── converter.py           # Word 转换器（DocxConverter、OfficeConverter）
│   ├── excel_converter.py    # Excel 到 Markdown 转换器
│   ├── pptx_converter.py     # PowerPoint 到 Markdown 转换器
│   ├── image_optimizer.py    # 图像压缩工具
│   └── __init__.py           # 包初始化
├── test/
│   ├── test_converter.py     # 基本功能测试
│   └── test_all_formats.py   # 综合测试套件
├── pyproject.toml            # 项目配置
└── README.md                 # 文档
```

### 依赖项

| 包 | 版本 | 用途 |
|-----|------|------|
| `mcp` | >=1.0.0 | 模型上下文协议 SDK |
| `python-docx` | >=1.1.0 | DOCX 文件解析和操作 |
| `mammoth` | >=1.6.0 | DOC/DOCX 到 HTML 转换（备用） |
| `Pillow` | >=10.0.0 | 图像处理和优化 |
| `markdownify` | >=0.11.0 | HTML 到 Markdown 转换 |
| `openpyxl` | >=3.1.0 | Excel 文件解析 |
| `python-pptx` | >=0.6.21 | PowerPoint 文件解析 |

运行 `pip install -e .` 时会自动安装所有依赖项。

## 测试

### 运行测试

```bash
# 基本转换器测试
python test/test_converter.py

# 所有格式的综合测试套件
python test/test_all_formats.py

# 使用特定文档进行测试
python test/test_converter.py path/to/your/document.docx
```

### 测试覆盖

测试套件验证：
- 模块导入和初始化
- 所有格式的转换器功能
- 图像提取和优化
- 文件类型检测
- 缓存管理
- 元数据提取

## 配置

OfficeReader-MCP 支持多种配置方式来自定义缓存位置和行为。

### 快速配置（推荐）

1. 复制示例配置文件：
   ```bash
   cp config.example.json config.json
   ```

2. 编辑 `config.json` 设置你的缓存目录：
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

3. 配置文件将在启动时自动加载。

详细配置选项请参阅 [CONFIG.md](CONFIG.md)。

### 环境变量

| 变量 | 描述 | 默认值 |
|------|------|--------|
| `OFFICEREADER_CACHE_DIR` | 缓存转换的目录 | 系统临时目录 |

使用示例：
```bash
# 设置自定义缓存目录
export OFFICEREADER_CACHE_DIR=/path/to/custom/cache

# 或在 Windows 中
set OFFICEREADER_CACHE_DIR=C:\path\to\custom\cache
```

**注意**：环境变量的优先级高于配置文件设置。

## 使用示例

### 转换包含多个工作表的 Excel

```
用户：转换我的 Excel 文件 D:\Reports\Q4_sales.xlsx

Claude：我来转换这个 Excel 文件。每个工作表将转换为 markdown 中的独立部分，
        并保留表格格式...

[输出包括所有工作表的 markdown 表格，保留格式]
```

### 提取 PowerPoint 内容

```
用户：从 D:\Presentations\product_launch.pptx 提取所有文本和图像

Claude：正在转换 PowerPoint 演示文稿。我会提取每张幻灯片的文本，
        包括演讲者备注和所有嵌入的图像...

[输出包括逐幻灯片的分解，包含图像和备注]
```

### 批量处理

```
用户：转换 D:\Documents\ 中的所有 Office 文档

Claude：我会转换每个文档并缓存结果以便快速访问...

[处理所有支持的文件并提供摘要]
```

## 故障排除

### "Module not found" 错误

```bash
# 重新安装包
pip install -e .
```

### 配置未加载

1. 验证配置文件位置是否正确
2. 检查 JSON 语法是否有效（使用 JSON 验证器）
3. 完全重启 Claude Desktop 或 Claude Code
4. 检查日志中的错误消息

### 图像未提取

可能原因：
- 文档包含链接图像（未嵌入）
- 缓存目录的写权限不足
- 文档库不支持该图像格式

解决方案：
```bash
# 验证缓存目录可写
ls -la /path/to/cache  # Unix/Mac
dir /path/to/cache     # Windows

# 检查图像是否嵌入
# 使用 convert_document 时显式设置 extract_images=true
```

### 编码问题

转换器全程使用 UTF-8 编码。如果看到乱码文本：
- 检查源文档编码
- 确保终端/控制台支持 UTF-8
- 尝试使用不同的系统区域设置进行转换

## 更新日志

### v2.0.0 (2024-11)

**主要功能：**
- 添加 Excel (.xlsx, .xls) 支持，包含多工作表转换
- 添加 PowerPoint (.pptx, .ppt) 支持，包含幻灯片提取
- 实现智能图像优化，支持 WebP 压缩
- 添加统一的 OfficeConverter 接口，支持所有文档类型
- 增强所有格式的元数据提取

**改进：**
- 基于哈希的文件识别智能缓存系统
- 格式特定转换器的延迟加载，提高性能
- 更好的错误处理和验证
- 所有格式的综合测试套件

**工具：**
- 添加 `get_supported_formats` 工具
- 增强 `get_document_metadata`，支持所有格式
- 改进 `list_conversions`，提供详细的缓存信息

### v1.0.0 (2024-09)

- 首次发布
- Word 文档 (.docx, .doc) 转换
- 基本图像提取
- MCP 服务器实现

## 贡献

欢迎贡献！你可以这样帮助：

1. **报告错误**：开启 issue，附上详细信息和复现步骤
2. **建议功能**：描述你的想法和使用场景
3. **提交 Pull Request**：
   - Fork 仓库
   - 创建功能分支（`git checkout -b feature/amazing-feature`）
   - 提交更改（`git commit -m 'Add amazing feature'`）
   - 推送到分支（`git push origin feature/amazing-feature`）
   - 开启 Pull Request

### 开发环境设置

```bash
# 克隆并安装开发依赖
git clone https://github.com/Asunainlove/office-reader-mcp.git
cd office-reader-mcp
pip install -e ".[dev]"

# 运行测试
python test/test_all_formats.py

# 运行代码检查（如果配置）
black src/
ruff check src/
```

## 许可证

MIT 许可证 - 详见 [LICENSE](LICENSE) 文件。

## 作者

**Asunainlove**

- GitHub：[@Asunainlove](https://github.com/Asunainlove)
- 仓库：[office-reader-mcp](https://github.com/Asunainlove/office-reader-mcp)
- 问题反馈：[报告错误](https://github.com/Asunainlove/office-reader-mcp/issues)

## 致谢

本项目使用了以下开源库：
- [Model Context Protocol (MCP)](https://modelcontextprotocol.io/) by Anthropic
- [python-docx](https://python-docx.readthedocs.io/) 用于 Word 处理
- [openpyxl](https://openpyxl.readthedocs.io/) 用于 Excel 处理
- [python-pptx](https://python-pptx.readthedocs.io/) 用于 PowerPoint 处理
- [Pillow](https://pillow.readthedocs.io/) 用于图像处理

## 支持

如果你觉得这个项目有帮助，请：
- ⭐ 给仓库点星
- 🐛 报告错误和问题
- 💡 建议新功能
- 🔀 贡献代码改进

---

**祝转换愉快！** 🚀
