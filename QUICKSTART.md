# OfficeReader-MCP 快速入门

## 1. 安装

```bash
# 克隆仓库
git clone https://github.com/Asunainlove/office-reader-mcp.git
cd office-reader-mcp

# 安装依赖
pip install -e .
```

## 2. 配置缓存路径

### 方法一：使用配置文件（推荐）

```bash
# 复制示例配置
cp config.example.json config.json
```

编辑 `config.json`：

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

### 方法二：使用环境变量

**Windows:**
```cmd
set OFFICEREADER_CACHE_DIR=D:\MyDocuments\OfficeReaderCache
```

**Linux/Mac:**
```bash
export OFFICEREADER_CACHE_DIR=/home/user/cache/officereader
```

## 3. 配置 Claude

### Claude Desktop

编辑配置文件：
- **Windows**: `%APPDATA%\Claude\claude_desktop_config.json`
- **macOS**: `~/.config/Claude/claude_desktop_config.json`
- **Linux**: `~/.config/Claude/claude_desktop_config.json`

添加以下内容：

```json
{
  "mcpServers": {
    "officereader": {
      "command": "python",
      "args": ["-m", "officereader_mcp.server"]
    }
  }
}
```

如果使用环境变量：

```json
{
  "mcpServers": {
    "officereader": {
      "command": "python",
      "args": ["-m", "officereader_mcp.server"],
      "env": {
        "OFFICEREADER_CACHE_DIR": "D:/MyCache/OfficeReader"
      }
    }
  }
}
```

### Claude Code

编辑配置文件：
- **Windows**: `%LOCALAPPDATA%\claude-code\settings.json`
- **macOS/Linux**: `~/.config/claude-code/settings.json`

配置格式与 Claude Desktop 相同。

## 4. 测试配置

运行测试脚本验证配置：

```bash
python test_config.py
```

你应该看到类似输出：

```
============================================================
  OfficeReader-MCP Configuration Test
============================================================

Config file location: D:\OfficeReader-MCP\config.json
Config loaded: Yes

Cache Directory: D:\MyDocuments\OfficeReaderCache

[SUCCESS] All tests passed!
```

## 5. 使用

重启 Claude 后，你可以在对话中使用：

```
转换我的 Word 文档：D:\Documents\report.docx
```

```
将 Excel 文件转换为 Markdown：D:\Reports\sales_2024.xlsx
```

```
提取 PowerPoint 演示文稿的内容：D:\Presentations\keynote.pptx
```

## 配置优先级

1. **环境变量** `OFFICEREADER_CACHE_DIR` (最高优先级)
2. **配置文件** `config.json`
3. **默认值** (系统临时目录)

## 常见问题

### 找不到配置文件？

配置文件搜索顺序：
1. 当前工作目录：`./config.json`
2. 项目根目录：`OfficeReader-MCP/config.json`
3. 用户主目录：`~/.officereader-mcp/config.json`

### 缓存位置在哪？

使用 `list_conversions` 工具可以查看当前缓存位置：

```
列出所有已转换的文档
```

### 更改配置后不生效？

1. 确保 JSON 格式正确（无语法错误）
2. 重启 Claude Desktop/Code
3. 运行 `python test_config.py` 验证配置

## 更多帮助

- 详细配置选项：[CONFIG.md](CONFIG.md)
- 完整文档：[README.md](README.md)
- 中文文档：[README-CN.md](README-CN.md)
