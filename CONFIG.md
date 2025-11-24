# OfficeReader-MCP 配置说明

## 配置文件位置

OfficeReader-MCP 会按以下顺序查找 `config.json` 配置文件：

1. **当前工作目录** - `./config.json`
2. **项目根目录** - `OfficeReader-MCP/config.json`
3. **用户主目录** - `~/.officereader-mcp/config.json` (Windows: `C:\Users\<用户名>\.officereader-mcp\config.json`)

## 快速开始

1. 复制示例配置文件：
   ```bash
   cp config.example.json config.json
   ```

2. 编辑 `config.json` 修改缓存路径：
   ```json
   {
     "cache_dir": "D:/MyDocuments/OfficeReaderCache",
     "output_dir": null,
     "image_optimization": {
       "enabled": true,
       "max_dimension": 1920,
       "quality": 80
     },
     "default_settings": {
       "extract_images": true,
       "image_format": "file"
     }
   }
   ```

## 配置选项详解

### cache_dir (缓存目录)

**类型**: `string`
**默认值**: 系统临时目录
**说明**: 指定文档转换缓存的存储位置

**示例**:
```json
{
  "cache_dir": "cache"  // 相对路径，相对于 config.json 所在目录
}
```

```json
{
  "cache_dir": "D:/MyCache/OfficeReader"  // Windows 绝对路径
}
```

```json
{
  "cache_dir": "/home/user/cache/officereader"  // Linux/Mac 绝对路径
}
```

**注意**:
- 相对路径会相对于 `config.json` 文件所在目录
- 绝对路径直接使用指定路径
- 环境变量 `OFFICEREADER_CACHE_DIR` 的优先级高于配置文件

### output_dir (输出目录)

**类型**: `string | null`
**默认值**: `null` (使用 cache_dir/output)
**说明**: 自定义转换后的 Markdown 文件输出位置

**示例**:
```json
{
  "output_dir": "D:/Documents/Converted"
}
```

### image_optimization (图像优化设置)

**类型**: `object`
**说明**: 控制提取图像的优化参数

#### enabled
- **类型**: `boolean`
- **默认值**: `true`
- **说明**: 是否启用图像优化

#### max_dimension
- **类型**: `integer`
- **默认值**: `1920`
- **说明**: 图像最大尺寸（宽或高），超过将被缩放

#### quality
- **类型**: `integer` (0-100)
- **默认值**: `80`
- **说明**: 图像压缩质量

**示例**:
```json
{
  "image_optimization": {
    "enabled": true,
    "max_dimension": 2560,
    "quality": 90
  }
}
```

### default_settings (默认转换设置)

**类型**: `object`
**说明**: 文档转换的默认参数

#### extract_images
- **类型**: `boolean`
- **默认值**: `true`
- **说明**: 是否提取文档中的图像

#### image_format
- **类型**: `string`
- **可选值**: `"file"`, `"base64"`, `"both"`
- **默认值**: `"file"`
- **说明**: 图像处理方式
  - `"file"`: 保存为文件（推荐）
  - `"base64"`: 嵌入 Markdown 为 base64
  - `"both"`: 同时保存文件和嵌入

**示例**:
```json
{
  "default_settings": {
    "extract_images": true,
    "image_format": "file"
  }
}
```

## 配置优先级

配置参数的优先级从高到低：

1. **环境变量** - `OFFICEREADER_CACHE_DIR`
2. **配置文件** - `config.json` 中的设置
3. **默认值** - 代码中的默认值

## 环境变量配置

除了配置文件，你也可以使用环境变量：

### Windows
```cmd
set OFFICEREADER_CACHE_DIR=D:\MyCache\OfficeReader
```

### Linux/Mac
```bash
export OFFICEREADER_CACHE_DIR=/home/user/cache/officereader
```

### Claude Desktop/Code 配置
在 Claude 的配置文件中设置：

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

## 完整配置示例

### 示例 1: 使用默认缓存位置
```json
{
  "cache_dir": "cache",
  "output_dir": null,
  "image_optimization": {
    "enabled": true,
    "max_dimension": 1920,
    "quality": 80
  },
  "default_settings": {
    "extract_images": true,
    "image_format": "file"
  }
}
```

### 示例 2: 自定义所有路径
```json
{
  "cache_dir": "D:/Documents/OfficeReaderCache",
  "output_dir": "D:/Documents/ConvertedMarkdown",
  "image_optimization": {
    "enabled": true,
    "max_dimension": 2560,
    "quality": 90
  },
  "default_settings": {
    "extract_images": true,
    "image_format": "file"
  }
}
```

### 示例 3: 禁用图像优化
```json
{
  "cache_dir": "cache",
  "output_dir": null,
  "image_optimization": {
    "enabled": false
  },
  "default_settings": {
    "extract_images": true,
    "image_format": "both"
  }
}
```

## 故障排查

### 配置文件未加载
- 检查 JSON 语法是否正确（可使用 JSON 验证器）
- 确保文件编码为 UTF-8
- 查看启动日志，确认配置文件路径

### 缓存路径无效
- 确保指定的路径存在或可以被创建
- 检查文件系统权限
- Windows 路径使用正斜杠 `/` 或双反斜杠 `\\`

### 环境变量不生效
- 重启 Claude Desktop/Code
- 检查环境变量名称是否正确：`OFFICEREADER_CACHE_DIR`
- 在配置文件中的 `env` 对象中设置

## 查看当前配置

转换文档时，MCP 会在响应中显示当前使用的缓存位置：

```
[Cache Location] D:\Documents\OfficeReaderCache
```

你也可以使用 `list_conversions` 工具查看缓存信息。
