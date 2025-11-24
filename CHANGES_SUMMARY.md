# 配置系统更新总结

## 概述

为 OfficeReader-MCP 添加了完整的配置文件系统，允许用户通过 `config.json` 文件轻松自定义缓存存储位置和其他设置。

## 新增文件

### 1. 核心配置文件

| 文件 | 说明 |
|------|------|
| `src/officereader_mcp/config_loader.py` | 配置加载器，负责读取和解析配置文件 |
| `config.json` | 用户配置文件（已添加到 .gitignore） |
| `config.example.json` | 配置文件示例模板 |

### 2. 文档文件

| 文件 | 说明 |
|------|------|
| `CONFIG.md` | 详细的配置说明文档（中文） |
| `QUICKSTART.md` | 快速入门指南（中文） |
| `test_config.py` | 配置测试脚本 |

## 修改的文件

| 文件 | 修改内容 |
|------|----------|
| `src/officereader_mcp/server.py` | 集成配置加载器，替换原有的环境变量方式 |
| `README.md` | 添加配置章节，说明配置文件使用方法 |
| `README-CN.md` | 添加配置章节（中文版） |
| `.gitignore` | 添加 `config.json` 和 `.officereader-mcp/` 到忽略列表 |

## 主要功能

### 1. 配置文件位置

系统会按以下顺序搜索 `config.json`：
1. 当前工作目录
2. 项目根目录
3. 用户主目录 (`~/.officereader-mcp/config.json`)

### 2. 配置选项

```json
{
  "cache_dir": "cache",                    // 缓存目录（支持相对/绝对路径）
  "output_dir": null,                      // 自定义输出目录（可选）
  "image_optimization": {                  // 图像优化设置
    "enabled": true,                       // 是否启用优化
    "max_dimension": 1920,                 // 最大尺寸
    "quality": 80                          // 压缩质量
  },
  "default_settings": {                    // 默认转换设置
    "extract_images": true,                // 是否提取图像
    "image_format": "file"                 // 图像格式（file/base64/both）
  }
}
```

### 3. 配置优先级

1. **环境变量** `OFFICEREADER_CACHE_DIR`（最高优先级）
2. **配置文件** `config.json`
3. **默认值**（系统临时目录）

## 使用方法

### 快速开始

```bash
# 1. 复制示例配置
cp config.example.json config.json

# 2. 编辑配置文件
# 修改 cache_dir 为你想要的路径

# 3. 测试配置
python test_config.py

# 4. 重启 Claude
```

### 配置示例

#### Windows 用户
```json
{
  "cache_dir": "D:/Documents/OfficeReaderCache"
}
```

#### Linux/Mac 用户
```json
{
  "cache_dir": "/home/user/cache/officereader"
}
```

#### 使用相对路径
```json
{
  "cache_dir": "cache"  // 相对于 config.json 所在目录
}
```

## 向后兼容性

- ✅ 保留了原有的环境变量 `OFFICEREADER_CACHE_DIR` 支持
- ✅ 如果没有配置文件，系统会使用默认值（系统临时目录）
- ✅ 现有用户无需修改配置即可继续使用

## 测试

运行测试脚本验证配置：

```bash
python test_config.py
```

成功输出示例：

```
============================================================
  OfficeReader-MCP Configuration Test
============================================================

Config file location: D:\OfficeReader-MCP\config.json
Config loaded: Yes

Cache Directory: D:\Documents\OfficeReaderCache

[SUCCESS] All tests passed!
```

## 文档链接

- [配置详细说明](CONFIG.md)
- [快速入门指南](QUICKSTART.md)
- [README](README.md)
- [中文 README](README-CN.md)

## 技术细节

### config_loader.py 模块

提供以下功能：
- `Config` 类：配置管理器
- `get_config()`: 获取全局配置实例
- `load_config()`: 重新加载配置

### server.py 集成

```python
from .config_loader import get_config

def get_converter() -> OfficeConverter:
    config = get_config()
    cache_dir = config.get_cache_dir()
    return OfficeConverter(cache_dir=cache_dir)
```

## 下一步

用户可以根据需要：
1. 自定义缓存位置
2. 调整图像优化参数
3. 配置默认转换行为
4. 使用多个配置文件（通过改变工作目录）

---

**更新日期**: 2024-11-24
**版本**: v2.1.0
