# OfficeReader-MCP 启动脚本

此目录包含用于启动 OfficeReader MCP 服务器的便捷脚本。

## 使用方法

### Windows 用户
```cmd
Of-Run.bat
```

### Linux/macOS 用户
```bash
chmod +x Of-Run.sh
./Of-Run.sh
```

## 前置要求

在运行这些脚本之前，请确保：

1. 已安装 Python 3.10 或更高版本
2. 已安装项目依赖：
   ```bash
   cd ..
   pip install -e .
   ```
   或者
   ```bash
   pip install -r requirements.txt
   ```

## 注意事项

- 这些脚本会从项目根目录启动 MCP 服务器
- 服务器使用标准输入/输出进行通信（stdio）
- 通常这些脚本在 Claude Desktop/Code 配置中被调用，而不是手动运行
- 如果需要测试，可以手动运行并通过 stdin 发送 JSON-RPC 消息

## 配置

服务器会读取以下配置（按优先级）：
1. 环境变量 `OFFICEREADER_CACHE_DIR`
2. 项目根目录的 `config.json` 文件
3. 默认缓存目录（系统临时目录）

示例 Claude Desktop 配置：
```json
{
  "mcpServers": {
    "officereader": {
      "command": "path/to/OfficeReader-MCP/RunDemo/Of-Run.bat",
      "env": {
        "OFFICEREADER_CACHE_DIR": "D:/cache/officereader"
      }
    }
  }
}
```
