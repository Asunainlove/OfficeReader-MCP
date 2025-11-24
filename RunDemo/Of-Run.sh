#!/bin/bash
# OfficeReader-MCP Launcher
# This script starts the Word document MCP server

cd "$(dirname "$0")/.."
python -m src.officereader_mcp
