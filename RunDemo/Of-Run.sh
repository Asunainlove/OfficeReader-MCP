#!/bin/bash
# OfficeReader-MCP Launcher
# This script starts the OfficeReader MCP server

cd "$(dirname "$0")/.."
python -m officereader_mcp.server
