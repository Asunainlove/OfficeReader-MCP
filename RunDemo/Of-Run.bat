@echo off
REM OfficeReader-MCP Launcher
REM This script starts the OfficeReader MCP server

cd /d "%~dp0.."
python -m officereader_mcp.server
