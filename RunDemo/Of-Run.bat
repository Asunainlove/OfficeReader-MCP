@echo off
REM OfficeReader-MCP Launcher
REM This script starts the Word document MCP server

cd /d "%~dp0.."
python -m src.officereader_mcp
