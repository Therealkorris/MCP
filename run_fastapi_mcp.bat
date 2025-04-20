@echo off
echo Setting up FastAPI MCP for Visio...
echo.

rem Install required packages
echo Installing required packages...
pip install -r requirements.txt

echo.
echo Starting MCP server for Visio...
echo The MCP server will be available at http://localhost:8050/mcp
echo.

python mcp_server.py

echo.
echo Server stopped. 