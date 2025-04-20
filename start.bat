@echo off
echo Starting MCP-Visio server...
echo.

REM Activate virtual environment if it exists
if exist venv\Scripts\activate.bat (
    call venv\Scripts\activate.bat
)

REM Use SSE transport by default
set TRANSPORT=sse
set HOST=0.0.0.0
set PORT=8050

REM Check for command-line arguments
if "%1"=="stdio" (
    set TRANSPORT=stdio
    echo Using STDIO transport
) else (
    echo Using SSE transport on %HOST%:%PORT%
)

REM Start the server
python src/main.py --transport %TRANSPORT% --host %HOST% --port %PORT%

REM Deactivate virtual environment if it was activated
if exist venv\Scripts\activate.bat (
    call venv\Scripts\deactivate.bat
) 