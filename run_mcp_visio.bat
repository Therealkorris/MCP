@echo off
echo ===================================================
echo MCP Visio Server - All-in-One Setup
echo ===================================================

setlocal enabledelayedexpansion
set PORT=8051
set DOCKER_PORT=8050
set DEBUG=true

REM Check command line arguments
if "%1"=="relay" goto :start_relay
if "%1"=="docker" goto :start_docker
if "%1"=="local_mcp" goto :start_local_mcp
if "%1"=="test" goto :run_test
if "%1"=="stop" goto :stop_all

REM No arguments, show help menu
if "%1"=="" (
    echo Select an option:
    echo.
    echo 1. Start with Docker (recommended)
    echo 2. Start both services locally (no Docker)
    echo 3. Start Visio relay only
    echo 4. Start MCP Docker container only
    echo 5. Start MCP locally only
    echo 6. Run tests
    echo 7. Stop all services
    echo 8. Toggle auto-reload (currently !DEBUG!)
    echo 9. Exit
    echo.
    
    set /p choice=Enter your choice (1-9): 
    
    if "!choice!"=="1" goto :start_all_docker
    if "!choice!"=="2" goto :start_all_local
    if "!choice!"=="3" goto :start_relay
    if "!choice!"=="4" goto :start_docker
    if "!choice!"=="5" goto :start_local_mcp
    if "!choice!"=="6" goto :run_test
    if "!choice!"=="7" goto :stop_all
    if "!choice!"=="8" goto :toggle_debug
    if "!choice!"=="9" goto :end
    
    echo Invalid choice. Please try again.
    goto :eof
)

:start_all_docker
echo Starting complete MCP Visio environment with Docker...
start "Visio Relay" cmd /c "%~f0 relay"
timeout /t 3 /nobreak > nul
call :start_docker
echo.
echo All services started successfully!
echo Visio Relay: http://localhost:%PORT%
echo MCP Server (Docker): http://localhost:%DOCKER_PORT%
echo Auto-reload: %DEBUG%
echo.
echo Use this URL in your applications: http://localhost:%DOCKER_PORT%/mcp
goto :end

:start_all_local
echo Starting complete MCP Visio environment locally (no Docker)...
start "Visio Relay" cmd /c "%~f0 relay"
timeout /t 3 /nobreak > nul
start "MCP Server" cmd /c "%~f0 local_mcp"
echo.
echo All services started locally!
echo Visio Relay: http://localhost:%PORT%
echo MCP Server (Local): http://localhost:%DOCKER_PORT%
echo Auto-reload: %DEBUG%
echo.
echo Use this URL in your applications: http://localhost:%DOCKER_PORT%/mcp
goto :end

:start_relay
echo.
echo Starting Visio Relay Service...
echo This service allows the MCP container to communicate with Microsoft Visio.
echo Make sure Visio is installed on this machine.
echo.
echo Checking for required Python packages...
python -m pip install fastapi uvicorn pywin32 pydantic --quiet
echo Starting Visio Relay Service on port %PORT% with auto-reload: %DEBUG%...
echo.
echo Press Ctrl+C to stop the service
echo.
REM Run the relay service with auto-reload already enabled in the code
start "Visio Relay Service" cmd /c python host_visio_relay.py
echo Visio Relay Service started in a new window.
goto :end

:start_local_mcp
echo.
echo Starting MCP Server locally...
echo.
echo Checking for required Python packages...
python -m pip install fastapi uvicorn pydantic python-dotenv --quiet
echo Starting MCP Server on port %DOCKER_PORT% with auto-reload: %DEBUG%...
echo.
echo Press Ctrl+C to stop the service
echo.
REM Set environment variables for local MCP server
set VISIO_SERVICE_HOST=localhost
set VISIO_SERVICE_PORT=%PORT%
set HOST=0.0.0.0
set PORT=%DOCKER_PORT%
set TRANSPORT=sse
set DEBUG=%DEBUG%

REM Choose which MCP server implementation to run
if exist mcp_server.py (
    echo Running via mcp_server.py...
    if "%DEBUG%"=="true" (
        start "MCP Server" cmd /c python mcp_server.py --debug
    ) else (
        start "MCP Server" cmd /c python mcp_server.py
    )
) else (
    echo Running via src/main.py...
    if "%DEBUG%"=="true" (
        start "MCP Server" cmd /c python src/main.py --debug
    ) else (
        start "MCP Server" cmd /c python src/main.py
    )
)
echo MCP Server started in a new window.
goto :end

:start_docker
echo.
echo Building and starting MCP Visio Docker container...
docker stop mcp-visio 2>nul
docker rm mcp-visio 2>nul
docker build -t mcp-visio-windows -f Dockerfile.windows .
docker run -d --name mcp-visio ^
  -p %DOCKER_PORT%:%DOCKER_PORT% ^
  -e TRANSPORT=sse ^
  -e HOST=0.0.0.0 ^
  -e PORT=%DOCKER_PORT% ^
  -e VISIO_SERVICE_HOST=host.docker.internal ^
  -e VISIO_SERVICE_PORT=%PORT% ^
  -e DEBUG=%DEBUG% ^
  --isolation=process ^
  -v %CD%\src:C:\app\src ^
  -v %CD%\visio-files:C:\visio-files ^
  mcp-visio-windows
echo.
echo Docker container started with auto-reload: %DEBUG%!
goto :eof

:toggle_debug
if "%DEBUG%"=="true" (
    set DEBUG=false
) else (
    set DEBUG=true
)
echo Auto-reload has been set to: %DEBUG%
goto :eof

:run_test
echo.
echo Running Visio connection tests...
python test_visio_connection.py
goto :end

:stop_all
echo.
echo Stopping all MCP Visio services...
docker stop mcp-visio 2>nul
docker rm mcp-visio 2>nul
echo.
echo Stopped. You may need to manually close any Visio Relay terminal windows.
goto :end

:end
echo. 