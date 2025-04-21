@echo off
echo Setting up MCP Visio server with Relay Service...

REM Start the Visio relay service
echo Starting Visio Relay Service...
start "Visio Relay Service" python visio_relay.py

REM Wait for the relay service to start
echo Waiting for relay service to start...
timeout /t 3 /nobreak > nul

REM Stop any existing container
echo Stopping existing containers...
docker stop mcp-visio 2>nul
docker rm mcp-visio 2>nul

REM Build and run the container
echo Building and starting Docker container...
docker build -t mcp-visio-windows -f Dockerfile.windows .
docker run -d --name mcp-visio ^
  -p 8050:8050 ^
  -e TRANSPORT=sse ^
  -e HOST=0.0.0.0 ^
  -e PORT=8050 ^
  -e VISIO_SERVICE_HOST=host.docker.internal ^
  -e VISIO_SERVICE_PORT=8051 ^
  --isolation=process ^
  -v %CD%\src:C:\app\src ^
  -v %CD%\visio-files:C:\visio-files ^
  mcp-visio-windows

echo.
echo MCP Setup Complete!
echo.
echo Running Services:
echo - Visio Relay Service: http://localhost:8051
echo - MCP Server: http://localhost:8050/mcp
echo.
echo Use this MCP server URL in Cursor: http://localhost:8050/mcp
echo.
echo Note: To stop all services, run:
echo   docker stop mcp-visio
echo   Then close the Visio Relay Service window
echo. 