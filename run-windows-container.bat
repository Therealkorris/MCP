@echo off
echo Running MCP Visio in Windows container...

REM Check if Docker is running in Windows container mode
FOR /F "tokens=*" %%i IN ('docker info ^| findstr "OSType"') DO SET DOCKER_OS=%%i
echo %DOCKER_OS% | findstr "linux" > nul
IF %ERRORLEVEL% EQU 0 (
  echo ERROR: Docker is running in Linux container mode. Please switch to Windows containers.
  echo Right-click the Docker Desktop icon in the system tray and select "Switch to Windows containers..."
  exit /b 1
)

REM Stop existing container if running
echo Stopping existing container if running...
docker stop mcp-visio 2>nul
docker rm mcp-visio 2>nul

REM Build the container
echo Building Windows container...
docker build -t mcp-visio-windows -f Dockerfile.windows .

REM Run the container with proper Visio access
echo Running the container...
docker run -d --name mcp-visio ^
  -p 8050:8050 ^
  -e TRANSPORT=sse ^
  -e HOST=0.0.0.0 ^
  -e PORT=8050 ^
  -e VISIO_SERVICE_HOST=localhost ^
  -e VISIO_SERVICE_PORT=8051 ^
  --isolation=process ^
  -v %CD%\src:C:\app\src ^
  -v %CD%\visio-files:C:\visio-files ^
  mcp-visio-windows

echo.
echo Container started! Your Visio files should be placed in the 'visio-files' directory.
echo The MCP server is accessible at: http://localhost:8050/mcp
echo.
echo To verify everything is working, try:
echo docker logs mcp-visio
echo.
echo Done! 