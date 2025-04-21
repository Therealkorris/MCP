@echo off
echo Preparing Visio test environment...

REM Create the visio-files directory if it doesn't exist
if not exist visio-files mkdir visio-files

REM Display instructions
echo.
echo ==============================================
echo VISIO FILE TEST SETUP
echo ==============================================
echo.
echo Please copy your Visio (.vsdx) files to the 'visio-files' directory.
echo.
echo Example file paths to use in the MCP:
echo - C:\visio-files\my-diagram.vsdx (from container)
echo - visio-files\my-diagram.vsdx (relative path)
echo.
echo After copying your files, restart the container using:
echo run-windows-container.bat
echo.
echo ==============================================
echo.

REM Open the directory
explorer visio-files

echo Done! 