@echo off
echo Testing access to active Visio documents from container...
echo.
echo IMPORTANT: Make sure you have a Visio document open in Visio on your host machine!
echo.

REM Run the test script in the container
docker exec mcp-visio python src/test_active_visio.py

echo.
echo If you see "Active Visio document test successful!" above, the container can access your open Visio document.
echo. 
pause