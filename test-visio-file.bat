@echo off
echo Testing Visio file access from container...

REM Check if a parameter was passed
if "%~1"=="" (
  echo ERROR: Please provide a Visio file name as parameter.
  echo Example: test-visio-file.bat my-diagram.vsdx
  exit /b 1
)

REM Get the filename
set FILE_NAME=%~1

REM Check if the file exists
if not exist "visio-files\%FILE_NAME%" (
  echo ERROR: File not found: visio-files\%FILE_NAME%
  exit /b 1
)

echo Testing access to: %FILE_NAME%

REM Run the test script in the container
docker exec mcp-visio python src/test_visio_access.py "C:\visio-files\%FILE_NAME%"

echo.
echo If you see "Visio access test successful!" above, the container can properly access Visio files.
echo. 