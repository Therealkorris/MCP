# MCP Visio Server - Windows Container Setup

This document explains how to run the MCP Visio Server using Windows containers for full Visio integration.

## Prerequisites

1. Windows Server 2019/2022 or Windows 10/11 Pro/Enterprise
2. Docker Desktop for Windows installed
3. Docker configured to use Windows containers (not Linux containers)
4. Microsoft Visio installed on the host machine

## Setup Instructions

### 1. Switch Docker to Windows Containers Mode

Right-click the Docker Desktop icon in the system tray and select "Switch to Windows containers..."

### 2. Prepare Visio Files

1. Run the preparation script:
   ```
   prepare-visio-test.bat
   ```

2. Copy your Visio (.vsdx) files to the `visio-files` directory that opens.

3. These files will be accessible from within the container at `C:\visio-files\your-file.vsdx`.

### 3. Build and Run the Windows Container

```powershell
# Use the provided batch file
run-windows-container.bat

# Or run docker compose directly
docker compose -f docker-compose.windows.yml up -d --build
```

### 4. Verify the Container is Running

```powershell
docker ps
```

You should see the `mcp-visio` container running.

### 5. Test Visio Integration

To verify that the container can access and manipulate Visio files:

```powershell
# Check container logs
docker logs mcp-visio

# Run a test on a specific file
docker exec mcp-visio python src/test_visio_access.py C:\visio-files\your-file.vsdx
```

### 6. Configure MCP in Cursor

Add a new MCP server with the URL: `http://localhost:8050/mcp`

## Working with Visio Files

When using the MCP to work with Visio files, you can refer to files in several ways:

1. **Absolute path** within the container:
   ```
   C:\visio-files\your-diagram.vsdx
   ```

2. **Relative path** from the mounted directory:
   ```
   visio-files\your-diagram.vsdx
   ```

The container is configured with file sharing that makes the `visio-files` directory accessible both on your host machine and inside the container.

## Troubleshooting

### Issues with Visio Access

The Windows container runs in "process isolation" mode to allow access to Visio on the host. If you encounter issues with Visio access:

1. Ensure Visio is installed and accessible on the host
2. Check container logs: `docker logs mcp-visio`
3. Verify the container has the necessary permissions to access Visio

### Container Build Failures

Windows containers are larger and take longer to build than Linux containers. Be patient during the initial build process.

If the build fails, try:
```powershell
docker system prune -a
docker compose -f docker-compose.windows.yml build --no-cache
``` 