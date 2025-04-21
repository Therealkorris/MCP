# PowerShell script to run the appropriate Docker Compose file based on OS

# Check if Docker is running in Windows container mode
$containerPlatform = docker info --format '{{.OSType}}'
Write-Host "Current Docker container platform: $containerPlatform"

if ($containerPlatform -eq "windows") {
    Write-Host "Running in Windows container mode - starting Visio MCP server with Windows container..."
    docker compose -f docker-compose.windows.yml up -d --build
}
else {
    Write-Host "Running in Linux container mode - Windows container mode required for Visio integration."
    Write-Host "Switch to Windows containers in Docker Desktop and run this script again."
    Write-Host "Alternatively, to run without Visio integration, use: docker compose up -d --build"
}

Write-Host "Done!" 