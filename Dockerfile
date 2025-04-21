FROM python:3.9-slim

WORKDIR /app

# Copy requirements file
COPY requirements.txt .

# Install dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY . .

# Create directory for Visio files
RUN mkdir -p /visio-files

# Expose port for FastAPI
EXPOSE 8050

# Command to run the FastAPI application
CMD ["python", "mcp_server.py"] 