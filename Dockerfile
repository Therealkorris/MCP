FROM python:3.12-slim

# Build arguments
ARG PORT=8050

# Set working directory
WORKDIR /app

# Set environment variables
ENV PYTHONUNBUFFERED=1 \
    PORT=${PORT} \
    TRANSPORT=sse \
    HOST=0.0.0.0

# Copy requirements file
COPY requirements.txt .

# Install dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY . .

# Expose port
EXPOSE ${PORT}

# Set entrypoint
ENTRYPOINT ["python", "src/main.py"] 