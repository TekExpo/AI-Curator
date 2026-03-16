FROM python:3.10-slim

WORKDIR /app

# Install system dependencies
RUN apt-get update && apt-get install -y \
    gcc \
    curl \
    && rm -rf /var/lib/apt/lists/*

# Copy requirements
COPY requirements.txt .

# Install Python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copy application files
COPY . .

# Create data directory for SQLite database
RUN mkdir -p /app/data

# Expose port
EXPOSE 8000

# Health check endpoint
HEALTHCHECK --interval=30s --timeout=10s --start-period=5s --retries=3 \
    CMD curl -f http://localhost:${PORT:-8000}/health || exit 1

# Run with gunicorn and uvicorn worker for Azure compatibility
# Use PORT environment variable from Azure, default to 8000
CMD exec gunicorn --bind 0.0.0.0:${PORT:-8000} --timeout 600 --workers 4 --worker-class uvicorn.workers.UvicornWorker --access-logfile - --error-logfile - main:app
