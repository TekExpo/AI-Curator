#!/bin/bash
# Startup script for Azure App Service (Linux)

set -e

# Log startup
echo "Starting FastAPI application..."

# Set PORT if not set (Azure App Service provides PORT)
if [ -z "$PORT" ]; then
    export PORT=8000
    echo "PORT not set, defaulting to $PORT"
else
    echo "PORT is set to $PORT"
fi

# Check if we're in Python environment
if command -v python3 &> /dev/null; then
    echo "Python3 found at: $(which python3)"
    PYTHON_CMD="python3"
elif command -v python &> /dev/null; then
    echo "Python found at: $(which python)"
    PYTHON_CMD="python"
else
    echo "Python not found, trying /opt/python/3.10/bin/python3"
    PYTHON_CMD="/opt/python/3.10/bin/python3"
fi

# Check if pip is available
if command -v pip3 &> /dev/null; then
    PIP_CMD="pip3"
elif command -v pip &> /dev/null; then
    PIP_CMD="pip"
else
    echo "Pip not found, trying /opt/python/3.10/bin/pip3"
    PIP_CMD="/opt/python/3.10/bin/pip3"
fi

# Install dependencies if requirements.txt exists
if [ -f "requirements.txt" ]; then
    echo "Installing dependencies..."
    $PIP_CMD install -r requirements.txt
else
    echo "No requirements.txt found"
fi

# Check if gunicorn is available
if command -v gunicorn &> /dev/null; then
    echo "Starting with gunicorn..."
    exec gunicorn --bind=0.0.0.0:$PORT --timeout 600 -w 2 -k uvicorn.workers.UvicornWorker asgi:app
else
    echo "Gunicorn not found, trying with uvicorn directly..."
    if command -v uvicorn &> /dev/null; then
        exec uvicorn asgi:app --host 0.0.0.0 --port $PORT
    else
        echo "Uvicorn not found, trying with python..."
        exec $PYTHON_CMD -m uvicorn asgi:app --host 0.0.0.0 --port $PORT
    fi
fi
