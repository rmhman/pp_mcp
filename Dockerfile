# Use Python slim image
FROM python:3.11-slim

# Set working directory
WORKDIR /app

# Set Python unbuffered mode
ENV PYTHONUNBUFFERED=1

# Copy requirements first for better caching
COPY requirements.txt .

# Install dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copy the server code
COPY powerpoint_server.py .

# Create non-root user with home directory and PowerPoint directory
RUN useradd -m -u 1000 mcpuser && \
    chown -R mcpuser:mcpuser /app && \
    mkdir -p /home/mcpuser/PowerPoints && \
    chown -R mcpuser:mcpuser /home/mcpuser/PowerPoints

# Switch to non-root user
USER mcpuser

# Set working directory to user's home for proper file access
WORKDIR /home/mcpuser

# Run the server with stdio transport
CMD ["python", "/app/powerpoint_server.py"]