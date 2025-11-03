#!/bin/bash

# PowerPoint MCP Server Setup Script
# This script helps set up and test the PowerPoint MCP server

set -e

echo "🚀 PowerPoint MCP Server Setup"
echo "=============================="

# Create PowerPoint directory on host if it doesn't exist
echo "📁 Creating PowerPoint directory..."
mkdir -p /tmp/PowerPoints
echo "✅ PowerPoint directory created at /tmp/PowerPoints"

# Build Docker image
echo "🔨 Building Docker image..."
docker build -t powerpoint-mcp .

# Stop any existing container
echo "🛑 Stopping existing container..."
docker stop powerpoint-mcp-server 2>/dev/null || true
docker rm powerpoint-mcp-server 2>/dev/null || true

# Start container with volume mount for testing
echo "🚀 Starting PowerPoint MCP server for testing..."
docker run -d \
  --name powerpoint-mcp-server \
  -v /tmp/PowerPoints:/home/mcpuser/PowerPoints \
  -e PYTHONUNBUFFERED=1 \
  --user 1000:1000 \
  --entrypoint /bin/bash \
  powerpoint-mcp \
  -c "while true; do sleep 30; done"

echo "✅ PowerPoint MCP server is running!"
echo ""
echo "📋 Container Status:"
docker ps --filter name=powerpoint-mcp-server

echo ""
echo "🔍 To test the server:"
echo "1. Check container logs: docker logs powerpoint-mcp-server"
echo "2. Test MCP tools through your AI agent"
echo "3. Check /tmp/PowerPoints directory for created files"
echo ""
echo "🛠️  Management Commands:"
echo "- Stop server: docker stop powerpoint-mcp-server"
echo "- Start server: docker start powerpoint-mcp-server"
echo "- View logs: docker logs -f powerpoint-mcp-server"
echo "- Remove server: docker rm -f powerpoint-mcp-server"
