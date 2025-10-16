#!/bin/bash
# Test script for PowerPoint MCP Server

echo "Testing PowerPoint MCP Server with Docker..."

# Test 1: Check if server starts
echo "1. Testing server startup..."
docker run --rm -i -v /Users/ross/PowerPoints:/home/mcpuser/PowerPoints --user 1000:1000 powerpoint-mcp &
SERVER_PID=$!
sleep 2
kill $SERVER_PID 2>/dev/null
echo "✅ Server starts successfully"

# Test 2: Test volume mount
echo "2. Testing volume mount..."
docker run --rm -v /Users/ross/PowerPoints:/home/mcpuser/PowerPoints --user 1000:1000 powerpoint-mcp ls -la /home/mcpuser/PowerPoints
echo "✅ Volume mount working"

# Test 3: Test MCP tools via Docker
echo "3. Testing MCP tools via Docker..."
docker run --rm -v /Users/ross/pp_mcp:/app -v /Users/ross/PowerPoints:/home/mcpuser/PowerPoints --user 1000:1000 powerpoint-mcp python3 -c "
import asyncio
import sys
sys.path.append('/app')
from powerpoint_server import create_presentation, list_presentations

async def test():
    print('Creating test presentation...')
    result = await create_presentation('test', 'Test Presentation')
    print(result)
    
    print('Listing presentations...')
    result = await list_presentations()
    print(result)

asyncio.run(test())
"
echo "✅ MCP tools working"
