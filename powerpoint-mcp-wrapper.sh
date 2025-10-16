#!/bin/bash
# PowerPoint MCP Server Wrapper
# This script runs the PowerPoint server with proper volume mounting

docker run --rm -i \
  -v /Users/ross/PowerPoints:/home/mcpuser/PowerPoints \
  --user 1000:1000 \
  powerpoint-mcp
