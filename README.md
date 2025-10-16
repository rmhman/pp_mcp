# PowerPoint MCP Server

A Model Context Protocol (MCP) server that creates and manages PowerPoint presentations locally on your computer.

## Purpose

This MCP server provides a secure interface for AI assistants to create, modify, and manage PowerPoint (.pptx) files in your local ~/PowerPoints directory.

## Features

### Current Implementation

- **`create_presentation`** - Create a new PowerPoint presentation with optional title slide
- **`add_slide`** - Add new slides with title and content to existing presentations
- **`list_presentations`** - List all PowerPoint files in the ~/PowerPoints directory
- **`get_presentation_info`** - Get detailed information about a specific presentation

## Prerequisites

- Docker Desktop with MCP Toolkit enabled
- Docker MCP CLI plugin (`docker mcp` command)

## Installation

### For Docker MCP CLI (Recommended)

1. **Build the Docker image:**

   ```bash
   docker build -t powerpoint-mcp .
   ```

2. **Configure Claude Desktop** with volume mounting:

   Add this to your Claude Desktop MCP configuration:
   ```json
   {
     "mcpServers": {
       "powerpoint": {
         "command": "docker",
         "args": [
           "run",
           "--rm",
           "-i",
           "-v",
           "/Users/ross/PowerPoints:/home/mcpuser/PowerPoints",
           "--user",
           "1000:1000",
           "powerpoint-mcp"
         ]
       }
     }
   }
   ```
   
   **Important**: Use the full absolute path `/Users/ross/PowerPoints` instead of `~/PowerPoints` because Docker doesn't expand `~` when passed through JSON/CLI arguments.

### Alternative: Persistent Container Setup

1. **Run the setup script:**
   ```bash
   ./setup.sh
   ```

2. **Configure Claude Desktop** to use the persistent container

### Manual Setup

1. **Create PowerPoint directory:**
   ```bash
   mkdir -p ~/PowerPoints
   ```

2. **Build Docker image:**
   ```bash
   docker build -t powerpoint-mcp .
   ```

3. **Run with volume mount:**
   ```bash
   docker run -d \
     --name powerpoint-mcp-server \
     -v ~/PowerPoints:/home/mcpuser/PowerPoints \
     -e PYTHONUNBUFFERED=1 \
     --user 1000:1000 \
     powerpoint-mcp
   ```

### Docker Compose (Alternative)

```bash
docker-compose up -d
```

## VS Code Integration

To use this MCP server with VS Code:

1. **Install MCP Toolkit extension** in VS Code
2. **Add server configuration** to your VS Code settings (`settings.json`):

```json
{
  "mcp.servers": {
    "powerpoint": {
      "command": "docker",
      "args": [
        "run",
        "--rm",
        "-i",
        "-v",
        "/Users/ross/PowerPoints:/home/mcpuser/PowerPoints",
        "--user",
        "1000:1000",
        "powerpoint-mcp"
      ],
      "type": "stdio"
    }
  }
}
```

3. **Restart VS Code** to load the MCP server
4. **Use the PowerPoint tools** through the MCP Toolkit interface

## Usage Examples

In Claude Desktop, you can ask:

- "Create a new PowerPoint presentation called 'My Demo' with the title 'Introduction to MCP'"
- "Add a slide to 'My Demo' with title 'What is MCP?' and content 'MCP stands for Model Context Protocol...'"
- "List all my PowerPoint presentations"
- "Show me information about the 'My Demo' presentation"

## Architecture

```
Claude Desktop → MCP Gateway → PowerPoint MCP Server → Local File System
↓
~/PowerPoints Directory
```

## Development

### Local Testing

```bash
# Build Docker image (only needed the first time)
docker build -t powerpoint-mcp .

# Test with volume mount
docker run --rm -i -v /Users/ross/PowerPoints:/home/mcpuser/PowerPoints --user 1000:1000 powerpoint-mcp

# Test volume mount
docker run --rm -v /Users/ross/PowerPoints:/home/mcpuser/PowerPoints --user 1000:1000 powerpoint-mcp ls -la /home/mcpuser/PowerPoints

# Test MCP tools directly (requires Python)
./test_docker.sh
```

### Adding New Tools

1. Add the function to `powerpoint_server.py`
2. Decorate with `@mcp.tool()`
3. Update the catalog entry with the new tool name
4. Rebuild the Docker image

## Troubleshooting

### Presentations Not Persisting

**Problem:** Presentations are created but disappear when trying to add slides.

**Solution:** This is caused by missing volume mounts. Ensure you're running with:
```bash
docker run -v ~/PowerPoints:/home/mcpuser/PowerPoints ...
```

Or use the provided setup script: `./setup.sh`

### Tools Not Appearing

- Verify Docker image built successfully
- Check catalog and registry files
- Ensure Claude Desktop config includes custom catalog
- Restart Claude Desktop

### File Permission Errors

- Ensure ~/PowerPoints directory is writable
- Check Docker volume permissions
- Verify container is running as user 1000:1000

### Container Issues

- Check container logs: `docker logs powerpoint-mcp-server`
- Verify container is running: `docker ps --filter name=powerpoint-mcp-server`
- Restart container: `docker restart powerpoint-mcp-server`

## Security Considerations

- All files created in user's home directory
- No external API calls or network access
- Running as non-root user in Docker
- Safe filename handling prevents path traversal

## License

MIT License
