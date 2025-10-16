#!/usr/bin/env python3
"""
Test script for PowerPoint MCP Server
"""
import asyncio
import json
import sys
from powerpoint_server import create_presentation, add_slide, list_presentations


async def test_powerpoint_server():
    """Test the PowerPoint server functions"""
    print("🧪 Testing PowerPoint MCP Server")
    print("=" * 40)

    # Test 1: Create presentation
    print("\n1. Creating presentation...")
    result = await create_presentation("MCP-servers", "MCP Servers")
    print(f"Result: {result}")

    # Test 2: List presentations
    print("\n2. Listing presentations...")
    result = await list_presentations()
    print(f"Result: {result}")

    # Test 3: Add slide
    print("\n3. Adding slide...")
    result = await add_slide(
        "MCP-servers", "What is MCP?", "MCP stands for Model Context Protocol..."
    )
    print(f"Result: {result}")

    # Test 4: List presentations again
    print("\n4. Listing presentations again...")
    result = await list_presentations()
    print(f"Result: {result}")


if __name__ == "__main__":
    asyncio.run(test_powerpoint_server())
