"""
Test client for the MCP-Visio server.
This script sends test requests to the server to verify functionality.
"""

import sys
import json
import asyncio
import httpx
import uuid
import logging
from pprint import pprint

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)

# Configure server URL
SERVER_URL = "http://localhost:8050/sse"

async def send_request(method, params=None):
    """Send a request to the MCP server."""
    # Create request payload
    request_id = str(uuid.uuid4())
    payload = {
        "jsonrpc": "2.0",
        "id": request_id,
        "method": method
    }
    
    if params:
        payload["params"] = params
    
    logger.info(f"Sending request: {method}")
    logger.debug(f"Request payload: {json.dumps(payload, indent=2)}")
    
    try:
        # Send the request to the server
        async with httpx.AsyncClient(timeout=60.0) as client:
            response = await client.post(SERVER_URL, json=payload)
            
            # Check if request was successful
            if response.status_code == 200:
                result = response.json()
                logger.info(f"Received response for: {method}")
                logger.debug(f"Response: {json.dumps(result, indent=2)}")
                return result
            else:
                logger.error(f"Server error: {response.status_code} - {response.text}")
                return None
    
    except Exception as e:
        logger.error(f"Error sending request: {e}")
        return None

async def test_get_client_info():
    """Test the get_client_info method."""
    params = {
        "client_info": {
            "name": "MCP-Visio Test Client",
            "version": "1.0.0"
        }
    }
    
    result = await send_request("get_client_info", params)
    
    if result and "result" in result:
        logger.info("get_client_info test passed")
        tools = result.get("result", {}).get("tools", [])
        logger.info(f"Available tools: {len(tools)}")
        for tool in tools:
            logger.info(f"  - {tool.get('name')}: {tool.get('description')}")
    else:
        logger.error("get_client_info test failed")

async def test_get_active_document():
    """Test the get_active_document method."""
    result = await send_request("get_active_document")
    
    if result and "result" in result:
        logger.info("get_active_document test passed")
        doc_info = result.get("result", {})
        
        if "error" in doc_info:
            logger.warning(f"No active document: {doc_info.get('error')}")
        else:
            logger.info(f"Active document: {doc_info.get('name')}")
            logger.info(f"Pages: {doc_info.get('pages_count')}")
    else:
        logger.error("get_active_document test failed")

async def test_ping():
    """Test the ping method."""
    result = await send_request("ping")
    
    if result and result.get("result") == "pong":
        logger.info("ping test passed")
    else:
        logger.error("ping test failed")

async def main():
    """Run all tests."""
    logger.info("Starting MCP-Visio server tests")
    
    # Test ping
    await test_ping()
    
    # Test get_client_info
    await test_get_client_info()
    
    # Test get_active_document
    await test_get_active_document()
    
    logger.info("Tests completed")

if __name__ == "__main__":
    # Run the tests
    asyncio.run(main()) 