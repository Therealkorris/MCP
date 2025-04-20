"""
Standard I/O transport for the MCP-Visio server.
"""

import sys
import json
import logging
import asyncio
import traceback
from typing import Dict, Any, Optional

from services.mcp_service import process_message

# Configure logging
logger = logging.getLogger(__name__)

async def read_message() -> Optional[Dict[str, Any]]:
    """Read a message from stdin."""
    try:
        line = sys.stdin.buffer.readline().decode("utf-8").strip()
        if not line:
            return None
        
        return json.loads(line)
    except json.JSONDecodeError as e:
        logger.error(f"Failed to decode JSON message: {e}")
        return None
    except Exception as e:
        logger.error(f"Error reading message: {e}")
        return None

def write_message(message: Dict[str, Any]) -> None:
    """Write a message to stdout."""
    try:
        message_json = json.dumps(message)
        sys.stdout.write(message_json + "\n")
        sys.stdout.flush()
    except Exception as e:
        logger.error(f"Error writing message: {e}")

async def stdio_loop() -> None:
    """Main loop for the STDIO transport."""
    logger.info("STDIO transport started")
    
    while True:
        try:
            # Read message from stdin
            message = await read_message()
            if message is None:
                # No message, wait a bit and try again
                await asyncio.sleep(0.1)
                continue
            
            logger.debug(f"Received message: {json.dumps(message, indent=2)}")
            
            # Process the message
            response = await process_message(message)
            
            # Write response to stdout
            write_message(response)
            
        except Exception as e:
            logger.error(f"Error in STDIO loop: {e}")
            logger.error(traceback.format_exc())
            
            # Send error response if we have a message ID
            if message and "id" in message:
                error_response = {
                    "jsonrpc": "2.0",
                    "id": message["id"],
                    "error": {
                        "code": -32603,
                        "message": f"Internal error: {str(e)}"
                    }
                }
                write_message(error_response)

def run_stdio_transport() -> None:
    """Run the STDIO transport."""
    # Set up asyncio event loop
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    
    try:
        # Run the STDIO loop
        loop.run_until_complete(stdio_loop())
    except KeyboardInterrupt:
        logger.info("STDIO transport stopped by user")
    except Exception as e:
        logger.error(f"Error in STDIO transport: {e}")
        logger.error(traceback.format_exc())
    finally:
        loop.close() 