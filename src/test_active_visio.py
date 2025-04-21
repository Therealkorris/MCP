"""
Test script for accessing active Visio documents.
This script demonstrates how to connect to an already running Visio instance
and access the currently open document.
"""

import os
import sys
import json
import logging
import traceback

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)

# Add the src directory to the path if needed
sys.path.append(os.path.abspath(os.path.dirname(__file__)))

# Import the Visio service
from services.visio_service import VisioService

def test_active_document():
    """Test accessing the active Visio document."""
    logger.info("Testing access to active Visio document...")
    
    try:
        # Create Visio service
        visio_service = VisioService()
        
        # Check if connected
        if not visio_service.is_connected:
            logger.error("Failed to connect to Visio")
            return False
        
        logger.info("Successfully connected to Visio")
        
        # Get active document
        result = visio_service.get_active_document()
        
        # Check for errors
        if "error" in result:
            logger.error(f"Error: {result['error']}")
            return False
        
        # Check if there's an active document
        if result.get("status") == "no_document":
            logger.warning("No Visio document is currently open")
            return False
        
        # Print document info
        logger.info(f"Active document: {result['name']}")
        logger.info(f"Path: {result['full_path']}")
        logger.info(f"Pages: {result['pages_count']}")
        logger.info(f"Saved: {result['saved']}")
        logger.info(f"Read-only: {result['readonly']}")
        
        # Print info about all open documents
        logger.info("\nAll open documents:")
        for doc in result["open_documents"]:
            status = "(ACTIVE)" if doc["is_active"] else ""
            logger.info(f"- {doc['name']} {status}")
        
        # Print page info
        logger.info("\nPages in active document:")
        for page in result["pages"]:
            logger.info(f"- Page {page['index']}: {page['name']} (Shapes: {page['shapes_count']})")
        
        return True
        
    except Exception as e:
        logger.error(f"Error in test: {e}")
        logger.error(traceback.format_exc())
        return False

if __name__ == "__main__":
    success = test_active_document()
    
    if success:
        logger.info("\nActive Visio document test successful!")
        sys.exit(0)
    else:
        logger.error("\nActive Visio document test failed!")
        sys.exit(1) 