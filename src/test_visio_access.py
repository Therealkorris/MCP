"""
Test script for Visio file access and COM automation.
"""

import os
import sys
import logging
import win32com.client
import pythoncom

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)

def test_visio_file(file_path):
    """Test opening and reading a Visio file."""
    logger.info(f"Testing access to Visio file: {file_path}")
    
    if not os.path.exists(file_path):
        logger.error(f"File not found: {file_path}")
        return False
    
    logger.info(f"File exists, size: {os.path.getsize(file_path)} bytes")
    
    try:
        # Initialize COM
        logger.info("Initializing COM...")
        pythoncom.CoInitialize()
        
        # Create Visio application object
        logger.info("Creating Visio application...")
        visio = win32com.client.Dispatch("Visio.Application")
        
        # Make Visio invisible
        visio.Visible = 0
        
        # Open the document
        logger.info(f"Opening document: {file_path}")
        doc = visio.Documents.Open(file_path)
        
        # Get document properties
        logger.info(f"Document name: {doc.Name}")
        logger.info(f"Document pages: {doc.Pages.Count}")
        
        # Get shapes from the first page
        page = doc.Pages.Item(1)
        logger.info(f"Page name: {page.Name}")
        logger.info(f"Page shapes: {page.Shapes.Count}")
        
        # List shape names
        for i in range(1, page.Shapes.Count + 1):
            shape = page.Shapes.Item(i)
            logger.info(f"Shape {i}: {shape.Name}, Text: {shape.Text}")
        
        # Close document
        doc.Close()
        
        # Quit Visio
        visio.Quit()
        
        return True
        
    except Exception as e:
        logger.error(f"Error accessing Visio file: {e}")
        return False
    
    finally:
        # Uninitialize COM
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    # Default test file path
    test_file = r"C:\visio-files\test.vsdx"
    
    # Use command line argument if provided
    if len(sys.argv) > 1:
        test_file = sys.argv[1]
    
    success = test_visio_file(test_file)
    
    if success:
        logger.info("Visio access test successful!")
        sys.exit(0)
    else:
        logger.error("Visio access test failed!")
        sys.exit(1) 