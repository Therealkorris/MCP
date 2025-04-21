"""
Test script for MCP-Visio connection

This script tests the connection to the MCP-Visio server and performs
basic operations on a Visio diagram.
"""

import sys
import json
import requests
from pprint import pprint

# MCP server URL
MCP_URL = "http://localhost:8050"

def print_section(title):
    """Print a section title."""
    print("\n" + "=" * 50)
    print(f" {title}")
    print("=" * 50 + "\n")

def test_get_active_document():
    """Test getting the active Visio document."""
    print_section("Testing Get Active Document")
    
    try:
        # Prepare request
        request = {
            "jsonrpc": "2.0",
            "id": 1,
            "method": "get_active_document",
            "params": {}
        }
        
        # Send request
        response = requests.post(MCP_URL, json=request)
        
        # Check response
        if response.status_code == 200:
            result = response.json()
            print("Response received:")
            pprint(result)
            
            if "result" in result and result["result"].get("status") == "success":
                print("\n✅ Active document found!")
                return True, result["result"]
            else:
                print("\n❌ No active document found")
                return False, None
        else:
            print(f"\n❌ Error: {response.status_code} - {response.text}")
            return False, None
    
    except Exception as e:
        print(f"\n❌ Exception: {e}")
        return False, None

def test_analyze_diagram(file_path="active"):
    """Test analyzing a Visio diagram."""
    print_section(f"Testing Analyze Diagram: {file_path}")
    
    try:
        # Prepare request
        request = {
            "jsonrpc": "2.0",
            "id": 2,
            "method": "analyze_visio_diagram",
            "params": {
                "file_path": file_path,
                "analysis_type": "all"
            }
        }
        
        # Send request
        response = requests.post(MCP_URL, json=request)
        
        # Check response
        if response.status_code == 200:
            result = response.json()
            print("Response received:")
            
            # Print only summary to avoid too much output
            if "result" in result:
                if result["result"].get("status") == "success":
                    print("\n✅ Diagram analyzed successfully!")
                    
                    # Print summary
                    pages = result["result"].get("pages", [])
                    print(f"\nDocument: {result['result'].get('name', 'Unknown')}")
                    print(f"Pages: {len(pages)}")
                    
                    for i, page in enumerate(pages):
                        print(f"\nPage {i+1}: {page.get('name', 'Unnamed')}")
                        print(f"  Shapes: {len(page.get('shapes', []))}")
                        print(f"  Connections: {len(page.get('connections', []))}")
                        print(f"  Text Elements: {len(page.get('text_elements', []))}")
                    
                    return True, result["result"]
                else:
                    print("\n❌ Analysis failed:")
                    pprint(result["result"])
                    return False, None
            else:
                print("\n❌ Invalid response:")
                pprint(result)
                return False, None
        else:
            print(f"\n❌ Error: {response.status_code} - {response.text}")
            return False, None
    
    except Exception as e:
        print(f"\n❌ Exception: {e}")
        return False, None

def test_modify_diagram(file_path="active", operation="add_shape"):
    """Test modifying a Visio diagram."""
    print_section(f"Testing Modify Diagram: {operation}")
    
    try:
        # Prepare shape data based on operation
        shape_data = {}
        
        if operation == "add_shape":
            shape_data = {
                "master_name": "Rectangle",
                "stencil_name": "BASIC_M.vssx",
                "position": {"x": 4.0, "y": 4.0},
                "size": {"width": 2.0, "height": 1.0},
                "text": "Test Shape",
                "fill_color": "#00FF00",
                "line_color": "#000000"
            }
        elif operation == "add_connector":
            # First, we need to get shapes to connect
            # This is just a placeholder - in a real test, you'd need actual shape IDs
            shape_data = {
                "from_shape_id": 1,  # Replace with actual shape ID
                "to_shape_id": 2,    # Replace with actual shape ID
                "connector_type": "straight",
                "text": "Test Connection",
                "line_color": "#FF0000"
            }
        
        # Prepare request
        request = {
            "jsonrpc": "2.0",
            "id": 3,
            "method": "modify_visio_diagram",
            "params": {
                "file_path": file_path,
                "operation": operation,
                "shape_data": shape_data
            }
        }
        
        # Send request
        response = requests.post(MCP_URL, json=request)
        
        # Check response
        if response.status_code == 200:
            result = response.json()
            print("Response received:")
            pprint(result)
            
            if "result" in result and result["result"].get("status") == "success":
                print(f"\n✅ {operation} operation successful!")
                return True, result["result"]
            else:
                print(f"\n❌ {operation} operation failed")
                return False, None
        else:
            print(f"\n❌ Error: {response.status_code} - {response.text}")
            return False, None
    
    except Exception as e:
        print(f"\n❌ Exception: {e}")
        return False, None

def test_verify_connections(file_path="active", shape_ids=None):
    """Test verifying connections in a Visio diagram."""
    print_section("Testing Verify Connections")
    
    try:
        # Prepare request
        request = {
            "jsonrpc": "2.0",
            "id": 4,
            "method": "verify_connections",
            "params": {
                "file_path": file_path,
                "shape_ids": shape_ids or []
            }
        }
        
        # Send request
        response = requests.post(MCP_URL, json=request)
        
        # Check response
        if response.status_code == 200:
            result = response.json()
            print("Response received:")
            pprint(result)
            
            if "result" in result and result["result"].get("status") == "success":
                print("\n✅ Connections verified successfully!")
                print(f"Found {len(result['result'].get('data', {}).get('connections', []))} connections")
                return True, result["result"]
            else:
                print("\n❌ Connection verification failed")
                return False, None
        else:
            print(f"\n❌ Error: {response.status_code} - {response.text}")
            return False, None
    
    except Exception as e:
        print(f"\n❌ Exception: {e}")
        return False, None

def test_get_available_stencils():
    """Test getting available Visio stencils."""
    print_section("Testing Get Available Stencils")
    
    try:
        # Prepare request
        request = {
            "jsonrpc": "2.0",
            "id": 5,
            "method": "get_available_stencils",
            "params": {}
        }
        
        # Send request
        response = requests.post(MCP_URL, json=request)
        
        # Check response
        if response.status_code == 200:
            result = response.json()
            print("Response received:")
            
            if "result" in result and result["result"].get("status") == "success":
                data = result["result"].get("data", {})
                built_in_stencils = data.get("built_in_stencils", [])
                open_stencils = data.get("open_stencils", [])
                
                print("\n✅ Stencils retrieved successfully!")
                print(f"Built-in stencils: {len(built_in_stencils)}")
                print(f"Open stencils: {len(open_stencils)}")
                
                # Print some examples
                if built_in_stencils:
                    print("\nSample built-in stencils:")
                    for i, stencil in enumerate(built_in_stencils[:5]):
                        print(f"  - {stencil.get('name')}")
                    if len(built_in_stencils) > 5:
                        print(f"  ... and {len(built_in_stencils) - 5} more")
                
                return True, result["result"]
            else:
                print("\n❌ Failed to retrieve stencils")
                pprint(result)
                return False, None
        else:
            print(f"\n❌ Error: {response.status_code} - {response.text}")
            return False, None
    
    except Exception as e:
        print(f"\n❌ Exception: {e}")
        return False, None

def test_get_shapes_on_page(file_path="active", page_index=1):
    """Test getting shapes on a page."""
    print_section(f"Testing Get Shapes on Page {page_index}")
    
    try:
        # Prepare request
        request = {
            "jsonrpc": "2.0",
            "id": 6,
            "method": "get_shapes_on_page",
            "params": {
                "file_path": file_path,
                "page_index": page_index
            }
        }
        
        # Send request
        response = requests.post(MCP_URL, json=request)
        
        # Check response
        if response.status_code == 200:
            result = response.json()
            print("Response received:")
            
            if "result" in result and result["result"].get("status") == "success":
                data = result["result"].get("data", {})
                shapes = data.get("shapes", [])
                
                print("\n✅ Shapes retrieved successfully!")
                print(f"Document: {data.get('document_name')}")
                print(f"Page: {data.get('page_name')}")
                print(f"Total shapes: {len(shapes)}")
                
                # Print some shape info
                if shapes:
                    print("\nSample shapes:")
                    for i, shape in enumerate(shapes[:5]):
                        shape_type = "Connector" if shape.get("is_connector") else "Shape"
                        shape_text = shape.get("text", "")
                        shape_text_preview = shape_text[:20] + "..." if len(shape_text) > 20 else shape_text
                        print(f"  - [{shape.get('id')}] {shape.get('name')} ({shape_type}): {shape_text_preview}")
                    if len(shapes) > 5:
                        print(f"  ... and {len(shapes) - 5} more")
                
                return True, result["result"]
            else:
                print("\n❌ Failed to retrieve shapes")
                pprint(result)
                return False, None
        else:
            print(f"\n❌ Error: {response.status_code} - {response.text}")
            return False, None
    
    except Exception as e:
        print(f"\n❌ Exception: {e}")
        return False, None

def test_create_new_diagram(template="Basic.vst", save_path=None):
    """Test creating a new Visio diagram."""
    print_section("Testing Create New Diagram")
    
    try:
        # Prepare request
        request = {
            "jsonrpc": "2.0",
            "id": 7,
            "method": "create_new_diagram",
            "params": {
                "template": template
            }
        }
        
        if save_path:
            request["params"]["save_path"] = save_path
        
        # Send request
        response = requests.post(MCP_URL, json=request)
        
        # Check response
        if response.status_code == 200:
            result = response.json()
            print("Response received:")
            pprint(result)
            
            if "result" in result and result["result"].get("status") == "success":
                print("\n✅ New diagram created successfully!")
                doc_data = result["result"].get("data", {})
                print(f"Document name: {doc_data.get('name')}")
                print(f"Path: {doc_data.get('full_path')}")
                return True, result["result"]
            else:
                print("\n❌ Failed to create new diagram")
                return False, None
        else:
            print(f"\n❌ Error: {response.status_code} - {response.text}")
            return False, None
    
    except Exception as e:
        print(f"\n❌ Exception: {e}")
        return False, None

def test_save_diagram(file_path="active"):
    """Test saving a Visio diagram."""
    print_section("Testing Save Diagram")
    
    try:
        # Prepare request
        request = {
            "jsonrpc": "2.0",
            "id": 8,
            "method": "save_diagram",
            "params": {
                "file_path": file_path
            }
        }
        
        # Send request
        response = requests.post(MCP_URL, json=request)
        
        # Check response
        if response.status_code == 200:
            result = response.json()
            print("Response received:")
            pprint(result)
            
            if "result" in result and result["result"].get("status") == "success":
                print("\n✅ Diagram saved successfully!")
                doc_data = result["result"].get("data", {})
                print(f"Document name: {doc_data.get('name')}")
                print(f"Path: {doc_data.get('full_path')}")
                return True, result["result"]
            else:
                print("\n❌ Failed to save diagram")
                return False, None
        else:
            print(f"\n❌ Error: {response.status_code} - {response.text}")
            return False, None
    
    except Exception as e:
        print(f"\n❌ Exception: {e}")
        return False, None

def test_export_diagram(file_path="active", format="png", output_path=None):
    """Test exporting a Visio diagram."""
    print_section(f"Testing Export Diagram to {format.upper()}")
    
    try:
        # Prepare request
        request = {
            "jsonrpc": "2.0",
            "id": 9,
            "method": "export_diagram",
            "params": {
                "file_path": file_path,
                "format": format
            }
        }
        
        if output_path:
            request["params"]["output_path"] = output_path
        
        # Send request
        response = requests.post(MCP_URL, json=request)
        
        # Check response
        if response.status_code == 200:
            result = response.json()
            print("Response received:")
            pprint(result)
            
            if "result" in result and result["result"].get("status") == "success":
                print("\n✅ Diagram exported successfully!")
                export_data = result["result"].get("data", {})
                print(f"Document name: {export_data.get('document_name')}")
                print(f"Format: {export_data.get('format')}")
                print(f"Output path: {export_data.get('output_path')}")
                return True, result["result"]
            else:
                print("\n❌ Failed to export diagram")
                return False, None
        else:
            print(f"\n❌ Error: {response.status_code} - {response.text}")
            return False, None
    
    except Exception as e:
        print(f"\n❌ Exception: {e}")
        return False, None

def main():
    """Main test function."""
    print_section("MCP-Visio Connection Test")
    print("This test will check if the MCP-Visio server is working correctly.")
    print("Make sure you have both the MCP server and Visio relay running.")
    print("Also ensure you have a Visio document open in Visio.")
    
    tests_passed = 0
    tests_failed = 0
    
    # Test active document
    success, doc_data = test_get_active_document()
    if success:
        tests_passed += 1
    else:
        tests_failed += 1
        print("⚠️ Active document test failed - some other tests may be skipped")
    
    # If there's an active document, run the other tests
    if success:
        # Test analyzing active document
        success, _ = test_analyze_diagram()
        tests_passed += 1 if success else 0
        tests_failed += 0 if success else 1
        
        # Test adding a shape
        success, shape_data = test_modify_diagram()
        tests_passed += 1 if success else 0
        tests_failed += 0 if success else 1
        
        # Test verifying connections
        success, _ = test_verify_connections()
        tests_passed += 1 if success else 0
        tests_failed += 0 if success else 1
        
        # Test getting available stencils
        success, _ = test_get_available_stencils()
        tests_passed += 1 if success else 0
        tests_failed += 0 if success else 1
        
        # Test getting shapes on a page
        success, _ = test_get_shapes_on_page()
        tests_passed += 1 if success else 0
        tests_failed += 0 if success else 1
        
        # Test saving diagram
        success, _ = test_save_diagram()
        tests_passed += 1 if success else 0
        tests_failed += 0 if success else 1
        
        # Skip export test by default to avoid creating files
        # Uncomment this section to test exporting
        # success, _ = test_export_diagram()
        # tests_passed += 1 if success else 0
        # tests_failed += 0 if success else 1
    
    # Always test creating a new diagram, but don't save it
    success, _ = test_create_new_diagram()
    tests_passed += 1 if success else 0
    tests_failed += 0 if success else 1
    
    # Summary
    print_section("Test Summary")
    print(f"Tests passed: {tests_passed}")
    print(f"Tests failed: {tests_failed}")
    print(f"Total tests: {tests_passed + tests_failed}")
    
    if tests_failed == 0:
        print("\n✅ All tests passed successfully!")
    else:
        print(f"\n⚠️ {tests_failed} tests failed")

if __name__ == "__main__":
    main() 