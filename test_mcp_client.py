import requests
import json
import time
import uuid

def generate_message_id():
    """Generate a unique message ID."""
    return str(uuid.uuid4())

def send_mcp_request(url, method, params=None):
    """Send an MCP JSON-RPC request to the server."""
    if params is None:
        params = {}
    
    # Create the JSON-RPC message
    message = {
        "jsonrpc": "2.0",
        "id": generate_message_id(),
        "method": method,
        "params": params
    }
    
    print(f"Sending request: {json.dumps(message, indent=2)}")
    
    # Send the request
    response = requests.post(url, json=message)
    
    print(f"Response status: {response.status_code}")
    if response.status_code == 200:
        print(f"Response: {json.dumps(response.json(), indent=2)}")
        return response.json()
    else:
        print(f"Error: {response.text}")
        return None

def main():
    """Main function to test MCP client functionality."""
    # The URL of the MCP server
    url = "http://localhost:8050/mcp"
    
    # Get client info (initial handshake)
    print("\n=== Testing get_client_info ===")
    client_info_response = send_mcp_request(url, "get_client_info", {
        "client_info": {
            "name": "MCP Test Client",
            "version": "1.0.0"
        }
    })
    
    if not client_info_response:
        print("Failed to connect to MCP server")
        return
    
    # Ping the server
    print("\n=== Testing ping ===")
    ping_response = send_mcp_request(url, "ping")
    
    if not ping_response:
        print("Failed to ping MCP server")
        return
    
    print("\nAll tests completed successfully!")

if __name__ == "__main__":
    main() 