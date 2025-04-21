import requests
import json
import sys

def test_sse_connection(url):
    """Test if the SSE connection to the MCP server works."""
    print(f"Testing connection to {url}...")
    
    try:
        # Send a simple GET request to establish SSE connection
        response = requests.get(url, stream=True, timeout=5)
        
        print(f"Response status code: {response.status_code}")
        print(f"Response headers: {json.dumps(dict(response.headers), indent=2)}")
        
        if response.status_code == 200:
            print("Connection successful! Receiving events:")
            count = 0
            for line in response.iter_lines():
                if line:
                    print(f"Received: {line.decode('utf-8')}")
                    count += 1
                    if count >= 5:  # Just get a few events
                        break
            print("Test completed successfully.")
            return True
        else:
            print(f"Failed to connect: {response.text}")
            return False
            
    except Exception as e:
        print(f"Error connecting to the server: {e}")
        return False

if __name__ == "__main__":
    # Default URL
    url = "http://localhost:8050/mcp"
    
    # Use command line argument if provided
    if len(sys.argv) > 1:
        url = sys.argv[1]
        
    test_sse_connection(url) 