#!/usr/bin/env python3
"""
Test script for MCP server
"""
import subprocess
import json
import sys
import time

def test_mcp_server():
    """Test the MCP server functionality"""
    print("Testing MCP server...")
    
    # Start the server
    process = subprocess.Popen(
        [sys.executable, "server.py", "--mcp"],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True
    )
    
    try:
        # Test initialize
        init_request = {
            "jsonrpc": "2.0",
            "id": 1,
            "method": "initialize",
            "params": {
                "protocolVersion": "2024-11-05",
                "capabilities": {},
                "clientInfo": {"name": "test-client", "version": "1.0.0"}
            }
        }
        
        print("Sending initialize request...")
        process.stdin.write(json.dumps(init_request) + "\n")
        process.stdin.flush()
        
        # Read response
        response_line = process.stdout.readline()
        if response_line:
            response = json.loads(response_line.strip())
            print(f"Initialize response: {json.dumps(response, indent=2)}")
            
            if response.get("result"):
                print("✅ Initialize successful")
            else:
                print("❌ Initialize failed")
                return False
        else:
            print("❌ No response received")
            return False
        
        # Test tools/list
        tools_request = {
            "jsonrpc": "2.0",
            "id": 2,
            "method": "tools/list",
            "params": {}
        }
        
        print("\nSending tools/list request...")
        process.stdin.write(json.dumps(tools_request) + "\n")
        process.stdin.flush()
        
        # Read response
        response_line = process.stdout.readline()
        if response_line:
            response = json.loads(response_line.strip())
            print(f"Tools list response: {json.dumps(response, indent=2)}")
            
            if response.get("result") and "tools" in response["result"]:
                tools = response["result"]["tools"]
                print(f"✅ Found {len(tools)} tools")
                for tool in tools:
                    print(f"  - {tool['name']}: {tool['description']}")
            else:
                print("❌ Tools list failed")
                return False
        else:
            print("❌ No response received")
            return False
            
        print("\n✅ MCP server test completed successfully!")
        return True
        
    except Exception as e:
        print(f"❌ Test failed with error: {e}")
        return False
    finally:
        process.terminate()
        process.wait()

if __name__ == "__main__":
    success = test_mcp_server()
    sys.exit(0 if success else 1)
