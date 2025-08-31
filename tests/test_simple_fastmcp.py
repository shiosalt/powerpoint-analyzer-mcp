#!/usr/bin/env python3
"""
Test the simple FastMCP server
"""

import json
import subprocess
import sys
import time
import threading

def read_output(process, output_list, output_type):
    """Read output from process in a separate thread."""
    try:
        while True:
            if output_type == "stdout":
                line = process.stdout.readline()
            else:
                line = process.stderr.readline()
            
            if not line:
                break
                
            output_list.append(line.strip())
    except Exception as e:
        print(f"Error reading {output_type}: {e}")

def test_simple_fastmcp():
    """Test the simple FastMCP server."""
    print("Testing Simple FastMCP Server")
    print("=" * 40)
    
    # Start the server
    process = subprocess.Popen(
        [sys.executable, "simple_fastmcp_test.py"],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding='utf-8'
    )
    
    stdout_lines = []
    stderr_lines = []
    
    # Start threads to read output
    stdout_thread = threading.Thread(target=read_output, args=(process, stdout_lines, "stdout"))
    stderr_thread = threading.Thread(target=read_output, args=(process, stderr_lines, "stderr"))
    
    stdout_thread.daemon = True
    stderr_thread.daemon = True
    
    stdout_thread.start()
    stderr_thread.start()
    
    try:
        # Wait for server to start
        time.sleep(3)
        
        # Check if server is still running
        if process.poll() is not None:
            print(f"❌ Server terminated with code: {process.returncode}")
            print(f"Stdout: {stdout_lines}")
            print(f"Stderr: {stderr_lines}")
            return False
        
        print("✓ Server started successfully")
        
        # Test initialize
        request = {
            "jsonrpc": "2.0",
            "id": 1,
            "method": "initialize",
            "params": {
                "protocolVersion": "2024-11-05",
                "capabilities": {"tools": {}},
                "clientInfo": {"name": "test-client", "version": "1.0.0"}
            }
        }
        request_json = json.dumps(request) + "\n"
        
        print(f"\nSending initialize request...")
        print(f"Request: {request_json.strip()}")
        
        # Send request
        process.stdin.write(request_json)
        process.stdin.flush()
        
        # Wait for response
        time.sleep(2)
        
        # Check for response in stdout
        if stdout_lines:
            print(f"Stdout responses: {stdout_lines}")
            
            # Try to find JSON response
            for line in stdout_lines:
                if line.strip().startswith('{'):
                    try:
                        response = json.loads(line.strip())
                        if response.get("id") == 1 and "result" in response:
                            print("✓ Initialize successful")
                            
                            # Send notifications/initialized (no ID, no response expected)
                            notification = {"jsonrpc": "2.0", "method": "notifications/initialized"}
                            notification_json = json.dumps(notification) + "\n"
                            
                            print(f"\nSending notifications/initialized...")
                            print(f"Notification: {notification_json.strip()}")
                            
                            process.stdin.write(notification_json)
                            process.stdin.flush()
                            
                            # Wait a moment for notification processing
                            time.sleep(1)
                            
                            # Test tools/list (use same ID as initialize)
                            request2 = {"jsonrpc": "2.0", "id": 1, "method": "tools/list"}
                            request2_json = json.dumps(request2) + "\n"
                            
                            print(f"\nSending tools/list request...")
                            print(f"Request: {request2_json.strip()}")
                            
                            process.stdin.write(request2_json)
                            process.stdin.flush()
                            
                            # Wait for tools/list response
                            time.sleep(2)
                            
                            # Check for tools/list response in all stdout lines
                            print(f"All stdout lines after tools/list: {stdout_lines}")
                            for line2 in stdout_lines:
                                if line2.strip().startswith('{'):
                                    try:
                                        response2 = json.loads(line2.strip())
                                        if response2.get("id") == 1 and "result" in response2 and "tools" in response2["result"]:
                                            tools = response2["result"]["tools"]
                                            print(f"✓ TOOLS/LIST test passed! ({len(tools)} tools)")
                                            
                                            # Print tool names
                                            tool_names = [tool.get("name", "unknown") for tool in tools]
                                            print(f"Tools: {', '.join(tool_names)}")
                                            
                                            return True
                                    except json.JSONDecodeError:
                                        continue
                            
                            print("❌ TOOLS/LIST test failed: No valid response found")
                            return False
                        else:
                            print(f"❌ Initialize failed: {response}")
                            return False
                    except json.JSONDecodeError:
                        continue
            
            print("❌ Initialize failed: No valid JSON response found")
            return False
        else:
            print("❌ Initialize failed: No stdout output")
            return False
            
    finally:
        process.terminate()
        process.wait()
        print("\nServer stopped")

if __name__ == "__main__":
    success = test_simple_fastmcp()
    if success:
        print("\n🎉 Simple FastMCP server is working!")
    else:
        print("\n❌ Simple FastMCP server has issues")
    sys.exit(0 if success else 1)