#!/usr/bin/env python3
"""
CM5 Runner Script
Author: PamirAI
Description: Connects to CM5 via SSH and runs the specified command once
"""

import paramiko
import time
import signal
import sys
from datetime import datetime

# Configuration Constants
CM5_IP_ADDRESS = "distiller@192.168.0.30"
CM5_PASSWORD = "one"
CM5_COMMAND = "cd /opt/distiller-cm5-python/ && ./run.sh"

class CM5Runner:
    def __init__(self):
        self.ssh = None
        self.running = True
        
        # Setup signal handler for graceful shutdown
        signal.signal(signal.SIGINT, self.signal_handler)
        signal.signal(signal.SIGTERM, self.signal_handler)
    
    def signal_handler(self, sig, frame):
        """Handle Ctrl+C and other termination signals"""
        print(f"\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Received termination signal. Shutting down gracefully...")
        self.running = False
        if self.ssh:
            try:
                self.ssh.close()
                print("SSH connection closed.")
            except:
                pass
        sys.exit(0)
    
    def connect_ssh(self):
        """Establish SSH connection to CM5"""
        try:
            ssh = paramiko.SSHClient()
            ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            
            user, host = CM5_IP_ADDRESS.split('@')
            
            print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Connecting to {CM5_IP_ADDRESS}...")
            ssh.connect(host, username=user, password=CM5_PASSWORD, timeout=30)
            print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] SSH connection established successfully!")
            return ssh
            
        except Exception as e:
            print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] SSH connection failed: {e}")
            return None
    
    def execute_command(self, ssh, command):
        """Execute command via SSH and monitor output"""
        try:
            print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Executing command: {command}")
            
            # Use invoke_shell for interactive commands that need to stay running
            channel = ssh.invoke_shell()
            time.sleep(1)  # Wait for shell to be ready
            
            # Send the command
            channel.send(command + '\n')
            
            # Monitor output until command finishes or connection is lost
            while self.running:
                try:
                    if channel.recv_ready():
                        output = channel.recv(1024).decode('utf-8')
                        if output:
                            print(output, end='', flush=True)
                    
                    if channel.exit_status_ready():
                        exit_status = channel.recv_exit_status()
                        print(f"\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Command exited with status: {exit_status}")
                        return True
                    
                    # Check if channel is still active
                    if channel.closed:
                        print(f"\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] SSH channel closed.")
                        return False
                    
                    time.sleep(0.1)  # Small delay to prevent busy waiting
                    
                except (IOError, OSError, EOFError) as e:
                    # Connection lost or device unplugged
                    print(f"\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Connection lost: {e}")
                    return False
                    
        except Exception as e:
            print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Error executing command: {e}")
            return False
    
    def run(self):
        """Main run function"""
        print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Starting CM5 Runner...")
        print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Target: {CM5_IP_ADDRESS}")
        print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Command: {CM5_COMMAND}")
        print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Press Ctrl+C to stop\n")
        
        try:
            # Connect to SSH
            self.ssh = self.connect_ssh()
            if not self.ssh:
                print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Failed to establish SSH connection. Exiting.")
                return
            
            # Execute the command once
            success = self.execute_command(self.ssh, CM5_COMMAND)
            
            if success:
                print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Command completed successfully.")
            else:
                print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Command failed or connection was lost.")
            
            # Keep the connection open and wait
            print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Keeping SSH connection open. Press Ctrl+C to exit.")
            while self.running:
                time.sleep(1)
                
        except KeyboardInterrupt:
            print(f"\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Keyboard interrupt received.")
        except Exception as e:
            print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Unexpected error: {e}")
        finally:
            # Cleanup
            if self.ssh:
                try:
                    self.ssh.close()
                    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] SSH connection closed.")
                except:
                    pass
            
            print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] CM5 Runner finished.")

def main():
    """Main function"""
    runner = CM5Runner()
    runner.run()

if __name__ == "__main__":
    main() 