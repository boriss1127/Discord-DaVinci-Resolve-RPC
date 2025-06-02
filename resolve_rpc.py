print("Script starting...")
import time
import psutil
import os
import sys
import winreg
import win32gui
from pypresence import Presence
import traceback

print("Imports successful")

CLIENT_ID = "1378792739767980042"

def check_dependencies():
    print("Checking dependencies...")
    try:
        import win32gui
        import psutil
        from pypresence import Presence
        print("All dependencies are installed correctly")
        return True
    except ImportError as e:
        print(f"Error: Missing required dependency - {e}")
        print("Please make sure you have installed all requirements using:")
        print("pip install -r requirements.txt")
        return False

def get_resolve_version():
    try:
        # Try to get DaVinci Resolve version from registry
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Blackmagic Design\DaVinci Resolve")
        version = winreg.QueryValueEx(key, "Version")[0]
        winreg.CloseKey(key)
        print(f"Found Resolve version: {version}")
        return version
    except Exception as e:
        print(f"Error getting Resolve version: {e}")
        print("Full error:", traceback.format_exc())
        return "Unknown Version"

def is_resolve_running():
    print("Checking for Resolve processes...")
    resolve_processes = []
    for proc in psutil.process_iter(['name', 'pid']):
        try:
            process_name = proc.info['name'].lower()
            if process_name in ['resolve.exe', 'resolve studio.exe']:
                print(f"Found Resolve process: {process_name} (PID: {proc.info['pid']})")
                resolve_processes.append(proc)
        except (psutil.NoSuchProcess, psutil.AccessDenied) as e:
            print(f"Error checking process: {e}")
            continue
    
    if not resolve_processes:
        print("No Resolve processes found")
        return False
    
    print(f"Found {len(resolve_processes)} Resolve process(es)")
    return True

def get_resolve_window_title():
    try:
        def callback(hwnd, windows):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if "DaVinci Resolve" in title:
                    print(f"Found Resolve window: {title}")
                    windows.append(title)
            return True
        
        windows = []
        win32gui.EnumWindows(callback, windows)
        
        if not windows:
            print("No Resolve windows found")
            return "DaVinci Resolve"
            
        title = windows[0]
        print(f"Using window title: {title}")
        return title
    except Exception as e:
        print(f"Error getting window title: {e}")
        print("Full error:", traceback.format_exc())
        return "DaVinci Resolve"

def main():
    if not check_dependencies():
        sys.exit(1)
        
    try:
        print("Connecting to Discord...")
        rpc = Presence(CLIENT_ID)
        
        # Test Discord connection
        try:
            rpc.connect()
            print("Successfully connected to Discord!")
        except Exception as e:
            print(f"Failed to connect to Discord: {e}")
            print("Full error:", traceback.format_exc())
            print("Please make sure:")
            print("1. Discord is running")
            print("2. The Client ID is correct")
            print("3. The application is properly set up in Discord Developer Portal")
            sys.exit(1)
        
        print("Discord RPC for DaVinci Resolve is running...")
        
        while True:
            try:
                if is_resolve_running():
                    window_title = get_resolve_window_title()
                    version = get_resolve_version()
                    
                    # Determine if it's Studio or Free version
                    is_studio = "Studio" in window_title or "Studio" in version
                    print(f"Is Studio version: {is_studio}")
                    
                    # Determine if in Project Manager or specific project
                    if "Project Manager" in window_title:
                        print("Updating status: In Project Manager")
                        try:
                            rpc.update(
                                state="In Project Manager",
                                details=f"DaVinci Resolve {'Studio' if is_studio else 'Free'}",
                                large_image="resolve_logo",
                                large_text=f"DaVinci Resolve {'Studio' if is_studio else 'Free'} {version}"
                            )
                            print("Status updated successfully")
                        except Exception as e:
                            print(f"Failed to update status: {e}")
                            print("Full error:", traceback.format_exc())
                    else:
                        # Extract project name from window title and clean it up
                        project_name = window_title.replace("DaVinci Resolve", "").replace(" - Studio", "").replace(" - Free", "").strip()
                        if project_name and project_name != "":
                            print(f"Updating status: Editing {project_name}")
                            try:
                                rpc.update(
                                    state=f"Editing: {project_name}",
                                    details=f"DaVinci Resolve {'Studio' if is_studio else 'Free'}",
                                    large_image="resolve_logo",
                                    large_text=f"DaVinci Resolve {'Studio' if is_studio else 'Free'} {version}"
                                )
                                print("Status updated successfully")
                            except Exception as e:
                                print(f"Failed to update status: {e}")
                                print("Full error:", traceback.format_exc())
                        else:
                            print("Updating status: Editing (Project Name Unknown)")
                            try:
                                rpc.update(
                                    state="Editing",
                                    details=f"DaVinci Resolve {'Studio' if is_studio else 'Free'}",
                                    large_image="resolve_logo",
                                    large_text=f"DaVinci Resolve {'Studio' if is_studio else 'Free'} {version}"
                                )
                                print("Status updated successfully")
                            except Exception as e:
                                print(f"Failed to update status: {e}")
                                print("Full error:", traceback.format_exc())
                else:
                    print("Resolve not running, clearing status")
                    try:
                        rpc.clear()
                        print("Status cleared successfully")
                    except Exception as e:
                        print(f"Failed to clear status: {e}")
                        print("Full error:", traceback.format_exc())
                
                time.sleep(15)  # Update every 15 seconds
            except Exception as e:
                print(f"Error in main loop: {e}")
                print("Full error:", traceback.format_exc())
                time.sleep(15)  # Wait before retrying
    except Exception as e:
        print(f"Fatal error: {e}")
        print("Full error:", traceback.format_exc())
        sys.exit(1)

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nExiting...")
        sys.exit(0) 