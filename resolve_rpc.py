print("Script starting...")
import time
import psutil
import os
import sys
import winreg
import win32gui
from pypresence import Presence
import traceback
import logging
from datetime import datetime

# Set up logging
log_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'resolve_rpc.log')
logging.basicConfig(
    filename=log_file,
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def log_and_print(message):
    print(message)
    logging.info(message)

log_and_print("Script starting...")

CLIENT_ID = "1378792739767980042"

def check_dependencies():
    log_and_print("Checking dependencies...")
    try:
        import win32gui
        import psutil
        from pypresence import Presence
        log_and_print("All dependencies are installed correctly")
        return True
    except ImportError as e:
        log_and_print(f"Error: Missing required dependency - {e}")
        return False

def get_resolve_version():
    try:
        # Try to get DaVinci Resolve version from registry
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Blackmagic Design\DaVinci Resolve")
        version = winreg.QueryValueEx(key, "Version")[0]
        winreg.CloseKey(key)
        log_and_print(f"Found Resolve version: {version}")
        return version
    except Exception as e:
        log_and_print(f"Error getting Resolve version: {e}")
        return "Unknown Version"

def is_resolve_running():
    log_and_print("Checking for Resolve processes...")
    resolve_processes = []
    for proc in psutil.process_iter(['name', 'pid']):
        try:
            process_name = proc.info['name'].lower()
            if process_name in [
                'resolve.exe', 
                'resolve studio.exe', 
                'davinci resolve.exe', 
                'davinci resolve studio.exe',
                'resolveapp.exe',
                'resolveapp studio.exe',
                'davinciresolve.exe',
                'davinciresolve studio.exe'
            ]:
                log_and_print(f"Found Resolve process: {process_name} (PID: {proc.info['pid']})")
                resolve_processes.append(proc)
        except (psutil.NoSuchProcess, psutil.AccessDenied) as e:
            continue
    
    if not resolve_processes:
        log_and_print("No Resolve processes found")
        return False
    
    log_and_print(f"Found {len(resolve_processes)} Resolve process(es)")
    return True

def is_discord_running():
    log_and_print("Checking for Discord process...")
    for proc in psutil.process_iter(['name']):
        try:
            if proc.info['name'].lower() == 'discord.exe':
                log_and_print("Found Discord process")
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    log_and_print("Discord not found")
    return False

def get_resolve_window_title():
    try:
        def callback(hwnd, windows):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if "DaVinci Resolve" in title:
                    log_and_print(f"Found Resolve window: {title}")
                    windows.append(title)
            return True
        
        windows = []
        win32gui.EnumWindows(callback, windows)
        
        if not windows:
            log_and_print("No Resolve windows found")
            return "DaVinci Resolve"
            
        title = windows[0]
        log_and_print(f"Using window title: {title}")
        return title
    except Exception as e:
        log_and_print(f"Error getting window title: {e}")
        return "DaVinci Resolve"

def wait_for_discord():
    log_and_print("Waiting for Discord to start...")
    while not is_discord_running():
        time.sleep(5)
    log_and_print("Discord found!")

def main():
    if not check_dependencies():
        sys.exit(1)
        
    try:
        log_and_print("Starting Discord RPC for DaVinci Resolve...")
        rpc = None
        
        while True:
            try:
                # Check if Discord is running
                if not is_discord_running():
                    log_and_print("Discord not running, waiting...")
                    if rpc is not None:
                        try:
                            rpc.close()
                            rpc = None
                        except:
                            pass
                    time.sleep(5)
                    continue
                
                # Connect to Discord if not already connected
                if rpc is None:
                    try:
                        log_and_print("Connecting to Discord...")
                        rpc = Presence(CLIENT_ID)
                        rpc.connect()
                        log_and_print("Successfully connected to Discord!")
                    except Exception as e:
                        log_and_print(f"Failed to connect to Discord: {e}")
                        time.sleep(5)
                        continue
                
                if is_resolve_running():
                    window_title = get_resolve_window_title()
                    version = get_resolve_version()
                    
                    # Determine if it's Studio or Free version
                    is_studio = "Studio" in window_title or "Studio" in version
                    log_and_print(f"Is Studio version: {is_studio}")
                    
                    # Determine if in Project Manager or specific project
                    if "Project Manager" in window_title:
                        log_and_print("Updating status: In Project Manager")
                        try:
                            rpc.update(
                                state="In Project Manager",
                                details=f"DaVinci Resolve {'Studio' if is_studio else 'Free'}",
                                large_image="resolve_logo",
                                large_text=f"DaVinci Resolve {'Studio' if is_studio else 'Free'} {version}",
                                buttons=[{"label": "Project Manager", "url": "https://www.blackmagicdesign.com/products/davinciresolve/"}]
                            )
                            log_and_print("Status updated successfully")
                        except Exception as e:
                            log_and_print(f"Failed to update status: {e}")
                            rpc = None  # Reset connection on error
                    else:
                        # Extract project name from window title and clean it up
                        project_name = window_title.replace("DaVinci Resolve", "").replace(" - Studio", "").replace(" - Free", "").strip()
                        if project_name and project_name != "":
                            log_and_print(f"Updating status: Editing {project_name}")
                            try:
                                rpc.update(
                                    state=f"Editing: {project_name}",
                                    details=f"DaVinci Resolve {'Studio' if is_studio else 'Free'}",
                                    large_image="resolve_logo",
                                    large_text=f"DaVinci Resolve {'Studio' if is_studio else 'Free'} {version}",
                                    buttons=[{"label": "View Project", "url": "https://www.blackmagicdesign.com/products/davinciresolve/"}]
                                )
                                log_and_print("Status updated successfully")
                            except Exception as e:
                                log_and_print(f"Failed to update status: {e}")
                                rpc = None  # Reset connection on error
                        else:
                            log_and_print("Updating status: Editing (Project Name Unknown)")
                            try:
                                rpc.update(
                                    state="Editing",
                                    details=f"DaVinci Resolve {'Studio' if is_studio else 'Free'}",
                                    large_image="resolve_logo",
                                    large_text=f"DaVinci Resolve {'Studio' if is_studio else 'Free'} {version}",
                                    buttons=[{"label": "View Project", "url": "https://www.blackmagicdesign.com/products/davinciresolve/"}]
                                )
                                log_and_print("Status updated successfully")
                            except Exception as e:
                                log_and_print(f"Failed to update status: {e}")
                                rpc = None  # Reset connection on error
                else:
                    log_and_print("Resolve not running, clearing status")
                    try:
                        if rpc is not None:
                            rpc.clear()
                            log_and_print("Status cleared successfully")
                    except Exception as e:
                        log_and_print(f"Failed to clear status: {e}")
                        rpc = None  # Reset connection on error
                
                time.sleep(15)  # Update every 15 seconds
            except Exception as e:
                log_and_print(f"Error in main loop: {e}")
                rpc = None  # Reset connection on error
                time.sleep(15)  # Wait before retrying
    except Exception as e:
        log_and_print(f"Fatal error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        log_and_print("\nExiting...")
        sys.exit(0)




