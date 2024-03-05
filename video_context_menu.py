# Description: A Python script to add a context menu item to Explorer, that will download videos from most Video Sites using yt-dlp.

import os
import sys
import time
import subprocess

def install_package(package_name):
    print(f"{package_name} not found, installing...")
    time.sleep(1)
    # Download Package That Isn't Installed.
    subprocess.check_call([sys.executable, "-m", "pip", "install", package_name])

try:
    import yt_dlp
except ImportError:
    install_package("yt-dlp")
    import yt_dlp

try:
    import winreg as reg
except ImportError:
    install_package("winreg")
    import winreg as reg

try:
    from tkinter import Tk, simpledialog
except ImportError:
    install_package("tkinter")
    from tkinter import Tk, simpledialog

# Function to add the command to the context menu
def add_to_context_menu():
    application_path = os.path.abspath(sys.argv[0])
    
    # Notice the "%V" passed as an argument in the command, which Windows replaces with the clicked directory path
    key_path = r'Directory\\Background\\shell\\DownloadVideo'
    command_path = r'Directory\\Background\\shell\\DownloadVideo\\command'
    
    key = reg.CreateKey(reg.HKEY_CLASSES_ROOT, key_path)
    reg.SetValue(key, '', reg.REG_SZ, '&Download Video')
    reg.CloseKey(key)
    
    command_key = reg.CreateKey(reg.HKEY_CLASSES_ROOT, command_path)

    reg.SetValue(command_key, '', reg.REG_SZ, f'{sys.executable} "{application_path}" "%V"')
    reg.CloseKey(command_key)
    print("Context menu item added.")

# Function to display the URL input dialog and download the video
def download_video_dialog(directory):
    root = Tk()
    root.withdraw() # Hide the Tkinter root window
    video_url = simpledialog.askstring("Download Video", "Enter the Video URL:")
    
    if video_url:
        # Change to the directory where the context menu was clicked
        os.chdir(directory)
        # Assuming yt-dlp is installed and added to PATH
        subprocess.run(["yt-dlp", video_url])

        print('You can safely close this window now.')
        input("Press Enter to exit...")

# Main execution logic
if __name__ == "__main__":
    # Check if "install" is in the command line arguments
    if len(sys.argv) > 1 and sys.argv[1] == "--install":
        add_to_context_menu()
    else:
        # The directory path should be the second argument, if present
        directory = sys.argv[1] if len(sys.argv) > 1 else os.getcwd()
        download_video_dialog(directory)
