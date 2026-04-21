import os
import subprocess
import sys
from PIL import Image

def generate_icon(png_path, ico_path):
    print(f"Converting {png_path} to {ico_path}...")
    try:
        img = Image.open(png_path)
        img.save(ico_path, format='ICO', sizes=[(256, 256)])
        print("Icon generated successfully.")
    except Exception as e:
        print(f"Failed to generate icon: {e}")
        print("Note: ensure 'Pillow' is installed (pip install Pillow).")

def build():
    # Source image from artifact directory
    artifact_png = r"C:\Users\ringp\.gemini\antigravity\brain\c5af4d75-0809-430b-96c6-4d4be9fdb700\smart_ring_icon_1776313610943.png"
    ico_path = os.path.join(os.path.dirname(__file__), "app.ico")
    
    if os.path.exists(artifact_png):
        generate_icon(artifact_png, ico_path)
    else:
        print("Icon PNG not found, skipping icon generation.")
        
    print("Running PyInstaller...")
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm",
        "--onefile",
        "--windowed",
        f"--icon={ico_path}",
        "--name=SmartRingQC",
        "--add-data", f"app;app", # Include the app package essentially, although onefile might just need it via imports
        "app/main.py"
    ]
    
    # In Windows, paths might need slightly different handling for add-data if relying on MEIPASS.
    # PyInstaller usually automatically picks up python files. We might just let it auto-bundle.
    cmd_clean = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm",
        "--onefile",
        "--windowed",
        f"--icon={ico_path}",
        "--name=SmartRingQC",
        "app/main.py"
    ]
    
    print("Executing command:", " ".join(cmd_clean))
    subprocess.run(cmd_clean)
    print("Build complete. Executable should be in the 'dist' folder.")

if __name__ == "__main__":
    build()
