import sys
import os
from sotool import main

# Add the 'src' directory to Python's path
# This allows Python (and PyInstaller) to find the 'sotool' package
project_root = os.path.dirname(os.path.abspath(__file__))
src_path = os.path.join(project_root, "src")
sys.path.insert(0, src_path)

if __name__ == "__main__":
    main()

# pyinstaller --onefile --windowed --name sotool_app main.py
