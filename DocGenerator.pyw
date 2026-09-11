"""
DocGini - Batch Word Document Generator
Runs without a console window when launched via pythonw.
"""
import sys
import os

# Ensure script directory is on sys.path
script_dir = os.path.dirname(os.path.abspath(__file__))
if script_dir not in sys.path:
    sys.path.insert(0, script_dir)

from DocGenerator import main

if __name__ == "__main__":
    main()
