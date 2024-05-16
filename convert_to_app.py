import subprocess
import platform
import os
import sys

# Get the base path of the script
if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS  # Path to the bundled resources
else:
    base_path = os.path.dirname(os.path.abspath(__file__))

# Construct the full path to the image
image_path = os.path.join(base_path, "column_diff_large.png")

# Determine the correct path separator based on the OS
path_separator = ';' if platform.system() == 'Windows' else ':'

# Define the command as a string
command = (
    f'pyinstaller --onefile --windowed '
    f'--icon="{image_path}" '
    f'--add-data="{image_path}{path_separator}." '
    f'"column_diff.py"'
)

# Run the command
subprocess.run(command, shell=True)
