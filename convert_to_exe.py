import subprocess
import platform

# Determine the correct path separator based on the OS
path_separator = ';' if platform.system() == 'Windows' else ':'

# Define the command as a string
command = (
    f'pyinstaller --onefile --windowed '
    f'--icon="column_diff_large.png" '
    f'--add-data="column_diff_large.png{path_separator}." '
    f'"column_diff.py"'
)

# Run the command
subprocess.run(command, shell=True)
