import subprocess

# Define the command as a string
command = (
    'pyinstaller --onefile --windowed '
    '--icon="column_diff_large.png" '
    '--add-data="column_diff_large.png;." '
    '"column_diff.py"'
)

# Run the command
subprocess.run(command, shell=True)
