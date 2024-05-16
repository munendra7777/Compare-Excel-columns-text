import subprocess
import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from tkinter import PhotoImage, Tk
import threading
import webbrowser
import time
import sys

file_paths = []

# Function to open developer LinkedIn profile
def open_linkedin():
    webbrowser.open('https://www.linkedin.com/in/munendra7777/')

# Function to compare two texts and return the differences
def compare_text(text1, text2, delimiter=" | "):
    words1 = text1.split(" ")
    words2 = text2.split(" ")
    result1 = []
    result2 = []

    for word in words1[:]:
        if word in words2:
            words1.remove(word)
            words2.remove(word)

    for word in words1:
        if word != "":
            result1.append(word)

    for word in words2:
        if word != "":
            result2.append(word)

    if result1 or result2:
        return " ".join(result1) + delimiter + " ".join(result2)
    else:
        return ""

# Function to convert Excel column letter to index
def col_letter_to_index(col_letter):
    col_index = 0
    for i, letter in enumerate(reversed(col_letter)):
        col_index += (ord(letter.upper()) - 64) * (26 ** i)
    return col_index - 1  # Subtract 1 because Python uses zero-based indexing

# Function to open file dialog and select files
def select_file():
    global file_paths
    if multi_file_var.get():
        # Select multiple files
        file_paths = filedialog.askopenfilenames()
        if len(file_paths) > 15:
            file_paths = file_paths[:15]
        file_label.config(text=f"Selected Files: {len(file_paths)} files")
    else:
        # Select a single file
        file_paths = [filedialog.askopenfilename()]
        file_name = os.path.basename(file_paths[0])
        file_label.config(text=f"Selected File: {file_name}")

    if file_paths and file_paths[0] != '':
        # Check file format
        ext = os.path.splitext(file_paths[0])[-1].lower()
        if ext not in ['.xlsx', '.csv', '.tsv', '.numbers']:
            messagebox.showerror("Error", f"Unsupported file format: {ext}")
            return
        
        process_button.config(state='normal')

# Function to convert Numbers files to CSV using AppleScript
def convert_numbers_to_csv(file_path):
    script_path = 'export_to_csv.applescript'
    args = ['osascript', script_path, file_path]
    result = subprocess.run(args, capture_output=True, text=True)
    print("AppleScript stdout:", result.stdout)
    print("AppleScript stderr:", result.stderr)
    return file_path + '.csv'

# Function to show a popup when processing is completed
def show_completion_popup():
    messagebox.showinfo("Processing Completed", "All files have been processed successfully.")
    # Schedule the popup to be shown after a short delay
    root.after(100, show_popup)


# Function to process the selected file(s)
def process_file():
    if not file_paths:  # No file selected, so return without doing anything
        return

    col1 = col_letter_to_index(col1_entry.get())
    col2 = col_letter_to_index(col2_entry.get())
    col3 = col_letter_to_index(col3_entry.get())

    for file_path in file_paths:
        ext = os.path.splitext(file_path)[-1].lower()
        if ext == '.numbers':
            csv_file_path = convert_numbers_to_csv(file_path)
            if not os.path.exists(csv_file_path):
                time.sleep(2)  # Wait for the file to be generated
            file_path = csv_file_path
            ext = '.csv'

        try:
            # Read file based on its extension
            if ext == '.xlsx':
                df = pd.read_excel(file_path, header=None, engine='openpyxl')
            elif ext == '.csv':
                df = pd.read_csv(file_path, header=None)
            elif ext == '.tsv':
                df = pd.read_csv(file_path, header=None, delimiter='\t')
            else:
                messagebox.showerror("Error", f"Unsupported file format: {ext}")
                print(f"Unsupported file format: {ext}")
                continue
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read {ext} file: {e}")
            print(f"Failed to read {ext} file: {e}")
            continue

        # Set the progress bar maximum value
        progress['maximum'] = len(df)
        for i, row in df.iterrows():
            # Compare the text in the specified columns
            df.at[i, col3] = compare_text(str(row[col1]), str(row[col2]))
            progress['value'] = i + 1
            root.update_idletasks()

        try:
            # Write the updated DataFrame back to file
            if ext == '.xlsx':
                df.to_excel(file_path, index=False, header=False)
            elif ext == '.csv':
                df.to_csv(file_path, index=False, header=False)
            elif ext == '.tsv':
                df.to_csv(file_path, index=False, header=False, sep='\t')
        except Exception as e:
            messagebox.showerror("Error", f"Failed to write {ext} file: {e}")
            print(f"Failed to write {ext} file: {e}")
            continue

    status_label.config(text='Processing completed')
    process_button.config(state='disabled')
    # Show completion popup
    root.after(100, show_completion_popup)

# Get the base path of the script
if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS  # Path to the bundled resources
else:
    base_path = os.path.dirname(os.path.abspath(__file__))

# Construct the full path to the image
image_path = os.path.join(base_path, "column_diff_large.png")
print(image_path)

# Initialize the Tkinter root window
root = Tk()
img = PhotoImage(file=image_path)
root.iconphoto(True, img)
root.title("Compare Excel Columns")
root.geometry('800x210')
root.resizable(True, True)

# Add labels and entry fields for column inputs
tk.Label(root, text="Enter Column 1 to compare (Eg, A)").grid(row=0)
tk.Label(root, text="Enter Column 2 to compared with (Eg, B)").grid(row=1)
tk.Label(root, text="Enter the Output Column for the Compared Text (Eg, C)").grid(row=2)

col1_entry = tk.Entry(root)
col2_entry = tk.Entry(root)
col3_entry = tk.Entry(root)

col1_entry.grid(row=0, column=1)
col2_entry.grid(row=1, column=1)
col3_entry.grid(row=2, column=1)

# Label to display selected files
file_label = tk.Label(root, text="")
file_label.grid(row=3, column=0, columnspan=2)

# Progress bar for file processing
progress = ttk.Progressbar(root, length=200, mode='determinate')
progress.grid(row=6, column=1)

# Status label
status_label = tk.Label(root, text="")
status_label.grid(row=5, column=0, columnspan=1)

# Button to select files
select_button = tk.Button(root, text='Select File', command=select_file)
select_button.grid(row=6, column=0)

# Checkbox to select multiple files
multi_file_var = tk.BooleanVar()
tk.Checkbutton(root, text="Select Multiple Files (Provided input & output columns are same and Files are in the same directory)", variable=multi_file_var).grid(row=7, column=0, columnspan=2)

# Button to start processing files
process_button = tk.Button(root, text='Process File', command=lambda: threading.Thread(target=process_file).start(), state='disabled')
process_button.grid(row=6, column=2)

# Button to open developer LinkedIn profile
info_button = ttk.Button(root, text='Contact Developer', command=open_linkedin)
info_button.grid(row=9, column=0, sticky='w')

# Start the Tkinter event loop
root.mainloop()
