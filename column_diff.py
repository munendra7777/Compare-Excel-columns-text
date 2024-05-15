import pandas as pd
import tkinter as tk
from tkinter import filedialog, ttk
from tkinter import PhotoImage, Tk
import threading
import os
import webbrowser

file_paths = []
# Function to open developer LinkedIn profile
def open_linkedin():
    webbrowser.open('https://www.linkedin.com/in/munendra7777/')

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

def col_letter_to_index(col_letter):
    col_index = 0
    for i, letter in enumerate(reversed(col_letter)):
        col_index += (ord(letter.upper()) - 64) * (26 ** i)
    return col_index - 1  # Subtract 1 because Python uses zero-based indexing

def select_file():
    global file_paths
    if multi_file_var.get():
        file_paths = filedialog.askopenfilenames()
        if len(file_paths) > 15:
            file_paths = file_paths[:15]
        file_label.config(text=f"Selected Files: {len(file_paths)} files")
    else:
        file_paths = [filedialog.askopenfilename()]
        file_name = os.path.basename(file_paths[0])
        file_label.config(text=f"Selected File: {file_name}")
    
    if file_paths and file_paths[0] != '':
        process_button.config(state='normal')

def process_file():
    if not file_paths:  # No file selected, so return without doing anything
        return

    col1 = col_letter_to_index(col1_entry.get())
    col2 = col_letter_to_index(col2_entry.get())
    col3 = col_letter_to_index(col3_entry.get())

    for file_path in file_paths:
        try:
            df = pd.read_excel(file_path, header=None, engine='openpyxl')
        except Exception as e:
            print(f"Failed to read Excel file: {e}")
            continue

        progress['maximum'] = len(df)
        for i, row in df.iterrows():
            df.at[i, col3] = compare_text(str(row[col1]), str(row[col2]))
            progress['value'] = i
            root.update_idletasks()

        df.to_excel(file_path, index=False, header=False)

    status_label.config(text='Processing completed')
    process_button.config(state='disabled')

root = Tk()
img = PhotoImage(file='column_diff_large.png')
root.iconphoto(True, img)
root.title("Compare Excel Columns")
root.geometry('800x200')
root.resizable(True, True)

tk.Label(root, text="Enter Column 1 to compare (Eg, AC)").grid(row=0)
tk.Label(root, text="Enter Column 2 to compared with (Eg, D)").grid(row=1)
tk.Label(root, text="Enter the Output Column for the Compared Text (Eg, AD)").grid(row=2)

col1_entry = tk.Entry(root)
col2_entry = tk.Entry(root)
col3_entry = tk.Entry(root)

col1_entry.grid(row=0, column=1)
col2_entry.grid(row=1, column=1)
col3_entry.grid(row=2, column=1)

file_label = tk.Label(root, text="")
file_label.grid(row=3, column=0, columnspan=2)

progress = ttk.Progressbar(root, length=200, mode='determinate')
progress.grid(row=6, column=1)

status_label = tk.Label(root, text="")
status_label.grid(row=5, column=0, columnspan=1)

select_button = tk.Button(root, text='Select File', command=select_file)
select_button.grid(row=6, column=0)
multi_file_var = tk.BooleanVar()

tk.Checkbutton(root, text="Select Multiple Files (Provided input & output columns are same and Files are in the same directory)", variable=multi_file_var).grid(row=7, column=0, columnspan=2)

process_button = tk.Button(root, text='Process File', command=lambda: threading.Thread(target=process_file).start(), state='disabled')
process_button.grid(row=6, column=2)
# developer info
info_button = ttk.Button(root, text='Contact Developer', command=open_linkedin)
info_button.grid(row=9, column=0, sticky='w')

root.mainloop()
