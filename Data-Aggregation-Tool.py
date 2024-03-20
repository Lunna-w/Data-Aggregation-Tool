import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

def import_and_process_files(folder_path):
    for filename in os.listdir(folder_path):
        if filename.endswith('.txt'):
            file_path = os.path.join(folder_path, filename)
            try:
                with open(file_path, encoding='Windows-1252') as file:
                    delimiter = '¬' if '¬' in file.read() else 'ï¿½'
                    df = pd.read_csv(file_path, sep=delimiter, header=None, names=['Banner ID', 'Campus', 'Scanner Location', 'Date'], encoding='Windows-1252')
                    excel_output_path = os.path.splitext(file_path)[0] + '.xlsx'
                    df.to_excel(excel_output_path, index=False)
                    print(f'Excel file "{excel_output_path}" created successfully.')
            except UnicodeDecodeError:
                print(f'Error decoding file "{file_path}"')

def browse_folder():
    folder_path = filedialog.askdirectory()
    if folder_path:
        import_and_process_files(folder_path)
        messagebox.showinfo("Success", "Excel files created successfully.")

root = tk.Tk()
root.title("TXT to Excel Converter")

label = tk.Label(root, text="Select folder containing TXT files:")
label.pack(pady=10)
browse_button = tk.Button(root, text="Browse", command=browse_folder)
browse_button.pack(pady=5)

root.mainloop()
