import openpyxl as op
import pandas as pd
from fuzzywuzzy import process
from tkinter import filedialog, messagebox
import collections
import tkinter as tk
from tkinter import ttk

def read_columns(file):
    sheet = pd.read_excel(file, "Sheet1")
    cols = sheet.columns.tolist()
    tree.config(columns=cols)

    for col in cols:
        tree.heading(col, text=col)
        tree.column(col, anchor='center')

def select_file1():
    global file1_path
    file1_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    file1_label.config(text=f"Selected: {file1_path.split('/')[-1]}")
    read_columns(file1_path)
def select_file2():
    global file2_path
    file2_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    file2_label.config(text=f"Selected: {file2_path.split('/')[-1]}")

def start_matching():
    try:
        file1 = op.load_workbook(file1_path)
        sheet1 = file1.active

        file2 = op.load_workbook(file2_path)
        sheet2 = file2.active  

        result = op.Workbook()
        result_sheet = result.active
        result_sheet.title = "Result"

        row_tracker = 1
        email2 = [r[2] for r in sheet2.iter_rows(values_only=True) if r[2]]

       
        result_sheet['A1'].value = "Email1"
        result_sheet['B1'].value = "Matched Email"
        result_sheet['C1'].value = "Match Score"

        res = collections.defaultdict(list)

        for t in sheet1.iter_rows(values_only=True):
            email = t[2]
            if email:
                res[email] = process.extract(email, email2, limit=1)
                result_sheet[f'A{row_tracker + 1}'].value = email
                result_sheet[f'B{row_tracker + 1}'].value = res[email][0][0]
                result_sheet[f'C{row_tracker + 1}'].value = res[email][0][1]
                row_tracker += 1

        
        result.save("Result.xlsx")
        messagebox.showinfo("Success", f"Matching complete! Results saved as '{result_path}'.")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")


root = tk.Tk()
root.title("Excel Fuzzy Matching")
style = ttk.Style()
style.theme_use("default")

# Change Treeview heading text color (foreground)
style.configure("Treeview.Heading", foreground="red", font=('Helvetica', 10, 'bold'))
tree = ttk.Treeview(root)
tree.pack()
file1_label = tk.Label(root, text="Select first Excel file", width=50)
file1_label.pack(pady=5)
file1_btn = tk.Button(root, text="Choose File 1", command=select_file1)
file1_btn.pack(pady=5)

file2_label = tk.Label(root, text="Select second Excel file", width=50)
file2_label.pack(pady=5)
file2_btn = tk.Button(root, text="Choose File 2", command=select_file2)
file2_btn.pack(pady=5)

match_btn = tk.Button(root, text="Start Matching", command=start_matching, bg="green", fg="white")
match_btn.pack(pady=10)


root.mainloop()
