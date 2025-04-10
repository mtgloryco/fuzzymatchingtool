import openpyxl as op
import pandas as pd
from rapidfuzz import process
from tkinter import filedialog, messagebox, StringVar, simpledialog
import collections
import tkinter as tk
from datetime import datetime
from tkinter import ttk

def drop_down_menu(columns, file_label):
    global options
    options = [col for col in columns]
    global selected_col
    selected_col = StringVar()
    selected_col.set(options[0])

    drop_down = tk.OptionMenu(root, selected_col, *options)
    drop_down_label = tk.Label(root, text=f"Columns for {file_label} ")
    drop_down_label.pack()
    drop_down.pack()


def read_columns(file, file_label, tree):
    sheet = pd.read_excel(file, "Sheet1")
    cols = sheet.columns.tolist()
    tree.config(columns=cols)
    drop_down_menu(cols, file_label)
    for col in cols:
        tree.heading(col, text=col)
        tree.column(col, width=100)
    for _, row in sheet.iterrows():
        tree.insert("", "end", values=list(row))

def select_file1():
    global file1_path
    file1_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    file1_label.config(text=f"Selected: {file1_path.split('/')[-1]}")
    read_columns(file1_path,"File 1", tree1)
def select_file2():
    global file2_path
    file2_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    file2_label.config(text=f"Selected: {file2_path.split('/')[-1]}")
    read_columns(file2_path, "File 2", tree2)
def start_matching():
    try:
        custom_filename = simpledialog.askstring("Input", "Enter the name for the result file:")
        if not custom_filename:
            raise ValueError("Filename cannot be empty")
        
        file1 = op.load_workbook(file1_path)
        sheet1 = file1.active

        file2 = op.load_workbook(file2_path)
        sheet2 = file2.active  

        result = op.Workbook()
        result_sheet = result.active
        result_sheet.title = "Result"
        if selected_col.get() not in options:
            raise ValueError("Selected column is not valid.")
        index_of_item = options.index(selected_col.get())
        row_tracker = 1
        email2 = [r[index_of_item] for r in sheet2.iter_rows(values_only=True) if r[index_of_item]]

       
        result_sheet['A1'].value = "Email1"
        result_sheet['B1'].value = "Matched Email"
        result_sheet['C1'].value = "Match Score"

        res = collections.defaultdict(list)

        for t in sheet1.iter_rows(values_only=True):
            email = t[index_of_item]
            if email:
                res[email] = process.extract(email, email2, limit=1)
                result_sheet[f'A{row_tracker + 1}'].value = email
                result_sheet[f'B{row_tracker + 1}'].value = res[email][0][0]
                result_sheet[f'C{row_tracker + 1}'].value = res[email][0][1]
                row_tracker += 1

        current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        result.save(f"{custom_filename}_{current_time}.xlsx")
        messagebox.showinfo("Success", f"Matching complete! Results saved as '{custom_filename}.xlsx'.")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")


root = tk.Tk()
root.title("Excel Fuzzy Matching")
style = ttk.Style()
style.theme_use("default")

# Change Treeview heading text color (foreground)
style.configure("Treeview.Heading", foreground="red", font=('Helvetica', 10, 'bold'))

frame1 = ttk.LabelFrame(root, text="First Excell File")
frame1.pack(padx=5, pady=5, fill="both", expand=True)

frame2 = ttk.LabelFrame(root, text="Second Excell File")
frame2.pack(padx=5, pady=5, fill="both", expand=True)


#Iyi ni scroll frame

tree1_frame = ttk.Frame(frame1)
tree1_frame.pack(fill="both", expand=True)
tree2_frame = ttk.Frame(frame2)
tree2_frame.pack(fill="both", expand=True)

#scroll for frame one for vertical axis

tree1_scroll_y = ttk.Scrollbar(tree1_frame)
tree1_scroll_y.pack(side="right", expand=True)

tree2_scroll_y = ttk.Scrollbar(tree2_frame)
tree2_scroll_y.pack(side="right", fill="y")

# Create tree1 Treeview widget
tree1 = ttk.Treeview(tree1_frame, yscrollcommand=tree1_scroll_y.set)
tree1.pack(fill="both", expand=True)
tree1_scroll_y.config(command=tree1.yview)

#scroll for horizontal axis
tree1_scroll_x = ttk.Scrollbar(tree1_frame, orient="horizontal", command=tree1.xview)
tree1_scroll_x.pack(side="bottom", fill="x")
tree1.config(xscrollcommand=tree1_scroll_x.set)

tree2_scroll_x = ttk.Scrollbar(tree2_frame, orient="horizontal")
tree1_scroll_x.pack(side="bottom", fill="x")
# scroling for tree 2

tree2 = ttk.Treeview(tree2_frame, yscrollcommand=tree2_scroll_y.set, xscrollcommand=tree2_scroll_x.set)
tree2.pack(fill="both", expand=True)
tree2_scroll_y.config(command=tree2.yview)
tree2_scroll_x.config(command=tree2.xview)

#code to create styling of the created frames
styling = ttk.Style()
styling.theme_use("default")
styling.configure("Treeview.Heading", foreground="red", font=("Verdana", 10, "bold"))


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
