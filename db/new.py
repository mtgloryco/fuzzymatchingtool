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
    
    if file_label == "File 1":
        drop_down = tk.OptionMenu(control_frame, selected_col, *options)
        drop_down_label = tk.Label(control_frame, text=f"Columns for File 1")
        drop_down_label.grid(row=4, column=0, sticky="w", pady=5, padx=5)
        drop_down.grid(row=4, column=1, pady=5, padx=5, sticky="ew")
    else:
        # For File 2, we're using a different variable
        global selected_col2
        selected_col2 = StringVar()
        selected_col2.set(options[0])
        drop_down = tk.OptionMenu(control_frame, selected_col2, *options)
        drop_down_label = tk.Label(control_frame, text=f"Columns for File 2")
        drop_down_label.grid(row=5, column=0, sticky="w", pady=5, padx=5)
        drop_down.grid(row=5, column=1, pady=5, padx=5, sticky="ew")

def read_columns(file, file_label, tree):
    sheet = pd.read_excel(file, "Sheet1")
    cols = sheet.columns.tolist()
    tree.config(columns=cols)
    drop_down_menu(cols, file_label)
    for item in tree.get_children():
        tree.delete(item)
    for col in cols:
        tree.heading(col, text=col)
        tree.column(col, width=100, minwidth=100)
    for _, row in sheet.iterrows():
        tree.insert("", "end", values=list(row))

def select_file1():
    global file1_path
    file1_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if file1_path:
        filename = file1_path.split('/')[-1]
        file1_label.config(text=f"Selected: {filename}")
        read_columns(file1_path, "File 1", tree1)
    
def select_file2():
    global file2_path
    file2_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if file2_path:
        filename = file2_path.split('/')[-1]
        file2_label.config(text=f"Selected: {filename}")
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
        
        # Use the appropriate selected column variable for each file
        col1 = selected_col.get()
        col2 = selected_col2.get() if 'selected_col2' in globals() else selected_col.get()
        
        if col1 not in options:
            raise ValueError("Selected column for File 1 is not valid.")
            
        index_of_item = options.index(col1)
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
        messagebox.showinfo("Success", f"Matching complete! Results saved as '{custom_filename}_{current_time}.xlsx'.")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")


root = tk.Tk()
root.title("Excel Fuzzy Matching")
root.geometry("1200x800")  # Set initial window size

# Create main horizontal frame to hold the two sections
main_horizontal_frame = ttk.Frame(root)
main_horizontal_frame.pack(fill="both", expand=True)

# Create files frame on the left
files_frame = ttk.LabelFrame(main_horizontal_frame, text="Excel Files")
files_frame.pack(side="left", fill="both", expand=True, padx=5, pady=5)

# Create controls frame on the right with fixed width
control_frame = ttk.LabelFrame(main_horizontal_frame, text="Controls", width=250)
control_frame.pack(side="right", fill="y", padx=5, pady=5)
control_frame.pack_propagate(False)  # Prevent the frame from shrinking

# Set up the files frame with scrolling
files_canvas = tk.Canvas(files_frame)
files_scrollbar_y = ttk.Scrollbar(files_frame, orient="vertical", command=files_canvas.yview)
files_scrollbar_x = ttk.Scrollbar(files_frame, orient="horizontal", command=files_canvas.xview)
files_scroll_frame = ttk.Frame(files_canvas)

files_scroll_frame.bind(
    "<Configure>",
    lambda e: files_canvas.configure(
        scrollregion=files_canvas.bbox("all")
    )
)

# Create a window inside the canvas
files_canvas.create_window((0, 0), window=files_scroll_frame, anchor="nw")
files_canvas.configure(
    yscrollcommand=files_scrollbar_y.set,
    xscrollcommand=files_scrollbar_x.set
)

# Pack the scrollbars and canvas
files_scrollbar_y.pack(side="right", fill="y")
files_scrollbar_x.pack(side="bottom", fill="x")
files_canvas.pack(side="left", fill="both", expand=True)

# Add mouse wheel support for vertical scrolling
def _on_mousewheel(event):
    files_canvas.yview_scroll(int(-1*(event.delta/120)), "units")

# Add shift + mouse wheel for horizontal scrolling
def _on_shift_mousewheel(event):
    files_canvas.xview_scroll(int(-1*(event.delta/120)), "units")

files_canvas.bind_all("<MouseWheel>", _on_mousewheel)
files_canvas.bind_all("<Shift-MouseWheel>", _on_shift_mousewheel)

# Add these bindings for better scrolling experience
root.bind('<Up>', lambda event: files_canvas.yview_scroll(-1, "units"))
root.bind('<Down>', lambda event: files_canvas.yview_scroll(1, "units"))
root.bind('<Left>', lambda event: files_canvas.xview_scroll(-1, "units"))
root.bind('<Right>', lambda event: files_canvas.xview_scroll(1, "units"))
root.bind('<Prior>', lambda event: files_canvas.yview_scroll(-1, "pages"))
root.bind('<Next>', lambda event: files_canvas.yview_scroll(1, "pages"))

style = ttk.Style()
style.theme_use("default")

# Change Treeview heading text color (foreground)
style.configure("Treeview.Heading", foreground="red", font=('Helvetica', 10, 'bold'))

# First tree with scrollbars
frame1 = ttk.LabelFrame(files_scroll_frame, text="First Excel File")
frame1.pack(padx=5, pady=5, fill="both", expand=True)

# Create scrollbar frame1
tree1_frame = ttk.Frame(frame1)
tree1_frame.pack(fill="both", expand=True)

# Create vertical scrollbar
tree1_scroll_y = ttk.Scrollbar(tree1_frame)
tree1_scroll_y.pack(side="right", fill="y")

# Create horizontal scrollbar
tree1_scroll_x = ttk.Scrollbar(tree1_frame, orient="horizontal")
tree1_scroll_x.pack(side="bottom", fill="x")

tree1 = ttk.Treeview(tree1_frame, yscrollcommand=tree1_scroll_y.set, xscrollcommand=tree1_scroll_x.set)
tree1.pack(fill="both", expand=True)
tree1["show"] = "headings"

tree1_scroll_y.config(command=tree1.yview)
tree1_scroll_x.config(command=tree1.xview)

# Second tree with scrollbars
frame2 = ttk.LabelFrame(files_scroll_frame, text="Second Excel File")
frame2.pack(padx=5, pady=5, fill="both", expand=True)

# create scrollbar frame2
tree2_frame = ttk.Frame(frame2)
tree2_frame.pack(fill="both", expand=True)

# Create vertical scrollbar
tree2_scroll_y = ttk.Scrollbar(tree2_frame)
tree2_scroll_y.pack(side="right", fill="y")

tree2_scroll_x = ttk.Scrollbar(tree2_frame, orient="horizontal")
tree2_scroll_x.pack(side="bottom", fill="x")

tree2 = ttk.Treeview(tree2_frame, yscrollcommand=tree2_scroll_y.set, xscrollcommand=tree2_scroll_x.set)
tree2.pack(fill="both", expand=True)
tree2["show"] = "headings"

tree2_scroll_y.config(command=tree2.yview)
tree2_scroll_x.config(command=tree2.xview)

# Set up the control frame to match the image layout
# Configure grid columns to expand evenly
control_frame.grid_columnconfigure(0, weight=1)
control_frame.grid_columnconfigure(1, weight=1)

# File 1 selection label - full width
file1_label = tk.Label(control_frame, text="Selected: ", bg="light gray", relief="sunken", height=2)
file1_label.grid(row=0, column=0, columnspan=2, pady=5, padx=5, sticky="ew")

# Choose File 1 button - full width
file1_btn = tk.Button(control_frame, text="Choose File 1", command=select_file1, height=2)
file1_btn.grid(row=1, column=0, columnspan=2, pady=5, padx=5, sticky="ew")

# File 2 selection label - left side only
file2_label = tk.Label(control_frame, text="Selected: ", bg="light gray", relief="sunken", height=2)
file2_label.grid(row=2, column=0, pady=5, padx=5, sticky="ew")

# Choose File 2 button - right side only
file2_btn = tk.Button(control_frame, text="Choose File 2", command=select_file2, height=2)
file2_btn.grid(row=2, column=1, pady=5, padx=5, sticky="ew")

# Add dropdown placeholders (these will be replaced when files are loaded)
label_col_file1 = tk.Label(control_frame, text="Columns for File 1")
label_col_file1.grid(row=4, column=0, sticky="w", pady=5, padx=5)
dropdown_placeholder1 = tk.Label(control_frame, text="Email", relief="raised", width=15)
dropdown_placeholder1.grid(row=4, column=1, pady=5, padx=5, sticky="ew")

label_col_file2 = tk.Label(control_frame, text="Columns for File 2")
label_col_file2.grid(row=5, column=0, sticky="w", pady=5, padx=5)
dropdown_placeholder2 = tk.Label(control_frame, text="Email", relief="raised", width=15)
dropdown_placeholder2.grid(row=5, column=1, pady=5, padx=5, sticky="ew")

# Start matching button - full width with green background
match_btn = tk.Button(control_frame, text="Start Matching", command=start_matching, bg="green", fg="white", height=2)
match_btn.grid(row=6, column=0, columnspan=2, pady=10, padx=5, sticky="ew")

root.mainloop()