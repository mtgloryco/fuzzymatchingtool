import openpyxl as op
import pandas as pd
from rapidfuzz import process
from tkinter import filedialog, messagebox, StringVar, simpledialog
import collections
import tkinter as tk
from datetime import datetime
from tkinter import ttk

def get_sheet_names(file_path):
    """Get all sheet names from an Excel file."""
    workbook = op.load_workbook(file_path, read_only=True)
    return workbook.sheetnames

def update_sheet_dropdown(file_path, is_file1=True):
    """Update the sheet dropdown with sheet names from the selected file."""
    try:
        sheet_names = get_sheet_names(file_path)
        
        if is_file1:
            # Clear existing dropdown
            sheet1_dropdown['menu'].delete(0, 'end')
            # Reset the selected sheet variable
            selected_sheet1.set(sheet_names[0])
            # Add new options
            for sheet in sheet_names:
                sheet1_dropdown['menu'].add_command(label=sheet, 
                                                   command=lambda s=sheet: selected_sheet1.set(s))
        else:
            # Clear existing dropdown
            sheet2_dropdown['menu'].delete(0, 'end')
            # Reset the selected sheet variable
            selected_sheet2.set(sheet_names[0])
            # Add new options
            for sheet in sheet_names:
                sheet2_dropdown['menu'].add_command(label=sheet, 
                                                   command=lambda s=sheet: selected_sheet2.set(s))
    except Exception as e:
        messagebox.showerror("Error", f"Could not read sheets: {e}")

def drop_down_menu(columns, file_label):
    global options, dropdown1, dropdown2, selected_col1, selected_col2
    options = [col for col in columns]
    
    if file_label == "File 1":
        # Initialize selected_col1 as a global variable
        global selected_col1
        if not 'selected_col1' in globals():
            selected_col1 = StringVar(root)
        selected_col1.set(options[0] if options else "")
        
        # Store selected_col1 as an attribute of root
        root.selected_col1 = selected_col1
        
        # Clear the existing dropdown
        if 'dropdown1' in globals() and dropdown1 is not None:
            dropdown1.destroy()
            
        # Create new dropdown
        dropdown1 = tk.OptionMenu(control_frame, selected_col1, *options)
        dropdown1.grid(row=4, column=1, pady=5, padx=5, sticky="ew")
        
    else:
        # Initialize selected_col2 as a global variable
        global selected_col2
        if not 'selected_col2' in globals():
            selected_col2 = StringVar(root)
        selected_col2.set(options[0] if options else "")
        
        # Store selected_col2 as an attribute of root
        root.selected_col2 = selected_col2
        
        # Clear the existing dropdown
        if 'dropdown2' in globals() and dropdown2 is not None:
            dropdown2.destroy()
            
        # Create new dropdown
        dropdown2 = tk.OptionMenu(control_frame, selected_col2, *options)
        dropdown2.grid(row=5, column=1, pady=5, padx=5, sticky="ew")

def load_sheet_data(file_path, sheet_name, tree, file_label):
    """Load data from specified sheet into the treeview."""
    try:
        sheet = pd.read_excel(file_path, sheet_name)
        cols = sheet.columns.tolist()
        tree.config(columns=cols)
        drop_down_menu(cols, file_label)
        
        # Clear the tree
        for item in tree.get_children():
            tree.delete(item)
            
        # Set headers
        for col in cols:
            tree.heading(col, text=col)
            tree.column(col, width=100, minwidth=100)
            
        # Add data
        for _, row in sheet.iterrows():
            tree.insert("", "end", values=list(row))
            
    except Exception as e:
        messagebox.showerror("Error", f"Could not load sheet data: {e}")

def on_sheet1_selected(*args):
    """Called when a sheet is selected for file 1."""
    if hasattr(root, 'file1_path') and root.file1_path:
        # Load and display the selected sheet
        load_sheet_data(root.file1_path, selected_sheet1.get(), tree1, "File 1")
        
        try:
            # Open the workbook for reading and writing
            workbook = op.load_workbook(root.file1_path)
            # Set the active sheet
            workbook.active = workbook[selected_sheet1.get()]
            # Save the workbook with the new active sheet
            workbook.save(root.file1_path)
        except Exception as e:
            # Don't show error to user as this is an enhancement, not critical functionality
            print(f"Could not set active sheet in file: {e}")

def on_sheet2_selected(*args):
    """Called when a sheet is selected for file 2."""
    if hasattr(root, 'file2_path') and root.file2_path:
        # Load and display the selected sheet
        load_sheet_data(root.file2_path, selected_sheet2.get(), tree2, "File 2")
        
        # Open the workbook and set the active sheet to the selected one
        try:
            # Open the workbook for reading and writing
            workbook = op.load_workbook(root.file2_path)
            # Set the active sheet
            workbook.active = workbook[selected_sheet2.get()]
            # Save the workbook with the new active sheet
            workbook.save(root.file2_path)
        except Exception as e:
            # Don't show error to user as this is an enhancement, not critical functionality
            print(f"Could not set active sheet in file: {e}")

def select_file1():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if file_path:
        root.file1_path = file_path
        filename = file_path.split('/')[-1]
        file1_label.config(text=f"Selected: {filename}")
        
        # Update sheet dropdown
        update_sheet_dropdown(file_path, is_file1=True)
        
        # Load the first sheet data
        if hasattr(root, 'selected_sheet1'):
            load_sheet_data(file_path, selected_sheet1.get(), tree1, "File 1")
    
def select_file2():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if file_path:
        root.file2_path = file_path
        filename = file_path.split('/')[-1]
        file2_label.config(text=f"Selected: {filename}")
        
        # Update sheet dropdown
        update_sheet_dropdown(file_path, is_file1=False)
        
        # Load the first sheet data
        if hasattr(root, 'selected_sheet2'):
            load_sheet_data(file_path, selected_sheet2.get(), tree2, "File 2")
    
def start_matching():
    try:
        # Check if files are selected
        if not hasattr(root, 'file1_path') or not hasattr(root, 'file2_path'):
            raise ValueError("Please select both Excel files first")
        
        # Check if columns are selected
        if not hasattr(root, 'selected_col1') or not hasattr(root, 'selected_col2'):
            raise ValueError("Please select columns for matching from both files")
            
        # Check if selected columns have values
        if not selected_col1.get() or not selected_col2.get():
            raise ValueError("Please select columns for matching from both files")
            
        # Get filename from user
        custom_filename = simpledialog.askstring("Input", "Enter the name for the result file:")
        if not custom_filename:
            raise ValueError("Filename cannot be empty")
            
        # Load workbooks with selected sheets
        file1 = op.load_workbook(root.file1_path)
        sheet1 = file1[selected_sheet1.get()]

        file2 = op.load_workbook(root.file2_path)
        sheet2 = file2[selected_sheet2.get()]

        result = op.Workbook()
        result_sheet = result.active
        result_sheet.title = "Result"
        
        # Use the appropriate selected column variables for each file
        if not hasattr(root, 'selected_col1') or not hasattr(root, 'selected_col2'):
            raise ValueError("Please select columns for matching")
            
        col1 = selected_col1.get()
        col2 = selected_col2.get()
        
        # Get column indices
        df1 = pd.read_excel(root.file1_path, selected_sheet1.get())
        df2 = pd.read_excel(root.file2_path, selected_sheet2.get())
        
        if col1 not in df1.columns:
            raise ValueError(f"Column '{col1}' not found in File 1")
        if col2 not in df2.columns:
            raise ValueError(f"Column '{col2}' not found in File 2")
            
        col1_index = df1.columns.get_loc(col1)
        col2_index = df2.columns.get_loc(col2)
        
        row_tracker = 1
        values_to_match = [r[col2_index] for r in sheet2.iter_rows(values_only=True) if r[col2_index]]

        # Set up result headers
        result_sheet['A1'].value = col1
        result_sheet['B1'].value = f"Matched {col2}"
        result_sheet['C1'].value = "Match Score"

        res = collections.defaultdict(list)

        # Skip header row
        rows = list(sheet1.iter_rows(values_only=True))
        if rows:  # Skip header
            for t in rows[1:]:
                value = t[col1_index]
                if value:
                    res[value] = process.extract(value, values_to_match, limit=1)
                    result_sheet[f'A{row_tracker + 1}'].value = value
                    if res[value]:
                        result_sheet[f'B{row_tracker + 1}'].value = res[value][0][0]
                        result_sheet[f'C{row_tracker + 1}'].value = res[value][0][1]
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
control_frame = ttk.LabelFrame(main_horizontal_frame, text="Controls", width=280)
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


files_scrollbar_y.pack(side="right", fill="y")
files_scrollbar_x.pack(side="bottom", fill="x")
files_canvas.pack(side="left", fill="both", expand=True)


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

# Sheet selection for File 1
sheet1_label = tk.Label(control_frame, text="Sheet for File 1:")
sheet1_label.grid(row=3, column=0, sticky="w", pady=5, padx=5)
selected_sheet1 = StringVar(root)
selected_sheet1.set("Sheet1")  # Default value
sheet1_dropdown = tk.OptionMenu(control_frame, selected_sheet1, "Sheet1")
sheet1_dropdown.grid(row=3, column=1, pady=5, padx=5, sticky="ew")
selected_sheet1.trace('w', on_sheet1_selected)  # Call function when selection changes

sheet2_label = tk.Label(control_frame, text="Sheet for File 2:")
sheet2_label.grid(row=6, column=0, sticky="w", pady=5, padx=5)
selected_sheet2 = StringVar(root)
selected_sheet2.set("Sheet1")  # Default value
sheet2_dropdown = tk.OptionMenu(control_frame, selected_sheet2, "Sheet1")
sheet2_dropdown.grid(row=6, column=1, pady=5, padx=5, sticky="ew")
selected_sheet2.trace('w', on_sheet2_selected)  # Call function when selection changes

label_col_file1 = tk.Label(control_frame, text="Columns for File 1")
label_col_file1.grid(row=4, column=0, sticky="w", pady=5, padx=5)
dropdown_placeholder1 = tk.Label(control_frame, text="Email", relief="raised", width=15)
dropdown_placeholder1.grid(row=4, column=1, pady=5, padx=5, sticky="ew")


label_col_file2 = tk.Label(control_frame, text="Columns for File 2")
label_col_file2.grid(row=5, column=0, sticky="w", pady=5, padx=5)
dropdown_placeholder2 = tk.Label(control_frame, text="Email", relief="raised", width=15)
dropdown_placeholder2.grid(row=5, column=1, pady=5, padx=5, sticky="ew")

match_btn = tk.Button(control_frame, text="Start Matching", command=start_matching, bg="green", fg="white", height=2)
match_btn.grid(row=7, column=0, columnspan=2, pady=10, padx=5, sticky="ew")

root.mainloop()