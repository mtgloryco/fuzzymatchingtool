import openpyxl as op
import pandas as pd
from rapidfuzz import process
from tkinter import filedialog, messagebox, StringVar, simpledialog
import tkinter as tk
from datetime import datetime
from tkinter import ttk
import openpyxl.styles
import collections
import ctypes
from rounded_button import round_button

# def select_column(event):
#     region = 

def get_sheet_names(file_path):
    """Get all sheet names from an Excel file."""
    workbook = op.load_workbook(file_path, read_only=True)
    return workbook.sheetnames

def update_sheet_dropdown(file_path, is_file1=True):
    """Update the sheet dropdown with sheet names from the selected file."""
    try:
        sheet_names = get_sheet_names(file_path)
        
        if is_file1:
            
            sheet1_dropdown['menu'].delete(0, 'end')
            
            selected_sheet1.set(sheet_names[0])
            
            for sheet in sheet_names:
                sheet1_dropdown['menu'].add_command(label=sheet, 
                                                   command=lambda s=sheet: selected_sheet1.set(s))
        else:
            
            sheet2_dropdown['menu'].delete(0, 'end')
            
            selected_sheet2.set(sheet_names[0])
            
            for sheet in sheet_names:
                sheet2_dropdown['menu'].add_command(label=sheet, 
                                                   command=lambda s=sheet: selected_sheet2.set(s))
    except Exception as e:
        messagebox.showerror("Error", f"Could not read sheets: {e}")

def drop_down_menu(columns, file_label):
    global dropdown1, dropdown2, selected_col1, selected_col2

    if 'selected_col1' not in globals():
        selected_col1 = StringVar(root)
    if 'selected_col2' not in globals():
        selected_col2 = StringVar(root)

    # Normalize column names
    options = [col.strip() for col in columns]

    if file_label == "File 1":
        selected_col1.set(options[0] if options else "")  # Set the first column as default

        if 'dropdown1' in globals() and dropdown1 is not None:
            dropdown1.destroy()

        dropdown1 = tk.OptionMenu(control_frame, selected_col1, *options)
        dropdown1.grid(row=6, column=1, pady=5, padx=5, sticky="ew")

    else:
        selected_col2.set(options[0] if options else "")  # Set the first column as default

        if 'dropdown2' in globals() and dropdown2 is not None:
            dropdown2.destroy()

        dropdown2 = tk.OptionMenu(control_frame, selected_col2, *options)
        dropdown2.grid(row=7, column=1, pady=5, padx=5, sticky="ew")

def load_sheet_data(file_path, sheet_name, tree, file_label):
    """Load data from specified sheet into the treeview."""
    try:
        sheet = pd.read_excel(file_path, sheet_name)
        cols = sheet.columns.tolist()
        tree.config(columns=cols)
        drop_down_menu(cols, file_label)  # Update the dropdown menu with columns

        for item in tree.get_children():
            tree.delete(item)

        for col in cols:
            tree.heading(col, text=col)
            tree.column(col, width=100, minwidth=100)

        for _, row in sheet.iterrows():
            tree.insert("", "end", values=list(row))

    except Exception as e:
        messagebox.showerror("Error", f"Could not load sheet data: {e}")

def on_sheet1_selected(*args):
    """Called when a sheet is selected for file 1."""
    if hasattr(root, 'file1_path') and root.file1_path:
        
        load_sheet_data(root.file1_path, selected_sheet1.get(), tree1, "File 1")
        
        try:
            
            workbook = op.load_workbook(root.file1_path)
            
            workbook.active = workbook[selected_sheet1.get()]
            
            workbook.save(root.file1_path)
        except Exception as e:
            
            print(f"Could not set active sheet in file: {e}")

def on_sheet2_selected(*args):
    """Called when a sheet is selected for file 2."""
    if hasattr(root, 'file2_path') and root.file2_path:
        
        load_sheet_data(root.file2_path, selected_sheet2.get(), tree2, "File 2")
        
        
        try:
            
            workbook = op.load_workbook(root.file2_path)
            
            workbook.active = workbook[selected_sheet2.get()]
            
            workbook.save(root.file2_path)
        except Exception as e:
            
            print(f"Could not set active sheet in file: {e}")

def select_file1():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if file_path:
        root.file1_path = file_path
        filename = file_path.split('/')[-1]
        frame1.config(tex=filename)
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
        frame2.config(tex=filename)
        file2_label.config(text=f"Selected: {filename}")
        
        # Update sheet dropdown
        update_sheet_dropdown(file_path, is_file1=False)
        
        # Load the first sheet data
        if hasattr(root, 'selected_sheet2'):
            load_sheet_data(file_path, selected_sheet2.get(), tree2, "File 2")
    
def start_matching():
    try:
        # Validation checks
        if not hasattr(root, 'file1_path') or not hasattr(root, 'file2_path'):
            raise ValueError("Please select both Excel files first")

        if not selected_col1.get() or not selected_col2.get():
            raise ValueError("Please select columns for matching from both files")

        # Load data to validate selected columns
        df1 = pd.read_excel(root.file1_path, selected_sheet1.get())
        df2 = pd.read_excel(root.file2_path, selected_sheet2.get())

        # Normalize column names
        df1.columns = df1.columns.str.strip().str.lower()
        df2.columns = df2.columns.str.strip().str.lower()

        col1 = selected_col1.get().strip().lower()
        col2 = selected_col2.get().strip().lower()

        if col1 not in df1.columns:
            raise ValueError(f"Selected column '{selected_col1.get()}' is not valid for File 1")
        if col2 not in df2.columns:
            raise ValueError(f"Selected column '{selected_col2.get()}' is not valid for File 2")

        # Get filename from user
        custom_filename = simpledialog.askstring("Input", "Enter the name for the result file:")
        if not custom_filename:
            raise ValueError("Filename cannot be empty")
            
        # Get column names and threshold
        col1 = selected_col1.get()
        col2 = selected_col2.get()
        
        try:
            threshold = float(threshold_var.get())
        except ValueError:
            threshold = 80.0  # Default if invalid input

        # Load workbooks
        file1 = op.load_workbook(root.file1_path)
        sheet1 = file1[selected_sheet1.get()]
        file2 = op.load_workbook(root.file2_path)
        sheet2 = file2[selected_sheet2.get()]

        # Get column indices
        df1 = pd.read_excel(root.file1_path, selected_sheet1.get())
        df2 = pd.read_excel(root.file2_path, selected_sheet2.get())
        
        if col1 not in df1.columns:
            raise ValueError(f"Column '{col1}' not found in File 1")
        if col2 not in df2.columns:
            raise ValueError(f"Column '{col2}' not found in File 2")
            
        col1_index = df1.columns.get_loc(col1)
        col2_index = df2.columns.get_loc(col2)

        # Create workbook and sheets
        result = op.Workbook()
        filtered_sheet = result.active
        filtered_sheet.title = "Filtered Results"
        all_results_sheet = result.create_sheet(title="All Results")

        # Set up headers for both sheets
        for sheet in [filtered_sheet, all_results_sheet]:
            sheet['A1'].value = col1
            sheet['B1'].value = f"Matched {col2}"
            sheet['C1'].value = "Match Score"

        # Color definitions
        green_fill = op.styles.PatternFill(start_color='90EE90', end_color='90EE90', fill_type='solid')
        orange_fill = op.styles.PatternFill(start_color='FFB347', end_color='FFB347', fill_type='solid')
        red_fill = op.styles.PatternFill(start_color='FF6B6B', end_color='FF6B6B', fill_type='solid')

        # Initialize counters and get values to match
        row_tracker_filtered = 1
        row_tracker_all = 1
        values_to_match = [str(r[col2_index]) for r in sheet2.iter_rows(values_only=True) if r[col2_index]]

        # Process matches
        for t in sheet1.iter_rows(values_only=True, min_row=2):  # Skip header row
            value = t[col1_index]
            if value:
                # Convert to string to ensure compatibility
                str_value = str(value)
                matches = process.extract(str_value, values_to_match, limit=1)
                if matches and len(matches) > 0:
                    # Correctly handle the tuple unpacking - rapidfuzz returns (match, score, index)
                    match_tuple = matches[0]
                    # Extract just the first two elements we need
                    match_value = match_tuple[0]
                    match_score = match_tuple[1]
                    
                    # Add to filtered sheet if meets threshold
                    if match_score >= threshold:
                        row_tracker_filtered += 1
                        filtered_sheet.cell(row=row_tracker_filtered, column=1, value=value)
                        filtered_sheet.cell(row=row_tracker_filtered, column=2, value=match_value)
                        filtered_sheet.cell(row=row_tracker_filtered, column=3, value=match_score)
                    
                    # Add to all results sheet with color coding
                    row_tracker_all += 1
                    cells = [
                        all_results_sheet.cell(row=row_tracker_all, column=1, value=value),
                        all_results_sheet.cell(row=row_tracker_all, column=2, value=match_value),
                        all_results_sheet.cell(row=row_tracker_all, column=3, value=match_score)
                    ]
                    
                    fill = (green_fill if match_score >= threshold 
                           else orange_fill if match_score >= threshold * 0.8 
                           else red_fill)
                    
                    for cell in cells:
                        cell.fill = fill

        # Auto-adjust column widths and save
        for sheet in [filtered_sheet, all_results_sheet]:
            for col in sheet.columns:
                max_length = max(len(str(cell.value or "")) for cell in col)
                sheet.column_dimensions[col[0].column_letter].width = max_length + 2

        # Create safe filename
        result_sheet = result.active
        result_sheet.title = "Result"
        if selected_col1.get() not in dropdown1['menu'].entrycget(0, 'label'):
            raise ValueError("Selected column is not valid.")
        options = [col for col in df1.columns]  # Ensure options is defined
        index_of_item = options.index(selected_col1.get())
        row_tracker = 1
        email2 = [r[index_of_item] for r in sheet2.iter_rows(values_only=True) if r[index_of_item]]

       
        result_sheet['A1'].value = f"{selected_col1.get()} from first file "
        result_sheet['B1'].value = f"{selected_col2.get()} from second file"
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
        result_filename = f"{custom_filename}_{current_time}.xlsx"
        
        # Try to save the file
        try:
            result.save(result_filename)
            messagebox.showinfo("Success", 
                f"Matching complete!\nResults saved as '{result_filename}'\n" +
                f"Sheet 1: Matches >= {threshold}%\n" +
                "Sheet 2: All matches with color coding:\n" +
                "  Green: Meets threshold\n" +
                "  Orange: Near threshold\n" +
                "  Red: Below threshold")
        except PermissionError:
            messagebox.showerror("Error", 
                f"Could not save the file. The file '{result_filename}' may be in use or you don't have permission to write to this location.")
        except Exception as e:
            messagebox.showerror("Error", f"Could not save the file: {e}")
        fn = f"{custom_filename}_{current_time}.xlsx"
        result.save(fn)
        messagebox.showinfo("Success", f"Matching complete! Results saved as '{fn}'.")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

    print(f"Selected column for File 1: {selected_col1.get()}")
    print(f"Selected column for File 2: {selected_col2.get()}")
    print(f"Columns in File 1: {list(df1.columns)}")
    print(f"Columns in File 2: {list(df2.columns)}")
    print(f"Selected column for File 1: {selected_col1.get()}")
    print(f"Selected column for File 2: {selected_col2.get()}")


try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1) 
except:
    pass

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



def on_column_click(event):
    global selected_col1
    column_id = event.widget.identify_column(event.x)
    column_index = int(column_id.strip('#')) - 1 
    column_name = event.widget.heading(column_index, 'text')
    selected_col1.set(column_name)
    # print(f"Column {column_name} clicked!")
def on_column_click_2(event):
    global selected_col2
    column_id = tree2.identify_column(event.x)
    column_index = int(column_id.strip('#')) - 1 
    column_name = tree2.heading(column_index, 'text')
    selected_col2.set(column_name)
    # print(f"Column {column_name} clicked!")


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

root.bind('<Up>', lambda event: files_canvas.yview_scroll(-1, "units"))
root.bind('<Down>', lambda event: files_canvas.yview_scroll(1, "units"))
root.bind('<Left>', lambda event: files_canvas.xview_scroll(-1, "units"))
root.bind('<Right>', lambda event: files_canvas.xview_scroll(1, "units"))
root.bind('<Prior>', lambda event: files_canvas.yview_scroll(-1, "pages"))
root.bind('<Next>', lambda event: files_canvas.yview_scroll(1, "pages"))

style = ttk.Style()
style.theme_use("vista")


# Change Treeview heading text color (foreground)
style.configure("Treeview.Heading", foreground="#7b6cd9", font=('Helvetica', 10, 'bold'))

# First tree with scrollbars
frame1 = tk.LabelFrame(files_scroll_frame, text="First Excel File", background="#7b6cd9", foreground="white", font=("Poppins", 8, "bold"))
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
tree1.bind("<Button-1>", on_column_click)
tree1.pack(fill="both", expand=True)
tree1["show"] = "headings"

tree1_scroll_y.config(command=tree1.yview)
tree1_scroll_x.config(command=tree1.xview)

# Second tree with scrollbars
frame2 = tk.LabelFrame(files_scroll_frame, text="Second Excel File", background="#7b6cd9", foreground="white", font=("Poppins", 8, "bold"))
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
tree2.bind("<Button-1>", on_column_click_2)
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

# File 2 selection label - full width
file2_label = tk.Label(control_frame, text="Selected: ", bg="light gray", relief="sunken", height=2)
file2_label.grid(row=2, column=0, columnspan=2, pady=5, padx=5, sticky="ew")

# Choose File 2 button - full width
file2_btn = tk.Button(control_frame, text="Choose File 2", command=select_file2, height=2)
file2_btn.grid(row=3, column=0, columnspan=2, pady=5, padx=5, sticky="ew")

# Sheet selection for File 1
sheet1_label = tk.Label(control_frame, text="Sheet for File 1:")
sheet1_label.grid(row=4, column=0, sticky="w", pady=5, padx=5)
selected_sheet1 = StringVar(root)
selected_sheet1.set("Sheet1")  # Default value
sheet1_dropdown = tk.OptionMenu(control_frame, selected_sheet1, "Sheet1")
sheet1_dropdown.grid(row=4, column=1, pady=5, padx=5, sticky="ew")
selected_sheet1.trace('w', on_sheet1_selected)  # Call function when selection changes

# Sheet selection for File 2
sheet2_label = tk.Label(control_frame, text="Sheet for File 2:")
sheet2_label.grid(row=5, column=0, sticky="w", pady=5, padx=5)
selected_sheet2 = StringVar(root)
selected_sheet2.set("Sheet1")  # Default value
sheet2_dropdown = tk.OptionMenu(control_frame, selected_sheet2, "Sheet1")
sheet2_dropdown.grid(row=5, column=1, pady=5, padx=5, sticky="ew")
selected_sheet2.trace('w', on_sheet2_selected)  # Call function when selection changes

# Column selection for File 1
label_col_file1 = tk.Label(control_frame, text="Column for File 1:")
label_col_file1.grid(row=6, column=0, sticky="w", pady=5, padx=5)
dropdown_placeholder1 = tk.Label(control_frame, text="Select column", relief="raised", width=15)
dropdown_placeholder1.grid(row=6, column=1, pady=5, padx=5, sticky="ew")

# Column selection for File 2
label_col_file2 = tk.Label(control_frame, text="Column for File 2:")
label_col_file2.grid(row=7, column=0, sticky="w", pady=5, padx=5)
dropdown_placeholder2 = tk.Label(control_frame, text="Select column", relief="raised", width=15)
dropdown_placeholder2.grid(row=7, column=1, pady=5, padx=5, sticky="ew")

# Match settings
threshold_frame = ttk.LabelFrame(control_frame, text="Match Settings")
threshold_frame.grid(row=8, column=0, columnspan=2, pady=10, padx=5, sticky="ew")

threshold_label = ttk.Label(threshold_frame, text="Minimum Match %:")
threshold_label.pack(side="left", padx=5)

threshold_var = tk.StringVar(value="80")  # Default 80%
threshold_spinbox = ttk.Spinbox(
    threshold_frame,
    from_=0,
    to=100,
    textvariable=threshold_var,
    width=5
)
threshold_spinbox.pack(side="right", padx=5, pady=5)

# Match button

match_btn = round_button(control_frame, radius=25, fill="#7b6cd9",font=("Poppins", 9, "bold"), width=170, text="Start matching", command=start_matching)

# match_btn = tk.Button(control_frame, text="Start Matching", command=start_matching, bg="green", fg="#7b6cd9", height=2)
match_btn.grid(row=9, column=0, columnspan=2, pady=20, padx=10, sticky="nsew")

# Add a status label
status_label = tk.Label(control_frame, text="Ready", bd=1, relief="sunken", anchor="w")
status_label.grid(row=10, column=0, columnspan=2, pady=5, padx=5, sticky="ew")

# Start the application
root.mainloop()