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
from openpyxl import load_workbook

def get_sheet_names(file_path):
    """Get all sheet names from an Excel file."""
    try:
        workbook = op.load_workbook(file_path, read_only=True)
        return workbook.sheetnames
    except Exception as e:
        messagebox.showerror("Error", f"Could not read sheets: {e}")
        return ["Sheet1"]

def get_tables_in_sheet(file_path, sheet_name):
    """Return a list of real Excel tables (ListObjects) in the given sheet."""
    try:
        wb = load_workbook(filename=file_path)
        sheet = wb[sheet_name]
        tables = []
        for table in sheet.tables.values():
            tables.append(table.name)
        if not tables:
            tables.append("Full Sheet")
        return tables
    except Exception as e:
        messagebox.showerror("Error", f"Could not read tables: {e}")
        return ["Full Sheet"]

def get_columns_in_table(file_path, sheet_name, table_name):
    """Get columns from a specific table or the entire sheet."""
    try:
        if table_name == "Full Sheet":
            # Get all columns from the sheet
            df = pd.read_excel(file_path, sheet_name)
            return list(df.columns)
        elif table_name.startswith("Table:"):
            # Extract table name without the prefix
            pure_table_name = table_name.replace("Table: ", "").replace(" ++", "")
            
            # Try to read table data
            workbook = op.load_workbook(file_path)
            sheet = workbook[sheet_name]
            
            # Find the table range
            table_range = None
            
            # Check for Excel ListObjects first (native tables)
            if hasattr(sheet, '_tables'):
                for table in sheet._tables:
                    table_display_name = getattr(table, 'displayName', None)
                    if table_display_name == pure_table_name:
                        table_range = table.ref
                        break
            
            # If not found, check defined names
            if not table_range and hasattr(workbook, 'defined_names'):
                if pure_table_name in workbook.defined_names:
                    for sheet_title, cell_range in workbook.defined_names[pure_table_name].destinations:
                        if sheet_title == sheet_name:
                            table_range = cell_range
                            break
            
            # If we found a range, extract columns
            if table_range:
                # Read the specific range from the Excel file
                df = pd.read_excel(file_path, sheet_name, skiprows=None)
                
                # Convert range to excel coordinates
                # This is a simplified approach, might need adjustment for complex ranges
                if ':' in table_range:
                    start_cell, end_cell = table_range.split(':')
                    
                    # Read the header row from the range
                    cell_range = sheet[table_range]
                    if isinstance(cell_range, tuple) and len(cell_range) > 0:
                        header_row = cell_range[0]  # First row contains headers
                        return [cell.value for cell in header_row if cell.value]
                
                # Fallback: return all columns from the sheet
                return list(df.columns)
            else:
                # If table not found, return all columns from the sheet
                df = pd.read_excel(file_path, sheet_name)
                return list(df.columns)
    except Exception as e:
        messagebox.showerror("Error", f"Could not read columns: {e}")
        return []

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
            # Update tables after sheet selection
            update_table_dropdown(file_path, selected_sheet1.get(), is_file1=True)
        else:
            sheet2_dropdown['menu'].delete(0, 'end')
            selected_sheet2.set(sheet_names[0])
            
            for sheet in sheet_names:
                sheet2_dropdown['menu'].add_command(label=sheet, 
                                                  command=lambda s=sheet: selected_sheet2.set(s))
            # Update tables after sheet selection
            update_table_dropdown(file_path, selected_sheet2.get(), is_file1=False)
    except Exception as e:
        messagebox.showerror("Error", f"Could not read sheets: {e}")

def update_table_dropdown(file_path, sheet_name, is_file1=True):
    """Update the table dropdown with tables from the selected sheet."""
    try:
        table_names = get_tables_in_sheet(file_path, sheet_name)
        
        if is_file1:
            table1_dropdown['menu'].delete(0, 'end')
            selected_table1.set(table_names[0])
            
            for table in table_names:
                table1_dropdown['menu'].add_command(label=table, 
                                                  command=lambda t=table: selected_table1.set(t))
            # Update columns after table selection
            update_column_dropdown(file_path, sheet_name, selected_table1.get(), is_file1=True)
        else:
            table2_dropdown['menu'].delete(0, 'end')
            selected_table2.set(table_names[0])
            
            for table in table_names:
                table2_dropdown['menu'].add_command(label=table, 
                                                  command=lambda t=table: selected_table2.set(t))
            # Update columns after table selection
            update_column_dropdown(file_path, sheet_name, selected_table2.get(), is_file1=False)
    except Exception as e:
        messagebox.showerror("Error", f"Could not read tables: {e}")

def update_column_dropdown(file_path, sheet_name, table_name, is_file1=True):
    """Update the column dropdown with columns from the selected table."""
    try:
        columns = get_columns_in_table(file_path, sheet_name, table_name)
        
        if is_file1:
            # Update dropdown menu
            column1_dropdown['menu'].delete(0, 'end')
            if columns:
                selected_col1.set(columns[0])
                for col in columns:
                    column1_dropdown['menu'].add_command(label=col, 
                                                     command=lambda c=col: selected_col1.set(c))
                # Update status
                status_label.config(text=f"Ready - File 1 column set to: {selected_col1.get()}")
            else:
                selected_col1.set("")
                status_label.config(text="No columns found in selected table")
                
            # Update treeview to show selected table
            load_table_data(file_path, sheet_name, table_name, tree1, "File 1")
            
        else:
            # Update dropdown menu
            column2_dropdown['menu'].delete(0, 'end')
            if columns:
                selected_col2.set(columns[0])
                for col in columns:
                    column2_dropdown['menu'].add_command(label=col, 
                                                     command=lambda c=col: selected_col2.set(c))
                # Update status
                status_label.config(text=f"Ready - File 2 column set to: {selected_col2.get()}")
            else:
                selected_col2.set("")
                status_label.config(text="No columns found in selected table")
                
            # Update treeview to show selected table
            load_table_data(file_path, sheet_name, table_name, tree2, "File 2")
            
    except Exception as e:
        messagebox.showerror("Error", f"Could not update columns: {e}")

def load_table_data(file_path, sheet_name, table_name, tree, file_label):
    """Load only the selected table's data into the treeview."""
    try:
        wb = load_workbook(filename=file_path)
        ws = wb[sheet_name]
        if table_name == "Full Sheet":
            df = pd.read_excel(file_path, sheet_name)
        else:
            # Find the table by name and get its range
            table = ws.tables[table_name]
            data = ws[table.ref]
            rows = [[cell.value for cell in row] for row in data]
            df = pd.DataFrame(rows[1:], columns=rows[0])
        # Update treeview
        cols = df.columns.tolist()
        tree.config(columns=cols)
        for item in tree.get_children():
            tree.delete(item)
        for col in cols:
            tree.heading(col, text=col)
            tree.column(col, width=100, minwidth=100)
        for _, row in df.iterrows():
            tree.insert("", "end", values=list(row))
    except Exception as e:
        messagebox.showerror("Error", f"Could not load table data: {e}")

def on_sheet1_selected(*args):
    """Called when a sheet is selected for file 1."""
    if hasattr(root, 'file1_path') and root.file1_path:
        # Update tables in the selected sheet
        update_table_dropdown(root.file1_path, selected_sheet1.get(), is_file1=True)
        
        try:
            workbook = op.load_workbook(root.file1_path)
            workbook.active = workbook[selected_sheet1.get()]
            workbook.save(root.file1_path)
        except Exception as e:
            print(f"Could not set active sheet in file: {e}")

def on_sheet2_selected(*args):
    """Called when a sheet is selected for file 2."""
    if hasattr(root, 'file2_path') and root.file2_path:
        # Update tables in the selected sheet
        update_table_dropdown(root.file2_path, selected_sheet2.get(), is_file1=False)
        
        try:
            workbook = op.load_workbook(root.file2_path)
            workbook.active = workbook[selected_sheet2.get()]
            workbook.save(root.file2_path)
        except Exception as e:
            print(f"Could not set active sheet in file: {e}")

def on_table1_selected(*args):
    """Called when a table is selected for file 1."""
    if hasattr(root, 'file1_path') and root.file1_path:
        # Only show the selected table's data
        load_table_data(root.file1_path, selected_sheet1.get(), selected_table1.get(), tree1, "File 1")
        update_column_dropdown(root.file1_path, selected_sheet1.get(), selected_table1.get(), is_file1=True)

def on_table2_selected(*args):
    """Called when a table is selected for file 2."""
    if hasattr(root, 'file2_path') and root.file2_path:
        # Only show the selected table's data
        load_table_data(root.file2_path, selected_sheet2.get(), selected_table2.get(), tree2, "File 2")
        update_column_dropdown(root.file2_path, selected_sheet2.get(), selected_table2.get(), is_file1=False)

def select_file1():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if file_path:
        root.file1_path = file_path
        filename = file_path.split('/')[-1]
        frame1.config(text=filename)
        file1_label.config(text=f"Selected: {filename}")
        # Update sheet dropdown
        update_sheet_dropdown(file_path, is_file1=True)
    
def select_file2():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if file_path:
        root.file2_path = file_path
        filename = file_path.split('/')[-1]
        frame2.config(text=filename)
        file2_label.config(text=f"Selected: {filename}")
        
        # Update sheet dropdown
        update_sheet_dropdown(file_path, is_file1=False)

def highlight_column(tree, column_index):
    """Highlight only the selected column in the Treeview."""
    # Reset all cell styles
    for item in tree.get_children():
        tree.item(item, tags=())
    
    # Create a tag for this specific column
    tag_name = f"col_{column_index}"
    
    # Configure the tag style
    tree.tag_configure(tag_name, background="light gray")
    
    # Apply the tag to each cell in the column
    for item in tree.get_children():
        # Get all values for this row
        values = tree.item(item, "values")
        if values and len(values) > column_index:
            # Create a new item with the tag applied only to the selected column
            tree.item(item, tags=(tag_name,))

def on_column_click(event):
    """Handle column click for File 1."""
    column_id = event.widget.identify_column(event.x)
    if column_id:  # Check if a valid column was clicked
        try:
            column_index = int(column_id.strip('#')) - 1
            if column_index >= 0:  # Ensure valid column index
                column_name = tree1.heading(column_id)['text']
                if column_name:  # Ensure column name exists
                    selected_col1.set(column_name)
                    status_label.config(text=f"File 1 column selected: {column_name}")
                    
                    # Add custom rendering for the column cells
                    for col_id in tree1["columns"]:
                        col_idx = tree1["columns"].index(col_id)
                        tree1.tag_configure(f"col_{col_idx}", background="white")
                    
                    tree1.tag_configure(f"col_{column_index}", background="light gray")
                    
                    # Apply tags to all rows
                    for item_id in tree1.get_children():
                        tree1.item(item_id, tags=(f"col_{column_index}",))
                    
                    # Apply styling to the individual cells
                    style_column_cells(tree1, column_index)
        except Exception as e:
            print(f"Error in column selection: {e}")

def on_column_click_2(event):
    """Handle column click for File 2."""
    column_id = event.widget.identify_column(event.x)
    if column_id:  # Check if a valid column was clicked
        try:
            column_index = int(column_id.strip('#')) - 1
            if column_index >= 0:  # Ensure valid column index
                column_name = tree2.heading(column_id)['text']
                if column_name:  # Ensure column name exists
                    selected_col2.set(column_name)
                    status_label.config(text=f"File 2 column selected: {column_name}")
                    
                    # Add custom rendering for the column cells
                    for col_id in tree2["columns"]:
                        col_idx = tree2["columns"].index(col_id)
                        tree2.tag_configure(f"col_{col_idx}", background="white")
                    
                    tree2.tag_configure(f"col_{column_index}", background="light gray")
                    
                    # Apply tags to all rows
                    for item_id in tree2.get_children():
                        tree2.item(item_id, tags=(f"col_{column_index}",))
                    
                    # Apply styling to the individual cells
                    style_column_cells(tree2, column_index)
        except Exception as e:
            print(f"Error in column selection: {e}")

def style_column_cells(tree, column_index):
    """Style individual cells in the selected column."""
    # This function will be used for custom styling of individual cells
    # Since tkinter treeview doesn't support individual cell styling directly,
    # we use a workaround with item tags and rendering 
    
    # First, clear existing column styling
    for i in range(len(tree["columns"])):
        tag_name = f"col_{i}"
        tree.tag_configure(tag_name, background="white")
    
    # Set the style for the selected column
    tag_name = f"col_{column_index}"
    tree.tag_configure(tag_name, background="light gray")
    
    # Create a custom display column that shows the selection
    def fixed_map(option):
        # Fix for setting text colour for Tkinter ttk Treeview
        return [elm for elm in style.map("Treeview", query_opt=option)
                if elm[:2] != ("!disabled", "!selected")]
    
    style = ttk.Style()
    style.map("Treeview", 
              foreground=fixed_map("foreground"),
              background=fixed_map("background"))

def start_matching():
    try:
        status_label.config(text="Processing...")
        root.update()
        
        # Validation checks
        if not hasattr(root, 'file1_path') or not hasattr(root, 'file2_path'):
            raise ValueError("Please select both Excel files first")

        if not selected_col1.get() or not selected_col2.get():
            raise ValueError("Please select columns for matching from both files")

        # Load data based on selected tables and sheets
        if selected_table1.get() == "Full Sheet":
            df1 = pd.read_excel(root.file1_path, selected_sheet1.get())
        else:
            # For now, we read the entire sheet - in a more advanced version,
            # you could read only the specified table range
            df1 = pd.read_excel(root.file1_path, selected_sheet1.get())
            
        if selected_table2.get() == "Full Sheet":
            df2 = pd.read_excel(root.file2_path, selected_sheet2.get())
        else:
            # For now, we read the entire sheet - in a more advanced version,
            # you could read only the specified table range
            df2 = pd.read_excel(root.file2_path, selected_sheet2.get())

        # Debug prints
        print(f"Selected column for File 1: {selected_col1.get()}")
        print(f"Selected column for File 2: {selected_col2.get()}")
        print(f"Columns in File 1: {list(df1.columns)}")
        print(f"Columns in File 2: {list(df2.columns)}")

        # Convert column names to strings and normalize them for matching
        df1.columns = [str(col).strip() for col in df1.columns]
        df2.columns = [str(col).strip() for col in df2.columns]

        col1 = selected_col1.get().strip()
        col2 = selected_col2.get().strip()

        # First try exact match
        if col1 not in df1.columns:
            # Try case-insensitive match
            col1_match = next((col for col in df1.columns if col.lower() == col1.lower()), None)
            if col1_match:
                col1 = col1_match
            else:
                raise ValueError(f"Selected column '{selected_col1.get()}' is not valid for File 1")
        
        if col2 not in df2.columns:
            # Try case-insensitive match
            col2_match = next((col for col in df2.columns if col.lower() == col2.lower()), None)
            if col2_match:
                col2 = col2_match
            else:
                raise ValueError(f"Selected column '{selected_col2.get()}' is not valid for File 2")

        # Get filename from user
        custom_filename = simpledialog.askstring("Input", "Enter the name for the result file:")
        if not custom_filename:
            raise ValueError("Filename cannot be empty")
            
        # Get threshold
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

        # Auto-adjust column widths
        for sheet in [filtered_sheet, all_results_sheet]:
            for col in sheet.columns:
                max_length = max(len(str(cell.value or "")) for cell in col)
                sheet.column_dimensions[col[0].column_letter].width = max_length + 2

        # Create timestamp for filename
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
            status_label.config(text=f"Matching complete. Results saved as '{result_filename}'")
        except PermissionError:
            messagebox.showerror("Error", 
                f"Could not save the file. The file '{result_filename}' may be in use or you don't have permission to write to this location.")
            status_label.config(text="Error: Could not save file (permission denied)")
        except Exception as e:
            messagebox.showerror("Error", f"Could not save the file: {e}")
            status_label.config(text=f"Error: {str(e)}")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")
        status_label.config(text=f"Error: {str(e)}")


# Set DPI awareness for better display on Windows
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1) 
except:
    pass

# Create main window
root = tk.Tk()
root.title("Excel Fuzzy Matching")
root.geometry("1400x800")  # Wider window to accommodate three panels

# Apply style
style = ttk.Style()
style.theme_use("vista")
style.configure("Treeview.Heading", foreground="#7b6cd9", font=('Helvetica', 10, 'bold'))

# Create main horizontal frame to hold the three sections
main_horizontal_frame = ttk.Frame(root)
main_horizontal_frame.pack(fill="both", expand=True)

# Create sheet selection frame on the left
sheet_frame = ttk.LabelFrame(main_horizontal_frame, text="File Selection", width=250)
sheet_frame.pack(side="left", fill="y", padx=5, pady=5)
sheet_frame.pack_propagate(False)  # Prevent the frame from shrinking

# Create files frame in the center (for tree views)
files_frame = ttk.LabelFrame(main_horizontal_frame, text="Excel Data")
files_frame.pack(side="left", fill="both", expand=True, padx=5, pady=5)

# Create controls frame on the right
control_frame = ttk.LabelFrame(main_horizontal_frame, text="Hierarchical Selection", width=300)
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

# Configure scrolling functions
def _on_mousewheel(event):
    files_canvas.yview_scroll(int(-1*(event.delta/120)), "units")

def _on_shift_mousewheel(event):
    files_canvas.xview_scroll(int(-1*(event.delta/120)), "units")

files_canvas.bind_all("<MouseWheel>", _on_mousewheel)
files_canvas.bind_all("<Shift-MouseWheel>", _on_shift_mousewheel)

# Add keyboard binding for better scrolling
root.bind('<Up>', lambda event: files_canvas.yview_scroll(-1, "units"))
root.bind('<Down>', lambda event: files_canvas.yview_scroll(1, "units"))
root.bind('<Left>', lambda event: files_canvas.xview_scroll(-1, "units"))
root.bind('<Right>', lambda event: files_canvas.xview_scroll(1, "units"))
root.bind('<Prior>', lambda event: files_canvas.yview_scroll(-1, "pages"))
root.bind('<Next>', lambda event: files_canvas.yview_scroll(1, "pages"))

# CONFIGURE LEFT PANEL (FILE SELECTION) ------------------------------
sheet_frame.grid_columnconfigure(0, weight=1)

# File 1 selection label
file1_label = tk.Label(sheet_frame, text="Selected: ", bg="light gray", relief="sunken", height=2)
file1_label.grid(row=0, column=0, pady=5, padx=5, sticky="ew")

# Choose File 1 button
file1_btn = tk.Button(sheet_frame, text="Choose File 1", command=select_file1, height=2)
file1_btn.grid(row=1, column=0, pady=5, padx=5, sticky="ew")

# Add a separator
ttk.Separator(sheet_frame, orient='horizontal').grid(row=2, column=0, sticky='ew', pady=10)

# File 2 selection label
# File 2 selection label
file2_label = tk.Label(sheet_frame, text="Selected: ", bg="light gray", relief="sunken", height=2)
file2_label.grid(row=3, column=0, pady=5, padx=5, sticky="ew")

# Choose File 2 button
file2_btn = tk.Button(sheet_frame, text="Choose File 2", command=select_file2, height=2)
file2_btn.grid(row=4, column=0, pady=5, padx=5, sticky="ew")

# CONFIGURE CENTER PANEL (TREEVIEWS) ----------------------------------

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

# Create scrollbar frame2
tree2_frame = ttk.Frame(frame2)
tree2_frame.pack(fill="both", expand=True)

# Create vertical scrollbar
tree2_scroll_y = ttk.Scrollbar(tree2_frame)
tree2_scroll_y.pack(side="right", fill="y")

# Create horizontal scrollbar
tree2_scroll_x = ttk.Scrollbar(tree2_frame, orient="horizontal")
tree2_scroll_x.pack(side="bottom", fill="x")

tree2 = ttk.Treeview(tree2_frame, yscrollcommand=tree2_scroll_y.set, xscrollcommand=tree2_scroll_x.set)
tree2.bind("<Button-1>", on_column_click_2)
tree2.pack(fill="both", expand=True)
tree2["show"] = "headings"

tree2_scroll_y.config(command=tree2.yview)
tree2_scroll_x.config(command=tree2.xview)

# CONFIGURE RIGHT PANEL (HIERARCHICAL SELECTION) ----------------------------------
control_frame.grid_columnconfigure(0, weight=1)
control_frame.grid_columnconfigure(1, weight=2)

# Create File 1 hierarchical section with an inset frame
file1_hierarchy_frame = ttk.LabelFrame(control_frame, text="File 1 Hierarchy")
file1_hierarchy_frame.grid(row=0, column=0, columnspan=2, pady=5, padx=5, sticky="ew")

# Sheet selection for File 1
sheet1_label = ttk.Label(file1_hierarchy_frame, text="-- Sheet:")
sheet1_label.grid(row=0, column=0, sticky="w", pady=5, padx=5)
selected_sheet1 = StringVar(root)
selected_sheet1.set("Sheet1")  # Default value
sheet1_dropdown = ttk.OptionMenu(file1_hierarchy_frame, selected_sheet1, "Sheet1")
sheet1_dropdown.grid(row=0, column=1, pady=5, padx=5, sticky="ew")
selected_sheet1.trace('w', on_sheet1_selected)  # Call function when selection changes

# Table selection for File 1
table1_label = ttk.Label(file1_hierarchy_frame, text="    |-- Tables:")
table1_label.grid(row=1, column=0, sticky="w", pady=5, padx=5)
selected_table1 = StringVar(root)
selected_table1.set("Full Sheet")  # Default value
table1_dropdown = ttk.OptionMenu(file1_hierarchy_frame, selected_table1, "Full Sheet")
table1_dropdown.grid(row=1, column=1, pady=5, padx=5, sticky="ew")
selected_table1.trace('w', on_table1_selected)  # Call function when selection changes

# Column selection for File 1
column1_label = ttk.Label(file1_hierarchy_frame, text="        |-- Columns:")
column1_label.grid(row=2, column=0, sticky="w", pady=5, padx=5)
selected_col1 = StringVar(root)
selected_col1.set("Select column")  # Default value
column1_dropdown = ttk.OptionMenu(file1_hierarchy_frame, selected_col1, "Select column")
column1_dropdown.grid(row=2, column=1, pady=5, padx=5, sticky="ew")

# Add a separator
ttk.Separator(control_frame, orient='horizontal').grid(row=1, column=0, columnspan=2, sticky='ew', pady=10)

# Create File 2 hierarchical section with an inset frame
file2_hierarchy_frame = ttk.LabelFrame(control_frame, text="File 2 Hierarchy")
file2_hierarchy_frame.grid(row=2, column=0, columnspan=2, pady=5, padx=5, sticky="ew")

# Sheet selection for File 2
sheet2_label = ttk.Label(file2_hierarchy_frame, text="-- Sheet:")
sheet2_label.grid(row=0, column=0, sticky="w", pady=5, padx=5)
selected_sheet2 = StringVar(root)
selected_sheet2.set("Sheet1")  # Default value
sheet2_dropdown = ttk.OptionMenu(file2_hierarchy_frame, selected_sheet2, "Sheet1")
sheet2_dropdown.grid(row=0, column=1, pady=5, padx=5, sticky="ew")
selected_sheet2.trace('w', on_sheet2_selected)  # Call function when selection changes

# Table selection for File 2
table2_label = ttk.Label(file2_hierarchy_frame, text="    |-- Tables:")
table2_label.grid(row=1, column=0, sticky="w", pady=5, padx=5)
selected_table2 = StringVar(root)
selected_table2.set("Full Sheet")  # Default value
table2_dropdown = ttk.OptionMenu(file2_hierarchy_frame, selected_table2, "Full Sheet")
table2_dropdown.grid(row=1, column=1, pady=5, padx=5, sticky="ew")
selected_table2.trace('w', on_table2_selected)  # Call function when selection changes

# Column selection for File 2
column2_label = ttk.Label(file2_hierarchy_frame, text="        |-- Columns:")
column2_label.grid(row=2, column=0, sticky="w", pady=5, padx=5)
selected_col2 = StringVar(root)
selected_col2.set("Select column")  # Default value
column2_dropdown = ttk.OptionMenu(file2_hierarchy_frame, selected_col2, "Select column")
column2_dropdown.grid(row=2, column=1, pady=5, padx=5, sticky="ew")

# Add a separator
ttk.Separator(control_frame, orient='horizontal').grid(row=3, column=0, columnspan=2, sticky='ew', pady=10)

# Match settings
threshold_frame = ttk.LabelFrame(control_frame, text="Match Settings")
threshold_frame.grid(row=4, column=0, columnspan=2, pady=10, padx=5, sticky="ew")

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
match_btn = round_button(control_frame, radius=25, fill="green", font=("Poppins", 9, "bold"), width=170, text="Start matching", command=start_matching)
match_btn.grid(row=5, column=0, columnspan=2, pady=20, padx=10, sticky="nsew")

# Add a status label
status_label = tk.Label(control_frame, text="Ready", bd=1, relief="sunken", anchor="w")
status_label.grid(row=6, column=0, columnspan=2, pady=5, padx=5, sticky="ew")

# Start the application
root.mainloop()