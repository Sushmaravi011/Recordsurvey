import tkinter as tk
from tkinter import ttk
import openpyxl

# List of options for the combobox
combo_list = ["Physics", "Chemistry", "Botany", "Zoology"]
section_list = ["A", "B", "C"]

def load_data():
    path = r"C:\Users\Sushma R\Music\excel project\Record.xlsx"
    try:
        workbook = openpyxl.load_workbook(path)
        sheet = workbook.active

        list_values = list(sheet.values)
        if list_values:
            # Set column headings
            treeview["columns"] = list_values[0]
            for col_name in list_values[0]:
                treeview.heading(col_name, text=col_name)
            # Insert data into the treeview
            for value_tuple in list_values[1:]:
                treeview.insert('', tk.END, values=value_tuple)
    except Exception as e:
        print(f"Error loading data: {e}")

def insert_row():
    name = name_entry.get()
    section = section_combobox.get()
    department_status = status_combobox.get()
    record_status = "YES" if check_var.get() else "NO"
    
    path = r"C:\Users\Sushma R\Music\excel project\Record.xlsx"
    try:
        workbook = openpyxl.load_workbook(path)
        sheet = workbook.active
        row_value = [name, section, department_status, record_status]
        sheet.append(row_value)
        workbook.save(path)
        
        # Insert the new row into the treeview
        treeview.insert('', tk.END, values=row_value)
        
        # Clear the entry fields
        name_entry.delete(0, "end")
        name_entry.insert(0, "Name")
        section_combobox.set(section_list[0])  # Reset to default value after insertion
        
        status_combobox.set(combo_list[0])
        
        # Optionally, reset checkbutton to unchecked after inserting a row
        check_var.set(False)
    except Exception as e:
        print(f"Error inserting row: {e}")

# Initialize the main window
root = tk.Tk()
root.title("Default Color Mode Example")

# Frame for widgets
frame = ttk.Frame(root)
frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

widgets_frame = ttk.LabelFrame(frame, text="Insert Row")
widgets_frame.grid(row=0, column=0, padx=20, pady=10, sticky="nsew")

# Entry for name
name_entry = tk.Entry(widgets_frame)
name_entry.insert(0, "Name")
name_entry.bind("<FocusIn>", lambda e: name_entry.delete(0, 'end'))
name_entry.grid(row=0, column=0, padx=5, pady=(0,5), sticky="ew")

# Combobox for section
section_combobox = ttk.Combobox(widgets_frame, values=section_list)
section_combobox.current(0)  # Set default value
section_combobox.grid(row=1, column=0, padx=5, pady=(0,5), sticky="ew")

# Combobox for department
status_combobox = ttk.Combobox(widgets_frame, values=combo_list)
status_combobox.current(0)
status_combobox.grid(row=2, column=0, padx=5, pady=5, sticky="ew")

# Checkbutton for record status
check_var = tk.BooleanVar()
check_var.set(True)  # Set Checkbutton to be checked by default
checkbutton = tk.Checkbutton(widgets_frame, text="Record", variable=check_var)
checkbutton.grid(row=3, column=0, sticky="nsew")

# Button to insert row
insert_button = ttk.Button(widgets_frame, text="Insert", command=insert_row)
insert_button.grid(row=4, column=0, padx=5, pady=10, sticky="ew")

# Separator
separator = ttk.Separator(widgets_frame)
separator.grid(row=5, column=0, padx=(20,10), pady=10, sticky="ew")

# Frame for treeview
tree_frame = ttk.Frame(frame)
tree_frame.grid(row=0, column=1, padx=5, pady=10, sticky="nsew")

tree_scroll = ttk.Scrollbar(tree_frame)
tree_scroll.pack(side="right", fill="y")

# Treeview setup
cols = ("Name", "Section", "Department", "Record")
treeview = ttk.Treeview(tree_frame, columns=cols, show="headings", yscrollcommand=tree_scroll.set)
for col in cols:
    treeview.heading(col, text=col)
    treeview.column(col, width=100)

treeview.pack(side="left", fill="both", expand=True)
tree_scroll.config(command=treeview.yview)

# Load initial data
load_data()

# Run the application
root.mainloop()
