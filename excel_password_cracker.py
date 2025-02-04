import os
import re
import zipfile
import tkinter as tk
from tkinter import filedialog, messagebox

def remove_sheet_protection(xml_content):
    """
    Remove any self-closing <sheetProtection .../> tag from the XML content.
    """
    # This regex matches a self-closing <sheetProtection .../> tag.
    pattern = r'<sheetProtection\b[^>]*/>'
    cleaned = re.sub(pattern, '', xml_content, flags=re.DOTALL)
    return cleaned

def unprotect_excel(input_excel, output_excel):
    """
    Open the input Excel file (.xlsx), remove any <sheetProtection .../> tags from all worksheet XML files,
    and write the modified content to a new Excel file.
    """
    with zipfile.ZipFile(input_excel, 'r') as zin:
        with zipfile.ZipFile(output_excel, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                # Process only worksheet XML files (in xl/worksheets/)
                if item.filename.startswith("xl/worksheets/") and item.filename.endswith(".xml"):
                    try:
                        xml_content = data.decode('utf-8')
                    except UnicodeDecodeError:
                        xml_content = data.decode('ISO-8859-1')
                    xml_content = remove_sheet_protection(xml_content)
                    data = xml_content.encode('utf-8')
                zout.writestr(item, data)

def select_input_file():
    """Open a file dialog to select the input Excel file."""
    input_file = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel Files", "*.xlsx")]
    )
    if input_file:
        input_path_var.set(input_file)

def unprotect_file():
    """Process the selected file and save the unprotected version."""
    input_file = input_path_var.get()
    if not input_file or not os.path.isfile(input_file):
        messagebox.showerror("Error", "Please select a valid input Excel file.")
        return

    # Ask where to save the new unprotected Excel file.
    output_file = filedialog.asksaveasfilename(
        title="Save Unprotected Excel File",
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx")]
    )
    if not output_file:
        return

    try:
        unprotect_excel(input_file, output_file)
        messagebox.showinfo("Success", f"Unprotected Excel file saved as:\n{output_file}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred:\n{e}")

# --- Tkinter GUI Setup ---
root = tk.Tk()
root.title("Excel Unprotector")

# Variable to store the input file path
input_path_var = tk.StringVar()

# Create a label and entry to display the selected file path
tk.Label(root, text="Selected Excel File:").grid(row=0, column=0, padx=10, pady=10, sticky="w")
entry = tk.Entry(root, textvariable=input_path_var, width=50)
entry.grid(row=0, column=1, padx=10, pady=10)

# Button to choose the input file
select_button = tk.Button(root, text="Select Input File", command=select_input_file)
select_button.grid(row=0, column=2, padx=10, pady=10)

# Button to run the unprotection process
process_button = tk.Button(root, text="Unprotect Excel File", command=unprotect_file)
process_button.grid(row=1, column=0, columnspan=3, padx=10, pady=20)

root.mainloop()

