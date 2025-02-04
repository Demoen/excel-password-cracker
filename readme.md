# Excel Sheet Password Cracker

## Overview
This Python application removes sheet protection from Excel `.xlsx` files by eliminating `<sheetProtection .../>` tags from worksheet XML files. It uses the `zipfile` module to modify the XML content within the Excel archive structure.

## Features
- Select an Excel file (.xlsx) using a graphical user interface (GUI).
- Remove sheet protection from all worksheets within the file.
- Save the modified, unprotected version as a new `.xlsx` file.

## Prerequisites
Ensure you have Python installed on your system (Python 3.x recommended). You also need the following libraries:

- `tkinter` (built-in with Python)
- `zipfile` (built-in with Python)
- `os` (built-in with Python)
- `re` (built-in with Python)

## Installation
Clone the repository or download the script file.

```sh
# Clone the repository
git clone https://github.com/Demoen/excel-password-cracker.git
cd excel-password-cracker
```

No additional dependencies are required beyond standard Python libraries.

## Usage

1. Run the script:
   ```sh
   python excel_password_cracker.py
   ```

2. In the GUI:
   - Click "Select Input File" to choose the protected `.xlsx` file.
   - Click "Unprotect Excel File" and specify a location to save the new unprotected file.
   - The modified file will be saved without sheet protection.

## How It Works
1. The script extracts the `.xlsx` file, which is essentially a ZIP archive containing XML files.
2. It searches for `<sheetProtection .../>` tags inside `xl/worksheets/*.xml` files.
3. The script removes these protection tags and reconstructs the `.xlsx` file without modifying any other content.

## Limitations
- This tool only removes sheet protection; it does not remove workbook-level passwords.
- Does not recover lost passwords; it simply removes protection tags.
- Some Excel protection mechanisms might not be affected by this method.

## License
This project is licensed under the MIT License.

## Disclaimer
This tool is intended for educational and personal use only. Ensure you have permission to modify any Excel files you process with this script.

## Author
Demoen

---
Feel free to modify and contribute to this project!
