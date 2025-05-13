# Outlook to Excel Processing

Automation script for processing supplier Excel files sent via Outlook attachments.

The script performs the following operations:
- Connects to Outlook and scans recent emails  
- Filters messages with valid `.xlsx` / `.xlsm` attachments (must contain sheets: `DATA` and `offer`)  
- Checks Excel files for formatting errors and missing data  
- Moves files into appropriate folders (`To_process`, `Processed`, `Invalid_files`)  
- Updates log files with details of processed and invalid entries  

---

## Scripts Included

1. **`1. Choosing files location.py`** – sets up the `config.json` file with folder paths (for .exe users)  
2. **`2. Downloading_from_Outlook.py`** – downloads eligible attachments and matching `.msg` files from Outlook  
3. **`3. Processing_Excel_files.py`** – validates Excel files, logs errors, and updates the master workbook  

---

## Folder structure

```

├── Readme.md
├── 1. Choosing files location.py
├── 2. Downloading_from_Outlook.py
├── 3. Processing_Excel_files.py
├── Data
    ├── Sample Supplier Database.xlsm
    ├── config.json                     # created by "1. Choosing files location.py"
    ├── log.txt                         # created after using the script
    └── Data/
        ├── To_process/
        │   ├── Sup1.xlsx
        │   ├── Sup1.msg
        │   ├── Sup2.xlsx
        │   └── ...
        ├── Processed/
        │   └── Processed_msg/
        ├── Invalid_files/
        │   ├── 0.Invalid.xlsx
        └── tmp/
```

---

## Notes

> The script will log any errors to log.txt and keep a record of invalid entries in 0.Invalid.xlsx.

> .msg files are saved alongside attachments for reference.

> Supports repeated processing and overwrites with warnings if data already exists.

## Requirements

> Outlook (classic) installed on Windows

> Python 3.8+ recommended

> To run the scripts, the following Python packages are required:

```bash
pip install openpyxl pandas pywin32



