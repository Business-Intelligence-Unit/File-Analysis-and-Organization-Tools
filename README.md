# FBCRS Record Inventory & File Analysis Tools

## Overview

This repository contains a suite of Python utilities designed to function as a full-featured Information Governance and Records Management engine. The primary application scans local or network directories, reads various document types, automatically categorizes them based on record retention rules, and generates a formatted Excel summary report.

---

## Core Components

### 1. File Analysis App

The main application is designed to handle massive file directories with built-in crash resilience. It is available in two distinct processing modes:


> 1. Slow but Accurate Version:** Processes complete files and generates full hash fingerprints for absolute duplicate validation.


> 2. Fast but Less Accurate Version:** Optimized to scan large data drops quickly. It achieves this by reading only the first two pages of PDFs, the first two slides of PowerPoints, the first 50 paragraphs of Word documents, the first Excel sheet, and only the first 1 MB of a file for duplicate checking .



**Key Features:**


**High Performance:** Utilizes `ThreadPoolExecutor` for concurrent I/O processing and the `pyahocorasick` automaton to reduce algorithmic classification complexity, significantly speeding up large record volume scans.


**Crash Recovery (Checkpointing):** Features an incremental saving design that writes every processed row to a temporary CSV file (`_temp.csv`) in real-time. If the script crashes, data processed up to that exact moment is safely retained and automatically combined upon a successful restart.



**Graceful Error Handling:** Includes `try...except` blocks around metadata retrieval to handle inaccessible files, corrupted paths, or Windows UNC network paths that exceed the 260-character limit. The app logs an "Error" row and proceeds rather than terminating the process.



**Standardized Output:** Integrates Data Validation in the final Excel output, enforcing standardized dropdown menus for downstream auditing and applying clickable file hyperlinks.



### 2. Master Code Updater

A supplementary utility designed to safely update the live rulebook, `FBCRS_Master_Full.xlsx`, with missing Business Intelligence FBCRS codes.

**Dynamic Column Matching:** It searches the Excel file for the first column's actual name to map data correctly, preventing duplication.


**Safe Updating:** Utilizes the `.update()` function to inject new BI descriptions for existing codes or append new rows securely at the bottom. This ensures that existing, unrelated columns (like "Notes" or "Owner") are left completely untouched.



---

## Deployment & Usage Instructions

### For End-Users (Using the .exe)

1. **Keep Files Together:** Ensure the `FBCRS_Master_Full.xlsx` library file is located in the exact same folder as the `.exe` application.



**Rulebook Requirement:** The app will fail to start if the rulebook is not in the same folder.


**Keep the Excel File Closed:** Please ensure the `FBCRS_Master_Full.xlsx` file is closed before launching the application. If the file is open, the app may be unable to access the data, which can cause the scan to fail.




2. **Launch the App:** Double-click the `.exe` file. A black window will appear—this is the application engine. Keep this window open until the process is complete.


3. **Initiate Scan:** When prompted in the black window, type `Y` to start the analysis or `N` to exit.


4. **Selection:** A window will pop up asking you to select a folder. Choose the directory you wish to scan and click "OK".


5. **Monitoring Progress:** You can follow the progress of the analysis directly in the terminal window.


6. **Completion:** Once the process finishes, the app will display a success message and ask if you wish to run another folder.



### Troubleshooting

* **"Permission Denied" Errors:** Ensure the `_temp.csv` file or the final Excel report is not currently open in another program. Close all related Office applications and try again.



**Performance:** For very large network drives, we recommend dividing your files into smaller sub-folders and scanning them one at a time for the best performance.



**Long Path Errors:** If the scan skips files due to "path length" issues, ensure your system is optimized for long file paths.



---

## Future Versions & Potential Upgrades

We are actively looking to expand the capabilities of this system. Future updates may include:


**Empty Folder Skip:** A feature to automatically identify and skip empty directories to increase scanning speed.


* **Advanced Reporting:** Integration of automated "Summary Charts" within the Excel output to visualize record retention health.
* **Email Integration:** A feature to automatically flag records due for immediate disposition via automated email alerts.

---

## Installation & Setup (Developers)

If you are running the tools directly via Python, ensure you have **Python 3.10 or higher** installed and added to your system PATH.

Install all necessary dependencies via your terminal:

```bash
pip install pandas openpyxl pypdf python-docx python-pptx pyahocorasick

```

To build a standalone executable:

```bash
python -m PyInstaller --onefile "YourScript.py"

```

Note: Ensure the `FBCRS_Master_Full.xlsx` file is kept in the same directory as the generated `.exe`.
