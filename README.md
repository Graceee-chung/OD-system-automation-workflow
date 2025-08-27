# OD-system-automation-workflow
This project automates downloading, naming, and archiving official documents from the Vital OD system using Selenium + GUI automation (PyAutoGUI). The pipeline logs in, iterates pending items, opens each document, saves a PDF with a standardized filename, downloads attachments, and organizes everything into per-document folders with robust logging and retries.

### Features
1. Headless-optional Selenium login & navigation to “待辦理”.
2. Auto open each document tab and trigger Print → Save as PDF.
3. GUI automation fallback for Windows print dialogs.
4. Deterministic file/folder names from parsed metadata (doc no., sender, date).
5. Attachment downloading and structured archiving.
6. Resilient retries, window focus handling, and detailed logs.

### Requirements
Python 3.10+
Google Chrome + matching ChromeDriver
Windows (for GUI print dialogs)
Packages: selenium, pyautogui, pygetwindow, pyperclip, python-dotenv (if used), etc.

### File-by-file summary
- **main.py** – Orchestrates the whole workflow: initializes logging/driver, iterates documents, calls handlers to save PDFs and attachments, and controls retries.  
- **config.py** – Centralized settings (credentials, URLs, timeouts, paths, feature toggles) to adapt the automation without touching code.  
- **driver_util.py** – Builds and configures the Selenium Chrome WebDriver (options, user-agent, waits) and provides helper functions for window/tab management.  
- **logger_setup.py** – Sets up structured logging (console/file formatters, levels, emoji hints) used across all modules.  
- **document_util.py** – Parses HTML to extract metadata (doc number, sender, dates), normalizes strings, and generates canonical file/folder names.  
- **document_handler.py** – High-level document routines: open a row, switch tabs, invoke print/save, download attachments, and write results into the archive.  
- **pdf_handler.py** – Thin façade choosing the appropriate PDF-saving strategy and exposing `trigger_print` / `save_pdf_gui` wrappers.  
- **pdf_handler_v1.py** – Baseline GUI automation for “Print → Save as PDF” using fixed coordinates and image checks; includes a retry flow that reopens failed rows.  
- **pdf_handler_v2.py** – More robust GUI automation using template-matching (image recognition), Chrome window focusing, safer dialogs, and improved error screenshots.  

### Output Layout
archive_root/
└─ {doc_no} {sender} {ROC_date} {issue_no}/
   ├─ {same_string}.pdf
   ├─ attachments/
   │   └─ (all downloaded files)
   └─ metadata.log

