📊 Production Reporting Automation
🔎 Overview

This project provides an automated workflow for preparing daily production reports.
It integrates data from Excel files selected by the user into a single reporting workbook, ensuring accuracy, consistency, and reduced manual effort.

The solution is tailored for injection molding (brizganje) production planning and reporting for 2025.

📝 Workflow Summary

Select and Import Data

The user selects a source data file (Excel).

All data from the first sheet is copied and pasted into the reporting workbook → “Prilepi Gosoft” sheet (values only).

Load Production Plan

The user selects a monthly production plan file (Excel).

The program identifies the correct date group (yesterday’s date, or Friday if today is Monday).

A range (rows 6–44, 3 columns wide) is copied.

Insert into Reporting Workbook

The selected plan data is pasted into “brizganje izracun” sheet in the reporting workbook, aligned with the correct date.

Row Analysis

Rows 7–46 are scanned:

If column A contains a code and column M > 50 €, the code (e.g., B001) is recorded.

Per-Code Processing

For each recorded code:

Filter the Izbor sheet by that code.

Clear previous data in List2.

Copy filtered values (columns F–M) into List2 → T1.

Run macro gumb1 to process the data.

Copy processed results (B1:L8).

Paste results back into brizganje izracun, below the corresponding row (as an image).

📂 Repository Contents
.
├── automate_process.py        # Main script to automate data import and report updates
├── porocilo brizganje.py      # GUI script
├── REQUIREMENTS.txt           # Python dependencies
├── install_requirements.bat   # Helper script to install dependencies
└── README.md                  # Project documentation

⚙️ Requirements

Microsoft Excel (with macros enabled).

User-provided Excel files:

Source data file (daily production data).

Production plan file (monthly planning).

Reporting workbook (template with macros).

🚀 Usage Instructions

Run the program.

When prompted, select the required Excel files:

Source data file.

Production plan file.

Reporting workbook.

The program will automatically:

Import source data.

Insert the production plan for the correct date.

Process and analyze rows.

Generate per-code reports.

Review the updated reporting workbook.

🎯 Purpose

This project was developed to:

Simplify and standardize daily production reporting.

Reduce manual Excel operations and potential errors.

Allow flexibility by letting users select input files at runtime.
