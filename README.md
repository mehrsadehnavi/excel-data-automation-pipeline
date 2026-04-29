# Excel Data Automation Pipeline

## Overview
This project presents a Python-based data processing pipeline designed to automate the migration, transformation, and analysis of Excel-based instructor records.
The system handles large-scale datasets (500+ files) with heterogeneous structures and replaces manual workflows with automated, reproducible processes.

## Key Features
- Automated generation of Excel files from a template
- Migration of legacy data with schema differences
- Row-to-row and column-to-column data transformation
- Bulk addition and modification of worksheets
- Detection of missing or outdated records
- Computation of instructor performance metrics
- Batch file renaming and restructuring
- Removal of file-level restrictions (e.g., passwords)

## Project Structure
excel-data-automation-pipeline/
│
├── README.md
│
├── scripts/
│   ├── 01_generate_files.py
│   ├── 02_migrate_legacy_data.py
│   ├── 03_row_to_row_mapping.py
│   ├── 04_add_sheets.py
│   ├── 05_rename_sheet.py
│   ├── 06_remove_password.py
│   ├── 07_detect_missing_observations.py
│   └── 08_compute_average_scores.py
│
├── sample_data/
|   ├── old_format_sample.xlsx
│   └── new_template_sample.xlsx
│
└── outputs/
│   ├── missing_observations.txt
│   └── average_scores.txt

## Workflow
1. Generate standardized Excel files from template
2. Migrate legacy data into new structure
3. Apply row-level transformations
4. Update workbook structure (add/remove sheets)
5. Perform data validation and rule-based analysis
6. Generate reports (missing records, averages)

## Technologies
- Python
- pandas
- openpyxl
- os / pathlib

## Author
Mehrsa Dehnavi
