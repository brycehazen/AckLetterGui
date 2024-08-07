# Acknowledgment Letter and Mail Merge Tool

## Table of Contents
- [Description](#description)
- [File Descriptions](#file-descriptions)
- [Requirements](#requirements)
- [Usage](#usage)
- [File Descriptions](#file-descriptions)

## Description
The Acknowledgment Letter and Mail Merge Tool is designed to automate the process of generating acknowledgment letters for donors. It processes CSV files exported from the Mail module, ensures the correctness of titles and genders, and performs a mail merge using a DOCX template.

## File Descriptions

### CSV Files
- **_mail.csv**: Directly exported from the Mail module.
- **_export.csv**: Comes from a query created by the Mail Module, then use the Export module using the labels export as a gift export.
- **DOCX template**: A DOCX template needed for mail merge.

### Python Scripts
- **ack_letter.py**: Processes the CSV files to generate the acknowledgment letters data.
- **mail_merge.py**: Performs the mail merge operation using the processed data and the DOCX template.
- **labels.py**: Ensures the correctness of titles and genders in the exported CSV files.
- **ack_mail_merge_gui.py**: Provides a GUI for the entire process, allowing users to select input/output folders, run the label processor, run the acknowledgment letter processor, select a DOCX template for mail merge, and perform the mail merge.

## Requirements
- Python 3.7+
- `pandas`
- `tabulate`
- `PySide6`
- `python-docx`

## Usage
1. **Set up your environment**:
   ```bash
   python -m venv env
   source env/bin/activate  # On Windows, use `env\\Scripts\\activate`
   pip install -r requirements.txt
