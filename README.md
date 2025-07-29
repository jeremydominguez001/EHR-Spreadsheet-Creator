# EMR Spreadsheet Creator

A Python tool that automatically extracts patient data from PDF files and compiles it into an organized Excel spreadsheet for Electronic Medical Records (EMR) management.

## Overview

This script processes multiple PDF files containing patient EMR information and extracts structured data including patient names, EMR numbers, services, programs, descriptions, status, and dates. The extracted data is then saved into a well-formatted Excel spreadsheet for easy access and management.

## Features

- **Batch PDF Processing**: Automatically processes all PDF files in a specified directory
- **Smart Text Extraction**: Uses regular expressions to accurately extract structured data from PDFs
- **Multi-line Text Handling**: Aggregates text across lines to handle wrapped content
- **Excel Output**: Generates a clean, organized Excel spreadsheet with all extracted data
- **Console Feedback**: Displays extracted records in real-time during processing

## Requirements

- Python 3.6 or higher
- PyMuPDF (fitz)
- pandas
- openpyxl (for Excel support)

## Installation

1. Clone this repository:
```bash
git clone https://github.com/yourusername/EHR-Spreadsheet-Creator.git
cd EHR-Spreadsheet-Creator
```

2. Install required dependencies:
```bash
pip install PyMuPDF pandas openpyxl
```

## Usage

1. Update the folder paths in the script:
   - `folder_path`: Directory containing your PDF files
   - `output_excel`: Path for the output Excel file

2. Run the script:
```bash
python EMR_Spreadsheet_Creator.py
```

3. The script will:
   - Process all PDF files in the specified folder
   - Extract data matching the expected format
   - Display each extracted record in the console
   - Save all data to an Excel file

## Data Format

The script expects PDF content in the following format:
```
LastName, FirstName DMH123456 ACCS ServiceCode VIN ProgramCode Description Text Register MM/DD/YY
```

## Output Format

The generated Excel spreadsheet contains the following columns:
- **Person**: Full name (Last, First)
- **EMR#**: EMR identification number
- **Service**: Service code (ACCS)
- **Program**: Program code (VIN)
- **Description**: Service/visit description
- **Status**: Current status (typically "Register")
- **Date**: Date in MM/DD/YY format

## Configuration

To modify the script for different data formats:

1. Update the regular expression pattern to match your PDF structure
2. Modify the column names in the DataFrame creation
3. Adjust the output Excel filename and location

## Error Handling

- The script will skip any files that are not PDFs
- Text that doesn't match the expected pattern will be ignored
- Console output helps identify which records were successfully extracted

## Contributing

Feel free to submit issues, fork the repository, and create pull requests for any improvements.

## License

This project is open source and available under the [MIT License](LICENSE).

## Disclaimer

This tool is designed for internal EMR data management. Ensure compliance with all relevant healthcare data privacy regulations (HIPAA, etc.) when using this tool with real patient data.
