import os
import fitz  # PyMuPDF
import pandas as pd
import re

# Folder path where your PDF files are stored
folder_path = r"\python\Master List and IDs\EMRs"
output_excel = r"\python\Master List and IDs\EMRs\EMR_Spreadsheet_Results.xlsx"

# Regular expression pattern to match the data format
pattern = re.compile(
    r"(?P<last_name>[^,]+),\s*(?P<first_name>[^\s]+)\s+(?P<dmh>\w+)\s+(?P<accs>ACCS\s+\w+)\s+(?P<vin>VIN\s+\w+)\s+(?P<description>.+?)\s+Register\s+(?P<date>\d{2}/\d{2}/\d{2})"
)

# List to store extracted data
data = []

# Loop through each PDF in the folder
for filename in os.listdir(folder_path):
    if filename.endswith(".pdf"):
        # Open PDF file
        with fitz.open(os.path.join(folder_path, filename)) as pdf_file:
            aggregated_text = ""
            
            # Extract text from each page
            for page_num in range(pdf_file.page_count):
                page = pdf_file[page_num]
                page_text = page.get_text("text")
                
                # Aggregate lines
                for line in page_text.splitlines():
                    # Append to aggregated text
                    aggregated_text += line + " "
                    
                    # Attempt to match the pattern on the aggregated text
                    match = pattern.search(aggregated_text)
                    if match:
                        # If a match is found, extract data and reset the aggregated text
                        row = {
                            "Person": f"{match.group('last_name')}, {match.group('first_name')}",
                            "EMR#": match.group('dmh'),
                            "Service": match.group('accs'),
                            "Program": match.group('vin'),
                            "Description": match.group('description').strip(),
                            "Status": "Register",
                            "Date": match.group('date')
                        }
                        print(row)
                        data.append(row)
                        aggregated_text = ""  # Reset aggregated text after a successful match

# Create DataFrame from the extracted data
df = pd.DataFrame(data, columns=["Person", "EMR#", "Service", "Program", "Description", "Status", "Date"])

# Save the DataFrame to Excel
df.to_excel(output_excel, index=False)
print(f"Data extracted and saved to {output_excel}")
