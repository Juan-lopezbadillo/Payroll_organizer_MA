import re # Import the 're' module for regular expressions
import os # Import the 'os' module to interact with the operating system
import pandas as pd # Import the 'pandas' library and alias it as 'pd'

def process_and_organize_payroll_by_company(input_file_names, output_excel_filepath):
    """
    Processes multiple payroll text files and organizes employees
    by company into different sheets within a single Excel file.
    """
    # Dictionary to store employees grouped by company name
    # { "Company_Name_1": [[employee1_data], [employee2_data]],
    #   "Company_Name_2": [[employee3_data]], ... }
    all_employees_by_company = {}
    
    # Updated headers for employee sheets
    employee_headers = [
        "SSN",
        "Full_name",
        "Address",
        "City",
        "State",
        "Zip_code",
        "Date_hired",
        "Original_file" # Renamed and maintains its function
    ]

    for input_filepath in input_file_names:
        current_company_name = "Unknown_Company" # Default value if there's an error or 'B' line before 'A'
        print(f"Processing file: {input_filepath}")

        try:
            with open(input_filepath, 'r', encoding='utf-8') as infile:
                for line_num, line in enumerate(infile, 1):
                    line = line.strip()
                    if not line:
                        continue

                    if line.startswith('A'):
                        # Extract company name (same heuristic as before)
                        company_id = line[1:10]
                        remaining_after_id = line[10:]
                        match_address_start = re.search(r'(\d+\s*[A-Za-z]+)', remaining_after_id)
                        
                        if match_address_start:
                            company_name_end_index = match_address_start.start(1)
                            company_name = remaining_after_id[:company_name_end_index].strip()
                        else:
                            company_name = remaining_after_id.strip()
                            
                        # Clean company name of unwanted characters that might cause issues in Excel
                        # For example, / \ ? * [ ] : or very long names
                        cleaned_company_name = re.sub(r'[\\/?*\[\]:]', '', company_name).strip()
                        if not cleaned_company_name: # If it becomes empty after cleaning
                             cleaned_company_name = f"Company_Without_Name_{len(all_employees_by_company) + 1}"

                        # Excel has a 31-character limit for sheet names
                        if len(cleaned_company_name) > 31:
                            cleaned_company_name = cleaned_company_name[:28] + "..." # Truncate and add ellipsis

                        current_company_name = cleaned_company_name

                        # Ensure the company exists as a key in the dictionary
                        if current_company_name not in all_employees_by_company:
                            all_employees_by_company[current_company_name] = []
                        
                        # We don't add 'A' lines directly to the output sheets,
                        # we only use their information to group employees.

                    elif line.startswith('B'):
                        # Person's line (using fixed positions)
                        ssn = line[1:10].strip()
                        full_name = line[10:40].strip()
                        address = line[40:58].strip()
                        city = line[58:76].strip()
                        state = line[76:78].strip()
                        zip_code = line[78:83].strip()
                        date_hired = line[87:93].strip()
                        
                        employee_data = [
                            ssn,
                            full_name,
                            address,
                            city,
                            state,
                            zip_code,
                            date_hired,
                            os.path.basename(input_filepath) # Original_file
                        ]
                        
                        # Ensure the associated company exists before adding the employee
                        if current_company_name not in all_employees_by_company:
                            # This could happen if a file starts with a 'B' line or if the 'A' line is invalid
                            print(f"Warning: Employee '{full_name}' found without associated company on line {line_num} of '{input_filepath}'. Assigning to 'Unknown_Company'.")
                            if "Unknown_Company" not in all_employees_by_company:
                                all_employees_by_company["Unknown_Company"] = []
                            all_employees_by_company["Unknown_Company"].append(employee_data)
                        else:
                            all_employees_by_company[current_company_name].append(employee_data)

        except FileNotFoundError:
            print(f"Warning: The file '{input_filepath}' was not found. Skipping.")
            continue
        except Exception as e:
            print(f"An error occurred while processing line {line_num} in '{input_filepath}': {line}. Error: {e}")
            continue

    # Write data to an Excel file with multiple sheets
    try:
        # We use pd.ExcelWriter to create an Excel workbook with multiple sheets
        with pd.ExcelWriter(output_excel_filepath, engine='openpyxl') as writer:
            for company_name, employees_list in all_employees_by_company.items():
                if employees_list: # Only create sheet if there are employees for that company
                    df = pd.DataFrame(employees_list, columns=employee_headers)
                    # The sheet name will be the company name
                    df.to_excel(writer, sheet_name=company_name, index=False)
                    
                    # --- NEW CODE FOR AUTOFORMATTING ---
                    # Access the openpyxl worksheet object directly from the writer
                    # writer.sheets is a dictionary mapping sheet names to openpyxl worksheet objects
                    worksheet = writer.sheets[company_name] 

                    # Iterate over columns and set width based on content length
                    # This simulates Excel's AutoFit
                    for column_cells in worksheet.columns:
                        max_length = 0
                        column = column_cells[0].column_letter # Get the column letter (e.g., 'A', 'B', 'C')
                        for cell in column_cells:
                            try:
                                if cell.value is not None:
                                    # Convert value to string to measure length, handle non-string types
                                    cell_length = len(str(cell.value))
                                    if cell_length > max_length:
                                        max_length = cell_length
                            except Exception:
                                pass # Catch any errors during length calculation (e.g., complex objects)
                        
                        # Set the width, adding a small buffer (e.g., 2 characters) for readability
                        # A minimum width might also be desired, e.g., max(max_length, min_width)
                        adjusted_width = (max_length + 2) 
                        
                        # Apply width to the column
                        # Note: openpyxl column widths are typically based on font size and scaling,
                        # so '2' might not be a precise 2 characters but a general buffer.
                        worksheet.column_dimensions[column].width = adjusted_width
                    # --- END NEW CODE ---

                else:
                    print(f"Note: Company '{company_name}' has no associated employees to write a sheet.")
        
        print(f"\nAll data processed and organized by company in '{output_excel_filepath}'")
    except Exception as e:
        print(f"Error writing Excel file '{output_excel_filepath}'. Error: {e}")

# --- Script Configuration and Usage ---

# Get the directory path where the script is executed
script_dir = os.path.dirname(os.path.abspath(__file__))

# Find all .txt files in the same directory
input_files_to_process = [f for f in os.listdir(script_dir) if f.endswith('.txt') and os.path.isfile(os.path.join(script_dir, f))]



output_excel_file = 'Payroll_Organized_By_Company.xlsx'

# Execute the main function
process_and_organize_payroll_by_company(input_files_to_process, output_excel_file)