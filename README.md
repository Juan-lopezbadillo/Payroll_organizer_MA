# Payroll_organizer_MA
Py script that create an excel file base on .txt documents and the format to add a New Hire record for MA-USA https://www.mass.gov/info-details/new-hire-record-layout 

# Payroll Organizer MA

This repository contains a Python script designed to automate the organization of new hire data for the state of Massachusetts (MA-USA) into an easy-to-manage Excel spreadsheet format.

## Purpose

The primary goal of this script is to simplify the new hire record management process by converting structured data from plain text files into a tabular Excel-compatible format, ready for further manipulation or integration into payroll systems.

## General Functionality

The script operates as follows:

1.  **Data Source:** The script searches for and processes text files (`.txt`) which must be located **within the same directory** as the Python script itself.

2.  **Processing:** The script reads each `.txt` file, extracting company information and employee details based on a predefined fixed-width (or positional) format.

3.  **Data Output:** Once the `.txt` files are successfully processed, the extracted data is organized and exported to an **Excel file (.xlsx)**. This Excel file is structured to facilitate easy viewing, filtering, and any subsequent manipulation by the user, optimizing the management of new hire records.

## Input File Format (`.txt`)

For the script to process data correctly, each text file (`.txt`) must adhere to a **fixed-width or positional character structure**. The script is designed to read each line and classify the information based on a starting prefix:

* **Lines starting with `A`:** Correspond to **Company** information.
* **Lines starting with `B`:** Correspond to **Employee** information.

### Detailed Structure (Based on Provided Example):

Each `.txt` file must contain at least one company line followed by one or more associated employee records. Fields are distinguished by their position within the line and by the starting prefix.

1.  **Company Information Line (Prefix `A`):**
    * [cite_start]**Prefix:** Always starts with `A`[cite: 1].
    * **Field 1 (Company ID/Payroll Number):** 9 digits immediately following the 'A'. [cite_start]Example: `462915012`[cite: 1].
    * **Field 2 (Company Name):** A string of characters following the ID, ending before the address. [cite_start]Example: `Mass Alternative Care Inc`[cite: 1].
    * **Field 3 (Company Address):** Continues after the company name. [cite_start]Example: `1247 East Main St`[cite: 1].
    * **Field 4 (Company City):** Continues after the address. [cite_start]Example: `Chicopee`[cite: 1].
    * **Field 5 (Company State):** 2 characters (state abbreviation). [cite_start]Example: `MA`[cite: 1].
    * **Field 6 (Company Zip Code):** 5 digits. [cite_start]Example: `01020`[cite: 1].
    * **Field 7 (Control Number/Internal ID):** A final 8-digit number. [cite_start]Example: `25000001`[cite: 1].
    * **Full `A` Line Example:**
        ```
        A462915012Mass Alternative Care Inc     1247 East Main St Chicopee          MA01020    25000001 
        ```

2.  **Employee Information Line (Prefix `B`):**
    * [cite_start]**Prefix:** Always starts with `B`[cite: 1].
    * **Field 1 (SSN):** 9 digits without hyphens, immediately following the 'B'. [cite_start]Example: `596325420`[cite: 1].
    * **Field 2 (Full_name):** Last name(s) and First name(s) separated by a comma. [cite_start]Example: `Rodriguez Velez,Lireydaliz`[cite: 1].
    * **Field 3 (Address):** Employee's address. [cite_start]Example: `197 North Valley R` (note potential truncation, e.g., "Road")[cite: 1].
    * **Field 4 (City):** Employee's city. [cite_start]Example: `Pelham`[cite: 1].
    * **Field 5 (State):** 2 characters (state abbreviation). [cite_start]Example: `MA`[cite: 1].
    * **Field 6 (Zip_code):** 5 digits. [cite_start]Example: `01002`[cite: 1].
    * **Field 7 (Date_hired):** Hire date, in `MMDDYY` format (MonthDayYear two-digit). [cite_start]Example: `041025` (corresponding to April 10, 2025)[cite: 1].
    * **Full `B` Line Example:**
        ```
        B596325420Rodriguez Velez,Lireydaliz    197 North Valley RPelham            MA01002    041025
        ```
    * **Excel Headers:** The script will use the following headers for the employee data columns in the final Excel file: `SSN`, `Full_name`, `Address`, `City`, `State`, `Zip_code`, `Date_hired`.

**Important Considerations:**

* **Fixed-Width/Positional Consistency:** It is vital that the `.txt` files maintain this exact format consistently in terms of field position and length. Intermediate spaces (e.g., between company name and address) are likely significant for the script.
* **Implicit Delimiters:** The script must be designed to parse these lines based on positions or a pattern of spaces, not explicit field delimiters (except for the comma within `Full_name`).
* **Errors:** Any deviation in the format (incorrect prefixes, missing fields, altered order, incorrect spacing) could cause processing errors in the script.

---

## Requirements

* Python 3.x
* **(Please list any external Python libraries your script uses here, e.g., `pip install pandas openpyxl`)**

## How to Use

1.  Ensure your Python script (`**Payroll_organizar_MA.py**`) and the input `.txt` files are in the same directory.
2.  **(If executable from the terminal, add the command line instruction here, e.g.: `Payroll_organizar_MA.py`)**
3.  The generated Excel file will appear in the same directory where the script was executed.

---
