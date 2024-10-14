# WBS Data Processing Macro

This project features a macro designed to streamline the projects (WBS) data. It automates the process of filtering, copying, and transforming data between worksheets. The macro primarily focuses on filtering specific entries with statuses **REL** (Released) and **TECO** (Technically Completed), then processes the data by modifying and clearing certain columns based on the criteria. This helps in automating and speeding up the monthly or weekly reporting process.

## Macro Overview

The macro performs the following key tasks:

1. **Data Preparation and AutoFill**:
    - Activates the **WBS Raw** sheet and determines the last row containing data.
    - Switches to the **WBS Working** sheet and applies an autofill to replicate the data format from **Row 3** down to the last row.

2. **Data Filtering**:
    - Applies an auto filter on the **WBS Working** sheet, specifically filtering for entries with status **REL** (Released) or **TECO** (Technically Completed).
    - Copies the filtered data to the **WBS Data** sheet, pasting only the values without any formatting.

3. **Data Cleanup**:
    - Clears all the existing data in the **WBS Data** sheet before pasting the new filtered values.
    - Ensures that after pasting, the autofilter is turned off in the **WBS Working** sheet for further operations.

4. **TECO Entry Handling**:
    - Loops through the **WBS Data** sheet, identifying rows where the status is **TECO**.
    - For each **TECO** entry, the macro sets the values in columns **Y**, **Z**, **AA**, and **AB** to 0, as part of the data correction process.

5. **Refreshing and Final Steps**:
    - Once the processing is completed, the macro triggers a full workbook refresh to ensure all data connections and calculations are up-to-date.
    - Displays a "Done" message box to inform the user that the operation has been successfully completed.

## Key Features

- **Automated Data Filtering**: Automatically applies filters for key WBS statuses (**REL** and **TECO**) and copies the filtered results to the designated worksheet.
- **TECO Handling**: Modifies specific columns for **TECO** entries, setting them to 0 for further processing.
- **Data Cleanup**: Ensures the target sheet is cleaned and updated with the latest data without requiring manual intervention.
- **Workbook Refresh**: Refreshes the entire workbook after processing to ensure all formulas and data links are up-to-date.
