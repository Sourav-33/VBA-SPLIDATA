#Automated Data Splitting and Filtering with VBA

Overview
This project features a VBA macro designed to automate the process of splitting data based on unique values in a specified column and filtering this data into separate worksheets. The solution leverages Excel's built-in functions for efficient data processing and includes a user-friendly interface using Form Controls.

Features
1. Automated Data Splitting: Dynamically identify unique values within a specified range and create separate worksheets for each unique value.
2. Filtering and Copying Data: Filter data based on unique values and copy the filtered data into the respective worksheets.
3. User-Friendly Interface: Includes a Form Control button to trigger the data splitting process.
4. Error Handling: Ensures smooth execution and alerts users of any issues during the process.
Technical Details
5. Language: VBA (Visual Basic for Applications)
6. Environment: Microsoft Excel
7. Functions Used: WorksheetFunction.Match, WorksheetFunction.CountA, WorksheetFunction.Unique

Usage
Setup:
-> Open the workbook VBA Code Split Data.xlsm.
-> Ensure that the data to be split is in Sheet1.

Using the Form Control Button:
-> A Form Control button is included in the worksheet to trigger the split_data2 macro.
-> Click the button to execute the data splitting process.

Execution:
-> The macro will create new worksheets named after each unique value found in the specified column.
-> Filtered data for each unique value will be copied into the corresponding worksheet.
-> A temporary worksheet unique_values is created and deleted after the process is completed.
