# Excel-File-Merger
This script is designed to combine multiple Excel files located in one directory, add a "Brand" column based on the file name, and combine data from a specific sheet named "Data" in each file. The hyperlinks from the first column (column 'A') are also included in the final output as an additional column. The final results will be saved into a single Excel file with a name that can be customized by the user.
(You can modify this)
How it works!!!

1. User Input:
The user will be asked to enter the date and name of the output file.
The script will add .xlsx automatically if the extension is not included.

2.Excel File Filters:
All .xlsx format files in the working directory will be processed.

3.Column Addition:
Added a "Brand" column taken from the file name.
Added a "Link" column containing hyperlinks from column 'A' in the sheet named "Data".

4.Data Merging:
The script will combine all found Excel files into one DataFrame, replace empty values ​​with 0, and add a "Date" column according to user input.
Output Storage:

The combined files will be saved in one new Excel file according to the file name entered by the user.
