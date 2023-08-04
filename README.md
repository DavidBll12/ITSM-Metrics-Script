ITSM Metrics Main Script:

This script helps automate the generation of Problem Management metrics using data pulled from ServiceNow. It utilizes two Python libraries, pandas and openpyxl, for data manipulation and interaction with Excel files. The script includes several functions, which are described below:

sheet_created_closed(file_path, dest_sheet, column_name):
This function is responsible for generating 'Created' and 'Closed' Problem Records. It reads data from a ServiceNow export (Excel file), filters records according to the state ('Created' or 'Closed'), and appends the data to a specific sheet in a destination Excel file. The function also applies various formatting, such as adjusting row heights, column widths, and cell alignments.

sheet_open(file_path, dest_sheet, column_name):
This function is designed to generate 'Open' Problem Records. It reads data from a ServiceNow export (Excel file), filters records based on the 'State' column (excluding resolved, known error/closed, or canceled/closed states), and appends the data to a specific sheet in the destination Excel file. Similar to sheet_created_closed, it also handles various formatting tasks and additionally applies certain formulas to cells.

ppm_created(file_path, dest_sheet, column_name, ins_col_num, ins_col_name, formula_letter): 
This function is specifically used for dealing with data related to 'Proactive Problem' records that are either 'created' or 'closed'. It applies additional filtering based on the 'Organization' and 'Problem Type' columns and manipulates certain cell values based on specific conditions. A new column is inserted with a defined header and formula applied. This function, like the others, performs several formatting tasks and saves the modifications to the destination Excel file.

The script concludes with function calls to generate the necessary sheets ('Created', 'Closed', and 'Open' Problem Records, as well as 'Proactive Problem' records) in the final Excel report, which is saved as 'finished_metrics.xlsx'.

By automating the generation of these Problem Management metrics, this script saves significant time and reduces potential errors that could arise from manual reporting. It can be easily adapted for different data and reporting requirements, providing a versatile solution for ITSM reporting tasks.


##############################################################################################################################################################################################################################

ITSM Metrics Check Script:

This script is designed to ensure the consistency and correctness of data across multiple sheets within a single Excel workbook. The workbook in this script, titled 'WeeklyPMMetrics_20230428.xlsx', contains various sheets with data related to Problem Management Metrics.

The script uses the openpyxl Python library to load and interact with Excel data. The seven sheets in the workbook are loaded as individual objects at the start of the script.

The script then conducts several checks on these sheets. Each check compares data from specific cells in different sheets. If the data in the compared cells are not equal, it is an indication that the data are inconsistent, and the script appends an error message to an initially empty list, checks_output. These error messages are specific to the check that failed, allowing for quick identification of the inconsistency.

The checks performed are as follows:

Basic checks on the "Metrics - Weekly" sheet.
Checks on the "Weekly Trend - Overall" sheet.
Cross-sheet checks between the "Weekly Trend - Overall" and "Weekly Trend - AMS/Infra" sheets.
Addition checks on the "Weekly Trend - Overall" sheet.
If all checks pass successfully, the script will append a success message to the checks_output list and then prints all the contents of the list.

This script is instrumental in ensuring the accuracy and consistency of data in the 'WeeklyPMMetrics' Excel workbook. It provides a means of quickly identifying and locating discrepancies in the data, which greatly improves the efficiency of data validation and reduces the risk of decision-making based on incorrect data.

##############################################################################################################################################################################################################################

ITSM Metrics E-mail Script:

This Python script automates the creation of an email describing the weekly ITSM problem management metrics, extracted from an Excel file. The program follows these steps:

Excel File Loading:
The script uses the openpyxl library to load a pre-existing Excel workbook named 'WeeklyPMMetrics_20230428.xlsx'. It extracts data from two worksheets within the workbook, namely 'Metrics - Weekly' and 'Weekly Trend - Overall'.

Metric Extraction:
The program reads and rounds up various metric values from specified cells within the Excel worksheets. These metrics include overall backlog data, EIM backlog data, proactive backlog data, average age backlog data, etc.

Derived Metrics:
The script also calculates derived metrics such as the number and percentage of problems over 30 and 90 days old.

Week-to-Week Comparisons:
Two helper functions compare_to_last_week and compare_to_average are defined and used to make comparisons between the current week's data with the data from the last week and the 9-week average, respectively.

Email Content Generation:
The script then compiles all these metrics and comparisons into a string formatted as the content for an email to leadership. This content provides a week-to-week update on the status, trend, and detailed breakdown of ITSM problem management metrics.

Run the script using a Python interpreter. It will load the data from the specified Excel file and print the formatted email content string to the console.
