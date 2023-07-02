import openpyxl
import re

print('Running checks...\n')
wb = openpyxl.load_workbook('WeeklyPMMetrics_20230428.xlsx', data_only=True)

sheet1 = wb['Metrics - Weekly']
sheet2 = wb['Weekly Trend - Overall']
sheet3 = wb['Weekly Trend - AMS']
sheet4 = wb['Weekly Trend - Infra']
sheet5 = wb['EPIs']
sheet6 = wb['EPI Trending (Wkly) - Open ']
sheet7 = wb['EPI Trending (Wkly) - Closed']

# Creates an empty list that will hold failed checks
checks_output = []

# Checks on "Metrics - Weekly" sheet
if sheet1['B4'].value != sheet2['C13'].value:
    checks_output.append("Weekly Created check failed.")
if sheet1['B5'].value != sheet2['D13'].value:
    checks_output.append("Weekly Closed check failed.")
if sheet1['B6'].value != sheet2['E13'].value:
    checks_output.append("Weekly Open check failed.")

# Checks on "Weekly Trend - Overall" sheet
if sheet2['E13'].value != sheet2['F30'].value:
    checks_output.append("Open Total check failed.")
if sheet2['E47'].value != sheet2['F63'].value:
    checks_output.append("Open Detail Reactive check failed.")
if sheet2['F80'].value != sheet2['G96'].value:
    checks_output.append("Open Detail Reactive EIM check failed.")
if sheet2['F130'].value != sheet2['G147'].value:
    checks_output.append("Open Detail Proactive check failed.")
if sheet2['J165'].value != sheet2['J180'].value:
    checks_output.append("Open Age Buckets check failed.")
if sheet2['F232'].value != sheet2['F254'].value:
    checks_output.append("Open Average Age (Days) check failed.")

# Checks between "Weekly Trend - Overall" and "Weekly Trend - AMS/Infra" sheets
if sheet2['C30'].value != sheet3['J12'].value:
    checks_output.append("Open Total - AMS check failed.")
if sheet2['D30'].value != sheet4['M12'].value:
    checks_output.append("Open Total - Infra check failed.")
if sheet2['D96'].value != sheet3['K29'].value:
    checks_output.append("Reactive EIM - AMS check failed.")
if sheet2['E96'].value != sheet4['N29'].value:
    checks_output.append("Detail - Reactive EIM - Infra check failed.")
if sheet2['D147'].value != sheet3['K46'].value:
    checks_output.append("Detail - Proactive - AMS check failed.")
if sheet2['E147'].value != sheet4['N46'].value:
    checks_output.append("Detail - Proactive - Infra check failed.")
if sheet2['J177'].value != sheet3['J62'].value:
    checks_output.append("Age Buckets - AMS check failed.")
if sheet2['J178'].value != sheet4['J64'].value:
    checks_output.append("Age Buckets - Infra check failed.")
if sheet2['K193'].value != sheet3['J80'].value:
    checks_output.append("Reactive EIM - Age Buckets - AMS check failed.")
if sheet2['K194'].value != sheet4['J85'].value:
    checks_output.append("Reactive EIM - Age Buckets - Infra check failed.")
if sheet2['K208'].value != sheet3['J98'].value:
    checks_output.append("Proactive - Age Buckets - AMS check failed")
if sheet2['K209'].value != sheet4['J106'].value:
    checks_output.append("Proactive - Age Buckets - Infra check failed")
if sheet5['D11'].value != sheet6['K45'].value:
    checks_output.append("Open Total - AMS check failed")
if sheet5['F11'].value != sheet7['K45'].value:
    checks_output.append("Open Total - Infra check failed")

# Addtion Checks for "Weekly Trend - Overall"
if ((sheet2['E12'].value + sheet2['C13'].value) - sheet2['D13'].value) != sheet2['E13'].value :
    checks_output.append("Open Total - Addition check failed")    
if ((sheet2['E46'].value + sheet2['C47'].value) - sheet2['D47'].value) != sheet2['E47'].value :
    checks_output.append("Open Detail - Reactive - Addition check failed")    
if ((sheet2['F79'].value + sheet2['D80'].value) - sheet2['E80'].value) != sheet2['F80'].value :
    checks_output.append("Open Detail - Reactive EIM/MIM - Addition check failed")    
if ((sheet2['F129'].value + sheet2['D130'].value) - sheet2['E130'].value) != sheet2['F130'].value :
    checks_output.append("Open Detail - Proactive - Addition check failed")    

    
else:
    checks_output.append("Success: All Checks Passed!")
    
print(checks_output)