## Program to create the weekly emails based on the Weekly ITSM Problem Management Metric Report
import openpyxl
wb = openpyxl.load_workbook('WeeklyPMMetrics_20230428.xlsx', data_only=True )
sheet1 = wb['Metrics - Weekly']
sheet2 = wb['Weekly Trend - Overall']
# Assigning overall backlog data to variables
overall_current = round(sheet2['E13'].value)
overall_last = round(sheet2['E12'].value)
overall_avg = round(sheet2['E15'].value)

# Assigning EIM backlog data to variables
eim_current = round(sheet2['F80'].value)
eim_last = round(sheet2['F79'].value)
eim_avg = round(sheet2['F82'].value)

# Assigning proactive backlog data to variables
proactive_current = round(sheet2['F130'].value)
proactive_last = round(sheet2['F129'].value)
proactive_avg = round(sheet2['F132'].value)

# Assigning average age backlog data to variables
age_overall_current = round(sheet2['F232'].value)
age_overall_last = round(sheet2['F231'].value)
age_overall_avg = round(sheet2['F234'].value)

# Assigning EIM average age backlog data to variables
age_eim_current = round(sheet2['F273'].value)
age_eim_last = round(sheet2['F272'].value)
age_eim_avg = round(sheet2['F275'].value)

# Assigning Proactive average age backlog data to variables
age_proactive_current = round(sheet2['F292'].value)
age_proactive_last = round(sheet2['F291'].value)
age_proactive_avg = round(sheet2['F294'].value)





# Open - Age Buckets - Tower: Adds together problems over 30 days old assigns result as "num_over_30"
num_cell_range = sheet2['F180':'I180']
num_over_30 = 0
for column in num_cell_range:
    for cell in column:
        num_over_30 += cell.value

# Open - Age Buckets - Tower: Adds together percentages of problems over 30 days old and assigns result as "percent_over_30"
percent_cell_range = sheet2['F181':'I181']
percent_over_30 = 0
for column in percent_cell_range:
    for cell in column:
        percent_over_30 += cell.value      
# Num was rounded, multiplied by 100, and turned into an "int" value
round_percent_over_30 = int(round(percent_over_30, 2) * 100)

# Open Detail - Reactive (EIM/MIM) - Age Buckets - Tower: Adds togethe EIM problems over 30 days old and assigns result as "eim_num_over_30"
eim_num_cell_range = sheet2['G196':'J196']
eim_num_over_30 = 0
for column in eim_num_cell_range:
    for cell in column:
        eim_num_over_30 += cell.value

# Open Detail - Reactive (EIM/MIM) - Age Buckets - Tower: Adds together percentages of problems over 30 days old and assigns result as "eim_percent_over_30"
eim_percent_cell_range = sheet2['G197':'J197']
eim_percent_over_30 = 0
for column in eim_percent_cell_range:
    for cell in column:
        eim_percent_over_30 += cell.value      
eim_round_percent_over_30 = int(round(eim_percent_over_30, 2) * 100)

# Open Detail - Proactive - Age Buckets - Tower: Adds togethe EIM problems over 30 days old and assigns result as "eim_num_over_30"
ppm_num_cell_range = sheet2['G211':'J211']
ppm_num_over_30 = 0
for column in ppm_num_cell_range:
    for cell in column:
        ppm_num_over_30 += cell.value

# Open Detail - Proactive - Age Buckets - Tower: Adds together percentages of problems over 30 days old and assigns result as "eim_percent_over_30"
ppm_percent_cell_range = sheet2['G212':'J212']
ppm_percent_over_30 = 0
for column in ppm_percent_cell_range:
    for cell in column:
        ppm_percent_over_30 += cell.value      
ppm_round_percent_over_30 = int(round(ppm_percent_over_30, 2) * 100)


# Open - Age Buckets - Tower: gets num and percentage of problems over 90 days old
num_over_90 = sheet2['I180'].value
percent_over_90 = int(round(sheet2['I181'].value, 2) * 100)

# Open Detail - Reactive (EIM/MIM) - Age Buckets - Tower: gets num and percentage of problems over 90 days old
eim_num_over_90 = sheet2['J196'].value
eim_percent_over_90 = int(round(sheet2['J197'].value, 2) * 100)

# Open Detail - Proactive - Age Buckets - Tower: gets num and percentage of problems over 90 days old
ppm_num_over_90 = sheet2['J211'].value
ppm_percent_over_90 = int(round(sheet2['J212'].value, 2) * 100)

# Compare a data point from this week to last week
def compare_to_last_week(current_week, previous_week):
    if current_week > previous_week:
        return "up from last week"
    elif current_week == previous_week:
        return "same as last week"
    elif current_week < previous_week:
        return "down from last week"

# Compare a data point from this week to 9 week average
def compare_to_average(current_week, average):
    if current_week > average:
        return "above 9 week average"
    elif current_week == average:
        return "same as 9 week average"
    elif current_week < average:
        return "below 9 week average"
    
    # Runs functino to compare data to last week and returns the result as a string
overall_last_week = compare_to_last_week(overall_current, overall_last)
eim_last_week = compare_to_last_week(eim_current, eim_last)
proactive_last_week = compare_to_last_week(proactive_current, proactive_last)

# Runs function to compare data to 9 week average and returns the result as a string
overall_average = compare_to_average(overall_current, overall_avg)
eim_average = compare_to_average(eim_current, eim_avg)
proactive_average = compare_to_average(proactive_current, proactive_avg)

# AVG AGE: Runs function to compare data to last week and returns the result as a string
age_overall_last_week = compare_to_last_week(age_overall_current, age_overall_last)
age_eim_last_week = compare_to_last_week(age_eim_current, age_eim_last)
age_proactive_last_week = compare_to_last_week(age_proactive_current, age_proactive_last)

# AVG AGE: Runs function to compare data to 9 week average and returns the result as a string
age_overall_average = compare_to_average(age_overall_current, age_overall_avg)
age_eim_average = compare_to_average(age_eim_current, age_eim_avg)
age_proactive_average = compare_to_average(age_proactive_current, age_proactive_avg)

print(f"""Week ending overall problem backlog is ({overall_current}) problems; {overall_last_week} ({overall_last}), {overall_average} ({overall_avg}), below 6 month average (40).
EIM Reactive Problems: Status: Neutral; Trend: Negative 
Week ending EIM Reactive Problem backlog is ({eim_current}) problems; {eim_last_week} ({eim_last}), {eim_average} ({eim_avg}), below 6 month trending average (18).
Proactive Problems: Status: Neutral; Trend: Negative - [objective is to increase # of proactive problems]
Week ending Proactive Problem backlog is ({proactive_current}) problems; {proactive_last_week} ({proactive_last}), {proactive_average} ({proactive_avg}), below 6 month average 18.""")

print(f"""All Problems
Avg Age (Days) is {age_overall_current} days; {age_overall_last_week} ({age_overall_last}), {age_overall_average} ({age_overall_avg}), above 6 month average (63). Target for overall problem resolution is > 30 days. 
{num_over_30} ({round_percent_over_30}%) problems are > 30 days old,  {num_over_90} ({percent_over_90}%) are > 90 days old. 
EIM Reactive Problems: Status: Neutral; Trend: Negative 
Avg Age (Days) is {age_eim_current} days; {age_eim_last_week} ({age_eim_last}), {age_eim_average} ({age_eim_avg}), above 6 month average (55). Target for problem resolution is > 30 days.
{eim_num_over_30} ({eim_round_percent_over_30}%) problems are > 30 days old, {eim_num_over_90} ({eim_percent_over_90}%) are > 90 days old.
Proactive Problems: Status: Neutral; Trend: Positive
Avg Age (Days) is {age_proactive_current} days; {age_proactive_last_week} ({age_proactive_last}), {age_proactive_average} ({age_proactive_avg}), above 6 month average (76). Target for problem resolution is > 30 days, but is expected proactive problems will take longer to resolve than reactive problems, as are lower priority.
{ppm_num_over_30} ({ppm_round_percent_over_30}%) problems are > 30 days old, {ppm_num_over_90} ({ppm_percent_over_90}%) are > 90 days old.""")
