# %%
import os
# import tarfile
import pandas as pd
import datetime


# %%
# Date Formate
date = datetime.datetime.now().strftime("%d-%m-%Y")

# File locations 
# ucm_cdr_tar_file = '~/Documents/logs/ucm_pri/Master_condition_ucmadmin.tar.gz'
pri_cdr_csv_file = '~/Documents/logs/ucm_pri/Master_condition_admin.csv'
# tar_destination = '~/Documents/logs/ucm_pri/'
check_for_UCM = '~/Documents/logs/ucm_pri/Master_condition_ucmadmin.csv'
new_UCM_report_name = '~/Documents/logs/ucm_pri/UCM_CDR_' + date + '.csv'
new_PRI_report_name = '~/Documents/logs/ucm_pri/PRI_CDR_' + date + '.csv'
new_UCM_tar_name = '~/Documents/logs/ucm_pri/UCM_CDR_' + date + '.tar.gz'
report_location = '~/Documents/logs/ucm_pri/UCM_CDR_Report_' + date + '.xlsx'


# # %%
# # Open tar.gz file
# ucm_cdr = tarfile.open(ucm_cdr_tar_file)
# # Extract the tar.gz file in specified location
# ucm_cdr.extractall(tar_destination)
# # close tar.gz file
# ucm_cdr.close()

# %%
# check if file exist, if true rename the file

if os.path.isfile(pri_cdr_csv_file):
    os.rename(pri_cdr_csv_file, new_PRI_report_name)        
else:
    print(f"{pri_cdr_csv_file} not found")
    exit

if os.path.isfile(check_for_UCM):
    os.rename(check_for_UCM, new_UCM_report_name)
    # os.rename(ucm_cdr_tar_file, new_UCM_tar_name)
else:
    print(f"{check_for_UCM} not found")
    exit

# %%
# Read UCM and PRI CDR 
df_ucm = pd.read_csv(new_UCM_report_name)
df_pri = pd.read_csv(new_PRI_report_name)
# %%

# Split Start Time into Start Date and Start Time
df_ucm[['Start_Date', 'Start_Time']] = df_ucm['Start Time'].str.split(' ', expand=True)
df_pri[['Start_Date', 'Start_Time']] = df_pri['start time'].str.split(' ', expand=True)

# %%
# Create pivot table for UCM and PRI including all columns
df_ucm_pivot = pd.pivot_table(df_ucm, values="Talk Time", index=["Start_Date", "Caller Number", "Answered by", "Call Type"], aggfunc="sum")
df_pri_pivot = pd.pivot_table(df_pri, values="talk time", index=["Start_Date", "caller number", "answer by"], aggfunc="sum")

# Calculate Time in hours
df_ucm_pivot["Time"] = df_ucm_pivot["Talk Time"] / 3600  # Convert seconds to hours
df_pri_pivot["Time"] = df_pri_pivot["talk time"] / 3600   # Convert seconds to hours

# Convert seconds to timedelta
df_ucm_pivot["Time"] = pd.to_timedelta(df_ucm_pivot["Time"], unit='h')
df_pri_pivot["Time"] = pd.to_timedelta(df_pri_pivot["Time"], unit='h')

# # Format timedelta to hh:mm:ss as strings
df_ucm_pivot["Time"] = df_ucm_pivot["Time"].apply(lambda x: '{:02}:{:02}:{:02}'.format(x.components.hours, x.components.minutes, x.components.seconds))
df_pri_pivot["Time"] = df_pri_pivot["Time"].apply(lambda x: '{:02}:{:02}:{:02}'.format(x.components.hours, x.components.minutes, x.components.seconds))



with pd.ExcelWriter(report_location) as writer:
    df_ucm.to_excel(writer, sheet_name='UCM_CDR_' + date, index=False)
    
    # Save the second sheet after a certain row on the first sheet
    df_pri.to_excel(writer, sheet_name='PRI_CDR_' + date, index=False)
    df_ucm_pivot.to_excel(writer, sheet_name='UCM_Pivot_' + date)
    df_pri_pivot.to_excel(writer, sheet_name='PRI_Pivot_' + date)
