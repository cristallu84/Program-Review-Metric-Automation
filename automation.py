import numpy as np 
import pandas as pd 
from openpyxl import Workbook, load_workbook 

excel_file = 'Mavenlink Export Current.xlsx'
extraction_file = 'Live Metric Extraction.xlsx'
allo_excel_file = 'Allocated vs Actual Hours.xlsx'
itd_file = 'Allocated vs Actual Hours (ITD).xlsx'
burnup_file = 'Burn up Chart Data.xlsx'
date = pd.Timestamp.today().strftime('%Y-%m-%d') #Adding date that code is run 

#---------------------------Get EAC metrics 

df = pd.read_excel(excel_file, sheet_name=0, usecols=['Project: Name','Task: Top Level', 'Actual Hrs', "Budgeted Hrs", "Remaining Hrs"]) #Create dataframe for actual, budgeted and remaining hrs 
projects = df[['Project: Name','Actual Hrs', "Budgeted Hrs", "Remaining Hrs"]].where(df['Task: Top Level'] == 'Rollup').dropna() #Filter for Rollup, N/A, 


activeproj = pd.read_excel(extraction_file, sheet_name="Active Projects")
activeproj['Total Hours Used (from start of program until today)'] = ""
activeproj['Original EAC (in hours)'] = ""
activeproj['Estimated Hours to Complete'] = ""

#For every project in the projects dataframe, it will loop through the allo_projects dataframe to find the same project and add its actual & allocated hrs to its respective row in projects dataframe
for j, row in activeproj.iterrows():
    for i, row in projects.iterrows(): 
        if activeproj['Active Projects'][j][:10] == projects['Project: Name'][i][:10]: #Finds the row for the project in allo_projects 
            activeproj['Total Hours Used (from start of program until today)'][j] = projects['Actual Hrs'][i]
            activeproj['Original EAC (in hours)'][j] = projects['Budgeted Hrs'][i]
            activeproj['Estimated Hours to Complete'][j] = projects['Remaining Hrs'][i]
#print(activeproj) #- this is for checking if outputs are correct (leave as comment)

#-----------------------------Populate Erik's Burn Up Chart Data Sheet 
burnup_df = pd.read_excel(burnup_file, sheet_name=0)
burnup_df['Total Hours Used (from start of program until today) ' + date] = ""
burnup_df['Estimated Hours to Complete ' + date] = ""
burnup_df['Revised EAC ' + date] = ""

for a, row in burnup_df.iterrows():
    for b, row in projects.iterrows(): 
        if burnup_df['Project Name'][a][:10] == projects['Project: Name'][b][:10]: #Finds the row for the project in allo_projects 
            burnup_df['Total Hours Used (from start of program until today) ' + date][a] = projects['Actual Hrs'][b]
            burnup_df['Estimated Hours to Complete ' + date][a] = projects['Remaining Hrs'][b]
            burnup_df['Revised EAC ' + date][a] = projects['Actual Hrs'][b] + projects['Remaining Hrs'][b]

            
with pd.ExcelWriter(burnup_file, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer: 
  burnup_df.to_excel(writer, index=False, sheet_name='Sheet 1') #Exports projects dataframe to Burn up Chart Data

#-----------------------------Get Allocated and Actual Hours 
allodf = pd.read_excel(allo_excel_file, sheet_name=0) #Create dataframe for allocated vs actual hours 

allo_projects = allodf.where((allodf.iloc[:,1] == 'Rollup')) #Filter for Rollup in second column only 


allo_projects.rename(columns={allo_projects.columns[0]: 'Project: Name',allo_projects.columns[1]: 'Task: Top Level', allo_projects.columns[2]: 'Allocated Hrs', allo_projects.columns[3]: 'Actual Hrs'}, inplace= True)#Renaming Columns 
allo_projects.dropna(subset = ['Project: Name'], inplace=True) #Filtering out any N/A 

#print(allo_projects) #- this is for checking if outputs are correct (leave as comment)

activeproj["Allocated Hrs"] = ""
activeproj["Actual Hrs"] = "" 


#For every project in the projects dataframe, it will loop through the allo_projects dataframe to find the same project and add its actual & allocated hrs to its respective row in projects dataframe
for j, row in activeproj.iterrows():
    for i, row in allo_projects.iterrows(): 
        if allo_projects['Project: Name'][i][:10] == activeproj['Active Projects'][j][:10]: #Finds the row for the project in allo_projects 
            activeproj['Actual Hrs'][j] = allo_projects['Actual Hrs'][i] #Inputs numbers from allo_projects to projects 
            activeproj['Allocated Hrs'][j] = allo_projects['Allocated Hrs'][i]
#print(activeproj) #- this is for checking if outputs are correct (leave as comment)

#----------------------------Get allocated ITD
itd_df = pd.read_excel(itd_file,sheet_name=0 )#Create dataframe for allocated vs actual hours ITD 

itd_df.columns = itd_df.iloc[0]
itd_df = itd_df[1:].reset_index(drop=True) # Removing first row and using the second as the column headers

itd_projects = itd_df.where((itd_df.iloc[:,1] == 'Rollup')) # Filtering for Rollup only 
itd_projects.dropna(subset = ['Project: Name'], inplace=True) 

#print(itd_projects)

activeproj["Allocated Hrs (ITD)"] = ""
activeproj["Actual Hrs (ITD)"] = ""
activeproj[date] = ""#Add columns to project dataframe

itd_projects['Allocated (ITD)'] = itd_projects['Hours Allocated'].sum(axis=1) #Summing "Hours Allocated" columns 
itd_projects['Actual (ITD)'] = itd_projects['Hours Actual'].sum(axis=1) #Summing "Hours Actual" columns 

itd_projects = itd_projects[['Project: Name', 'Allocated (ITD)', 'Actual (ITD)']] #Filtering out weekly data 

for k, row in activeproj.iterrows(): 
   for g, row in itd_projects.iterrows(): 
        if activeproj['Active Projects'][k][:10] == itd_projects['Project: Name'][g][:10]: 
            activeproj['Allocated Hrs (ITD)'][k] = itd_projects['Allocated (ITD)'][g]
            activeproj['Actual Hrs (ITD)'][k] = itd_projects['Actual (ITD)'][g]

#print(activeproj)

with pd.ExcelWriter('Live Metric Extraction.xlsx', mode='a', engine='openpyxl', if_sheet_exists='replace') as writer: 
  activeproj.to_excel(writer, index=False, sheet_name='Interim Data (This Week)') #Exports projects dataframe to Live Metric Extraction 
