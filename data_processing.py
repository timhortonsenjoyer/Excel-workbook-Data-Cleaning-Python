import numpy as np
import pandas as pd 
from pathlib import Path
import os

#Get user string input and convert it into a path object
file_path_string = input("Welcome! Please enter the folder path here:\n") 
file_path = Path(file_path_string)
input("Before running this script, ensure all workbooks have sheets that share the same names! Press enter to continue\n")

#if directory type is invalid, prompt to re-enter
while file_path.exists()==False:
    file_path_string=input("File path not detected. Please try again. Try copying the file path from the File Explorer window: ")
    file_path = Path(file_path_string)

#Get all Excel workbooks in directory provided       
workbook_list = [pd.ExcelFile(f) for f in file_path.glob("*.xlsx")]

#Gets a list of string names, where the names are for each common sheet among the workbooks
sheet_category = input("All workbooks have been found, please list all sheet names:\n").split(",")

#Creates a dictionary of key-value pairs. The key being each name of spreadsheet, and value is the dataframe
dataframe_dict ={}

for name in sheet_category:
    #Gets each sheet that matches the user input, and puts them all in a list
    sheet_list=[]
    for excel_workbook in workbook_list:
        spreadsheet_found = pd.read_excel(excel_workbook,sheet_name=name)
        sheet_list.append(spreadsheet_found)

    #Appends them together, and cleans out any empty rows that tend to come from Excel's formatting
    combined_df = pd.concat(sheet_list,ignore_index=True)
    combined_df = combined_df.replace(r'^\s*$',np.nan,regex=True)

    #Removes the empty rows that came from Excel's formatting based on a pre-determined row that is supposed to be full regardless. Can be substituted again here with a user input
    combined_df = combined_df.dropna(subset=['important date column'])

    #Converts the string datetime info into pandas version, filters errors and sorts 
    combined_df['important date column'] = combined_df['important date column'].apply(lambda x: pd.to_datetime(x,dayfirst=True,errors='coerce'))
    combined_df.sort_values(by='important date column',na_position='last')

    #Replaces the original Number of Entries column with a new one for the final dataframe
    final_row_count = range(1,len(combined_df)+1)
    combined_df.loc[:,'No']= list(final_row_count)
    dataframe_dict[name] = combined_df

#Combines all the end result dataframes into one Excel workbook
file_location = os.path.join(file_path_string,'merged_output.xlsx')
with pd.ExcelWriter(file_location,engine='openpyxl') as writer:
    for name in sheet_category:
        dataframe_dict[name].to_excel(writer,sheet_name=name,index=False)

input("Process finished! You can view the final sheet in the same folder entered before. Press enter to exit\n")