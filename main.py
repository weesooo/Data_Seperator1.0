import time
import pandas as pd
import os
from datetime import datetime


time.sleep(1)
print('Files under current directory:')
files = [f for f in os.listdir('.') if os.path.isfile(f)]
print(files)
user_input_workbook = input("Enter the name of your excel workbook file (.xlsx): ")
user_input_sheet_name = input("Enter the name of your excel workbook sheet name: ")
time.sleep(2)
__location__ = os.path.realpath(
    os.path.join(os.getcwd(), os.path.dirname(__file__)))
excel_file_path = os.path.join(__location__, user_input_workbook)
print('Excel File Path Specified: ' + '\n' + excel_file_path)
time.sleep(2)
print('Excel File Sheet Name Specified: ' + '\n' + user_input_sheet_name)
time.sleep(3)
print('time to read in the excel file')
time.sleep(1)
try:
    df = pd.read_excel(excel_file_path, sheet_name = user_input_sheet_name )
    time.sleep(2)
    df_cols = df.columns
    print('\n')
    print('DataFrame Columns: '+ df_cols)
    print('\n')
    user_input_column_name = input('Enter name of Column you want to run the script on: ')
    user_input_seperator = input('Enter Name of seperator? for example: AND,OR,ALSO' + '\n')
    print('Column name to Explode: ' + str(user_input_column_name))
    print('Row Seperator: ' + str(user_input_seperator))
    new_df = df.assign(Location=df.Location.str.split(user_input_seperator)).explode(str(user_input_column_name))
    time.sleep(3)
    print('New DF rows for spot checking: ')
    print(new_df.head(5))
    time.sleep(6)
    print('Converting Pandas dataframe to excel workbook and saving the new output:')
    time.sleep(2)
    output_name = (user_input_column_name+'_'+user_input_seperator+'_'+'DONE'+'.xlsx')
    time.sleep(2)
    print('Expecetd output name of the file: '+ output_name)
    time.sleep(4)
    new_df.to_excel(output_name)
except Exception as e:
    print('Unexpected error:' + str(e))
    time.sleep(10)

