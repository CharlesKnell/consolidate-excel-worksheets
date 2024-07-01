# consolidate-excel-worksheets.py

# tag=charlesknell Version 2
# Credits: article by Gaurav Siyal https://www.makeuseof.com/consolidate-excel-workbooks-python/
# 30 Jun 24 - Updated for Python 3.12.3, added usage instructions, a few helpful changes, and updated comments
# 01 Jul 24 - Create the output folder if it doesn't exist

# Description: This python program will consolidate/merge a large number of .xlsx Excel files, limited
#              only by your computer's memory.

# Usage:  1. python must be installed. See https://python.org
#         2. pandas must be installed. Install at a Windows Command Prompt using "pip install pandas"
#         3. openpyxl must be installed. Install at a Windows Command Prompt using "pip install openpyxl"
#         4. Copy this file into the folder containing the .xlsx files to be consolidated/merged.
#         5. Edit the input_file_path and output_file_path to meet your needs. The output_file_path should be
#            different compared to the input_file_path, for repeatable results.
#         6. Execute this program at a Windows Command Prompt using "python consolidate-excel-worksheets.py"

import pandas as pd
import os

input_file_path = "./"  # currently set to the current working folder
output_file_path = "./output/"  # currently set to a folder in the current working folder named output

# create the output folder if it doesn't exist
if not os.path.exists(output_file_path):
    os.makedirs(output_file_path)

# create a list of all the file references of the input folder
file_list = os.listdir(input_file_path)

# create a new blank consolidated dataframe, to accumulate the Excel file imports
df = pd.DataFrame()

# loop through each file in the file_list
for file in file_list:
    # append data to the data frame from .xlsx suffix files only
    if file.endswith(".xlsx"):
        print("Processing " + str(file))
        # create a new dataframe to read/open each Excel file from the list of files created above
        df1 = pd.read_excel(input_file_path + file)
        # append the individual file data frame to the consolidated data frame
        df = pd.concat([df, df1], ignore_index=False)

# transfer final output to an Excel (xlsx) file on the output path
# to add a 1st column containing the original row indexes, change the following to index=True
df.to_excel(output_file_path + "Consolidated_file.xlsx", index=False)

print("Output written to: " + output_file_path + "Consolidated_file.xlsx")
