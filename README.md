# consolidate-excel-worksheets.py

Credits: Original program by Gaurav Siyal 
(https://www.makeuseof.com/consolidate-excel-workbooks-python/)

30 Jun 24 - Updated for Python 3.12.3, added usage instructions, a few helpful changes, and updated comments.
01 Jul 24 - Create the output folder if it doesn't exist.

## Description
This python program will consolidate/merge a large number of .xlsx Excel files, limited only by your computer's memory.

## Usage
1. python must be installed. See (https://python.org)
2. pandas must be installed. Install at a Windows Command Prompt using "pip install pandas".
3. openpyxl must be installed. Install at a Windows Command Prompt using "pip install openpyxl".
4. Copy this file into the folder containing the .xlsx files to be consolidated or merged.
5. Edit the input_file_path and output_file_path to meet your needs. The output_file_path should be different compared to the input_file_path, for repeatable results.
6. Execute this program at a Windows Command Prompt using "python consolidate-excel-worksheets.py"

