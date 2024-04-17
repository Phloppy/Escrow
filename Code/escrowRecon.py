import os
import pandas as pd
import openpyxl
import numpy as np
from fuzzywuzzy import process
from tqdm import tqdm

# Replace with the path to the folder - omit last backslash '\'
input_folder_path = r'C:\Users\jasonjasinski\OneDrive - At World Properties\Documents\Test\Escrow\Input' + '\\'

# Specify file names
cibcFile = 'CIBC.xlsx'
lwFile = 'Lonewolf.xlsx'

# Creates full file paths
cibcPath = os.path.join(input_folder_path, cibcFile)
lwPath = os.path.join(input_folder_path, lwFile)

# Read files at previous paths - sheet name is currently hard coded
cibcinfo = pd.read_excel(cibcPath, sheet_name='cibc')
lonewolf = pd.read_excel(lwPath, sheet_name='lonewolf')

# Specify Output Path
output_folder_path = r'C:\Users\jasonjasinski\OneDrive - At World Properties\Documents\Test\Escrow\Output' + '\\'

# Specify name of output Excel file
output_file = 'escrowRecon.xlsx'

# Concatenate full output path
output_path = os.path.join(output_folder_path, output_file)

# Function to categorize data based on specific keywords in given columns
def categorize_row(row):
    bai_description = row['BAI Description'] if pd.notna(row['BAI Description']) else ''
    detail = row['Detail'] if pd.notna(row['Detail']) else ''

    # Check both conditions
    if "ACH CREDIT" in bai_description.upper() and "EARNNEST" in detail.upper():
        return "EARNNEST"
    return "Other"  # Default category if conditions are not met

# Apply the categorization function
cibcinfo['Category'] = cibcinfo.apply(categorize_row, axis=1)


# Create an ExcelWriter object and use it to write data to separate sheets
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    cibcinfo.to_excel(writer, sheet_name='cibcFull', index=False)
    lonewolf.to_excel(writer, sheet_name='lonewolfFull', index=False)



'''
# Sample Code for conditional comparisons
earnnestCondition = (cibcinfo['A'] > 10) & (df['B'] < 20)
condition2 = (df['A'] <= 10) | (df['B'] >= 20)
'''

'''
# Define choices based on conditions
choices = ['Category 1', 'Category 2']

# The conditions list should be in the same order as the choices
conditions = [condition1, condition2]

# Apply conditions and choices to the new column
df['Category'] = np.select(conditions, choices, default='Other')
'''
