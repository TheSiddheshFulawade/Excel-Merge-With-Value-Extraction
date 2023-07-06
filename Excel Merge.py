import os
import pandas as pd

#Path to your excel files
directory = r'C:\Users\swara\Untitled Folder\Excel'

values = []
file_names = []

for filename in os.listdir(directory):
    if filename.endswith('.xlsx') or filename.endswith('.xls'):
        file_path = os.path.join(directory, filename)
        
        # Read the Excel file
        df = pd.read_excel(file_path, header=None)
        
        for row in df.values:
            for cell in row:
                if isinstance(cell, str) and cell.startswith('MUM'): #String starting with MUM
                    values.append(cell)
                    file_names.append(filename)

output_df = pd.DataFrame({'Value': values, 'File Name': file_names})

output_df.to_excel('Merged.xlsx', index=False)  #Creating the Excel File
