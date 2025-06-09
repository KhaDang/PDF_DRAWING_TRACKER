import os
import pandas as pd
import re
import openpyxl

# === Set your paths ===
folder_path = r'D:\Users\khamdang\Desktop\JASMINE IMU\CHECKINGDRAWING\PDF'
existing_excel = r'D:\Users\khamdang\Desktop\JASMINE IMU\CHECKINGDRAWING\BOM.xlsx'
output_file = 'comparison_with_drawing_revision.xlsx'

# === Set Excel column names ===
drawing_col = 'Model No.'
revision_col = 'Rev.'

# === Function to split filename into drawing number and revision ===
def split_name_and_revision(filename):
    name = os.path.splitext(filename)[0]
    match = re.match(r'(.+?)([A-Z])$', name)
    if match:
        base, rev = match.groups()
        return base, rev
    else:
        return name, ''  # No revision found

# === Extract PDF filenames and parse them ===
pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.pdf')]
pdf_data = [split_name_and_revision(f) for f in pdf_files]
folder_df = pd.DataFrame(pdf_data, columns=['Drawing Number', 'Folder Revision'])

# === Read Excel data ===
excel_df = pd.read_excel(existing_excel, usecols=[drawing_col, revision_col])
excel_df = excel_df.rename(columns={drawing_col: 'Drawing Number', revision_col: 'Excel Revision'})
excel_df = excel_df.dropna(subset=['Drawing Number'])

# === Merge data on Drawing Number ===
comparison = pd.merge(folder_df, excel_df, on='Drawing Number', how='outer', indicator=True)

# === Categorize ===
same_revision = comparison[
    (comparison['_merge'] == 'both') &
    (comparison['Folder Revision'] == comparison['Excel Revision'])
]

different_revision = comparison[
    (comparison['_merge'] == 'both') &
    (comparison['Folder Revision'] != comparison['Excel Revision'])
]

only_in_folder = comparison[comparison['_merge'] == 'left_only']
only_in_excel = comparison[comparison['_merge'] == 'right_only']

# === Save results to Excel ===
with pd.ExcelWriter(output_file) as writer:
    same_revision.to_excel(writer, sheet_name='Same Revision', index=False)
    different_revision.to_excel(writer, sheet_name='Different Revision', index=False)
    only_in_folder.to_excel(writer, sheet_name='Only in Folder', index=False)
    only_in_excel.to_excel(writer, sheet_name='Only in Excel', index=False)

print(f"Comparison complete. Results saved to '{output_file}'")
