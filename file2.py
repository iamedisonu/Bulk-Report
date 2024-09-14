from pathlib import Path
import pandas as pd
from docxtpl import DocxTemplate
import re

base_dir =Path(__file__).parent
word_template_path =base_dir /"Isomo_Report_Card _Template.docx"
excel_path = base_dir /"Isomo_2023_Y2_Final_Grades.xlsx"
output_dir = base_dir /"Final_Result"

#Created outup folder
output_dir.mkdir(exist_ok=True)

# Convert excel sheet into pandas dataframe
df = pd.read_excel(excel_path, sheet_name="Sheet1")

# iterate over each roe in df and render word document

for record in df.to_dict(orient="records"):
    doc =DocxTemplate(word_template_path)
    doc.render(record)
    output_path =output_dir /f"{record['Name']}.docx"
    doc.save(output_path)