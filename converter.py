import pandas as pd
from datetime import datetime
import pytz
from docx import Document
import os

def extract_excel_rows(file_path):
    df = pd.read_csv(file_path)
    for _, row in df.iterrows():
        yield row.tolist()

def get_current_date_chicago():
    chicago_tz = pytz.timezone('America/Chicago')
    now_utc = datetime.utcnow().replace(tzinfo=pytz.utc)
    now_chicago = now_utc.astimezone(chicago_tz)
    return now_chicago.strftime("%B %d, %Y")

def replace_text_in_docx(template_path, output_string, name):
    # Load the Word document
    doc = Document(template_path)

    # Replace "[REPLACE]" with output_string
    for paragraph in doc.paragraphs:
        if '[REPLACE]' in paragraph.text:
            paragraph.text = paragraph.text.replace('[REPLACE]', output_string)

    # Define the output directory
    output_dir = "Closing Letters"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Save the document with a new filename based on the name in a cross-OS compatible path
    new_file_path = os.path.join(output_dir, f"{name} Closing Letter.docx")
    doc.save(new_file_path)

# Read the template from 'letter_template.txt'
with open('letter_template.txt', 'r') as file:
    template = file.read()

# Replace 'your_file.xlsx' with your Excel file path
row_generator = extract_excel_rows('Closing letter data.csv')

# Replace 'word_template.docx' with your Word template file path
template_path = 'word_template.docx'

# Loop until there are no more rows
while True:
    try:
        row = next(row_generator)
        output_string = template.format(date=get_current_date_chicago(), name=row[0], add1=row[1], add2=row[2])
        replace_text_in_docx(template_path, output_string, row[0])
        print(f"Document for {row[0]} created.")
    except StopIteration:
        break  # No more rows
