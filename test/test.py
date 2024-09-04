import openpyxl
from PyPDF2 import PdfReader

# Open the existing Excel workbook
workbook = openpyxl.load_workbook('your_excel_file.xlsx')

# Select the desired sheet
sheet = workbook.active

# Specify the PDF file you want to insert
pdf_file_path = 'your_pdf_file.pdf'

# Open the PDF file and read its content
with open(pdf_file_path, 'rb') as pdf_file:
    pdf_reader = PdfReader(pdf_file)
    
    # Extract text from the PDF
    pdf_text = ''
    for page_num in range(len(pdf_reader.pages)):
        pdf_text += pdf_reader.pages[page_num].extract_text()

# Specify the cell where you want to insert the PDF content
cell = sheet['A1']

# Set the cell value to the extracted PDF text
cell.value = pdf_text

# Save the modified Excel workbook
workbook.save('output_excel_file.xlsx')
