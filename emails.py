import re
import PyPDF2
import pandas as pd
from openpyxl import Workbook

# Open the PDF file
with open('Members List.pdf', 'rb') as file:
    pdf = PyPDF2.PdfReader(file)
    
    email_list = []
    # Iterate over each page
    for page in range(len(pdf.pages)):
        text = pdf.pages[page].extract_text()
        
        # Search for email addresses in the text
        emails = re.findall(r'[\w\.-]+@[\w\.-]+', text)
        
        # Append the found emails to the list
        email_list.extend(emails)

# Create a dataframe from email list
df = pd.DataFrame(email_list,columns=['Emails'])

# Create an Excel file and write the dataframe
book = Workbook()
writer = pd.ExcelWriter('emails.xlsx', engine='openpyxl') 
writer.book = book

df.to_excel(writer, index=False)
writer.save()

print("Emails saved to excel file successfully!")

