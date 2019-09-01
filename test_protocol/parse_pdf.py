import pdfplumber
pdf = pdfplumber.open('E:\Temp\\test2.pdf')
page = pdf.pages[0]
print(page.extract_table())
# for table in pdf
