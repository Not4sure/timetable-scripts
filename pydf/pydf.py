from PyPDF2 import PdfReader

reader = PdfReader("3course-htf.pdf")

page = reader.pages[0]

text = page.extract_text()

print(text)