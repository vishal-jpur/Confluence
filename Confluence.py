from docx import Document
from docx.shared import Inches

def create_document():
    document = Document()

    # Add a title
    title = document.add_heading('Your Title Here', level=1)
    title.font.size = Inches(0.5)
    title.font.bold = True

    # Add a paragraph
    paragraph = document.add_paragraph('This is a simple paragraph in your predefined format.')
    paragraph.font.size = Inches(0.12)

    # Save the document
    document.save('output.docx')

if __name__ == '__main__':
    create_document()
