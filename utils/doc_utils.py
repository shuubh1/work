from docx import Document
from docx.shared import Inches

def process_word_template(template_path, context, image_stream=None):
    """
    Opens a Word template, replaces placeholders with context values,
    and inserts an image if the specific placeholder <<valuation.jpg>> is found.
    """
    doc = Document(template_path)
    
    # 1. Iterate through all paragraphs (Standard body text)
    for p in doc.paragraphs:
        # Check for image placeholder specifically
        if '<<valuation.jpg>>' in p.text and image_stream:
            p.text = "" # Clear the placeholder text
            run = p.add_run()
            run.add_picture(image_stream, width=Inches(5.0)) 
            continue

        # Standard Text Replacement
        for key, value in context.items():
            if key in p.text:
                # Simple replacement
                p.text = p.text.replace(key, str(value))

    # 2. Iterate through tables (Word docs often use tables for layout)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                     for key, value in context.items():
                        if key in p.text:
                            p.text = p.text.replace(key, str(value))
    
    return doc