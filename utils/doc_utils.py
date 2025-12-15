from docx import Document
from docx.shared import Inches
import pandas as pd
import matplotlib.pyplot as plt
import io

def generate_valuation_table_image(excel_file):
    """
    Analyzes an Excel file to determine if it's DCF or NAV,
    extracts the relevant table, and converts it to an image stream.
    """
    try:
        # Load Excel file to check sheet names
        xls = pd.ExcelFile(excel_file)
        sheet_names = xls.sheet_names
        
        target_df = None
        
        # Logic 1: Check for DCF
        # "if the keyword is 'DCF' then make the table in the sheet 'Financials' into an image"
        dcf_sheets = [s for s in sheet_names if 'DCF' in s.upper()]
        if dcf_sheets:
            if 'Financials' in sheet_names:
                target_df = pd.read_excel(xls, 'Financials')
            else:
                # Fallback if 'Financials' isn't exact match
                return None 

        # Logic 2: Check for NAV (Only if DCF didn't trigger)
        # "if its 'NAV' then insert the table from that sheet itself"
        if target_df is None:
            nav_sheets = [s for s in sheet_names if 'NAV' in s.upper()]
            if nav_sheets:
                # Use the first sheet that matched 'NAV'
                target_df = pd.read_excel(xls, nav_sheets[0])

        if target_df is None:
            return None

        # --- Generate Image from DataFrame ---
        
        # Cleanup: Drop completely empty rows/cols and fill NaNs
        target_df = target_df.dropna(how='all').dropna(axis=1, how='all')
        target_df = target_df.fillna('')

        # Create plot
        # Height is dynamic based on number of rows
        fig, ax = plt.subplots(figsize=(10, len(target_df) * 0.5 + 2)) 
        ax.axis('off')
        
        # Create table
        table = ax.table(
            cellText=target_df.values, 
            colLabels=target_df.columns, 
            cellLoc='center', 
            loc='center',
            colColours=['#f2f2f2']*len(target_df.columns)
        )
        
        table.auto_set_font_size(False)
        table.set_fontsize(10)
        table.scale(1.2, 1.5)

        # Save to buffer
        img_buffer = io.BytesIO()
        plt.savefig(img_buffer, format='png', bbox_inches='tight', dpi=300)
        img_buffer.seek(0)
        plt.close(fig)
        
        return img_buffer

    except Exception as e:
        # Print error to console for debugging, return None so app doesn't crash
        print(f"Error generating table image: {e}")
        return None

def process_word_template(template_path, context, image_stream=None):
    """
    Opens a Word template, replaces placeholders with context values,
    and inserts an image if the specific placeholder <<valuation.jpg>> is found.
    """
    doc = Document(template_path)
    
    # 1. Iterate through all paragraphs
    for p in doc.paragraphs:
        # Check for image placeholder specifically
        if '<<valuation.jpg>>' in p.text and image_stream:
            p.text = "" # Clear text
            run = p.add_run()
            # Reset stream position just in case
            image_stream.seek(0)
            run.add_picture(image_stream, width=Inches(6.0)) 
            continue

        # Standard Text Replacement
        for key, value in context.items():
            if key in p.text:
                p.text = p.text.replace(key, str(value))

    # 2. Iterate through tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                     for key, value in context.items():
                        if key in p.text:
                            p.text = p.text.replace(key, str(value))
    
    return doc