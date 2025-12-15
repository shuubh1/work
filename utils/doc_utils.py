from docx import Document
from docx.shared import Inches
import pandas as pd
import matplotlib.pyplot as plt
import io

def clean_and_trim_df(df):
    """
    Cleans the DataFrame to remove empty rows/cols and whitespace.
    Returns a dataframe containing ONLY the used range.
    """
    # 1. Replace strings that are just whitespace with NaN
    df = df.replace(r'^\s*$', pd.NA, regex=True)
    
    # 2. Drop rows that are ENTIRELY empty (NaN)
    df = df.dropna(how='all', axis=0)
    
    # 3. Drop columns that are ENTIRELY empty
    df = df.dropna(how='all', axis=1)
    
    # 4. Fill remaining NaNs with empty strings for clean display
    df = df.fillna('')
    
    return df

def generate_valuation_table_image(excel_file):
    """
    Analyzes an Excel file (DCF/NAV), extracts the table, trims it tightly,
    and converts it to a high-quality image stream.
    """
    try:
        xls = pd.ExcelFile(excel_file)
        sheet_names = xls.sheet_names
        target_df = None
        
        # --- Sheet Selection Logic ---
        
        # 1. Check for DCF -> Financials
        dcf_sheets = [s for s in sheet_names if 'DCF' in s.upper()]
        if dcf_sheets:
            if 'Financials' in sheet_names:
                target_df = pd.read_excel(xls, 'Financials')
            # If no 'Financials' sheet, we don't default to the first DCF sheet 
            # to avoid grabbing raw calculation data.
            
        # 2. Check for NAV (Fallback)
        if target_df is None:
            nav_sheets = [s for s in sheet_names if 'NAV' in s.upper()]
            if nav_sheets:
                target_df = pd.read_excel(xls, nav_sheets[0])

        if target_df is None:
            return None

        # --- Data Cleaning ---
        target_df = clean_and_trim_df(target_df)
        
        if target_df.empty:
            return None

        # --- Image Generation (Tight Layout) ---
        
        rows, cols = target_df.shape
        
        # Dynamic sizing: 
        # Width: 10 inches (standard matplotlib units, not print inches)
        # Height: calculated based on row count to ensure rows aren't squashed
        # 0.5 unit per row + 1 unit for header/padding
        calculated_height = (rows * 0.5) + 1 
        
        fig, ax = plt.subplots(figsize=(12, calculated_height))
        ax.axis('off')
        ax.axis('tight')
        
        # Create the table
        table = ax.table(
            cellText=target_df.values, 
            colLabels=target_df.columns, 
            cellLoc='center', 
            loc='center',
            colColours=['#f2f2f2']*cols # Light gray header
        )
        
        # Styling to remove "jankiness"
        table.auto_set_font_size(False)
        table.set_fontsize(10)
        table.scale(1, 1.5) # Scale height (1.5) to give text breathing room
        
        # Auto-adjust column widths to fit content
        table.auto_set_column_width(col=list(range(cols)))

        # Save tightly to buffer
        img_buffer = io.BytesIO()
        # bbox_inches='tight' removes the white frame around the plot
        # pad_inches=0.05 gives a TINY margin so borders aren't cut off
        plt.savefig(img_buffer, format='png', bbox_inches='tight', pad_inches=0.05, dpi=300)
        img_buffer.seek(0)
        plt.close(fig)
        
        return img_buffer

    except Exception as e:
        print(f"Error generating table image: {e}")
        return None

def process_word_template(template_path, context, image_stream=None):
    """
    Opens a Word template, replaces placeholders, and inserts an image.
    """
    doc = Document(template_path)
    
    # 1. Iterate through all paragraphs
    for p in doc.paragraphs:
        if '<<valuation.jpg>>' in p.text and image_stream:
            p.text = "" # Clear placeholder
            run = p.add_run()
            
            image_stream.seek(0) # Ensure we read from start
            
            # Using 5.5 inches is safer for A4 margins (usually 6.0 is max content width)
            # This helps prevent the "pushed to next page" issue
            run.add_picture(image_stream, width=Inches(5.5)) 
            continue

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