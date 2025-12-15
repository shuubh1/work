from docx import Document
from docx.shared import Inches, Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
import pandas as pd
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

def extract_valuation_table_data(excel_file):
    """
    Analyzes an Excel file (DCF/NAV), extracts the table data as a DataFrame.
    Returns: pandas.DataFrame or None
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

        return target_df

    except Exception as e:
        print(f"Error extracting table data: {e}")
        return None

def create_word_table(doc, df):
    """
    Creates a native Word table from a dataframe.
    """
    # Add a table with rows+1 (for header)
    table = doc.add_table(rows=df.shape[0] + 1, cols=df.shape[1])
    
    # Apply a standard grid style (borders)
    table.style = 'Table Grid'

    # --- Header ---
    for j, col_name in enumerate(df.columns):
        cell = table.cell(0, j)
        cell.text = str(col_name)
        # Bold the header
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.bold = True
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # --- Body ---
    for i, row in enumerate(df.itertuples(index=False)):
        for j, val in enumerate(row):
            cell = table.cell(i + 1, j)
            # Format numbers slightly? 
            # For now, just string conversion
            cell.text = str(val)
            # Center align cells for neatness
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
    return table

def move_table_after(table, paragraph):
    """
    Moves a table element to be immediately after a specific paragraph.
    This is required because python-docx adds tables to the end of the doc by default.
    """
    tbl, p = table._tbl, paragraph._p
    p.addnext(tbl)

def process_word_template(template_path, context, table_data=None):
    """
    Opens a Word template, replaces placeholders, and inserts a NATIVE TABLE
    if table_data (DataFrame) is provided.
    """
    doc = Document(template_path)
    
    # 1. Iterate through all paragraphs
    # We collect paragraphs first to avoid modifying the list while iterating
    paragraphs = list(doc.paragraphs)
    
    for p in paragraphs:
        # --- Table Insertion Logic ---
        if '<<valuation.jpg>>' in p.text and table_data is not None:
            p.text = "" # Clear the placeholder text
            
            # Create a new table (initially at the end of doc)
            new_table = create_word_table(doc, table_data)
            
            # Move it to right after this paragraph
            move_table_after(new_table, p)
            continue

        # --- Standard Text Replacement ---
        for key, value in context.items():
            if key in p.text:
                p.text = p.text.replace(key, str(value))

    # 2. Iterate through existing tables to replace text inside them
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                     for key, value in context.items():
                        if key in p.text:
                            p.text = p.text.replace(key, str(value))
    
    return doc