from docx import Document
from docx.shared import Inches, Pt
import pandas as pd
import io
import matplotlib.pyplot as plt
from pandas.plotting import table

def process_word_template(template_path, context, table_data=None, provided_image_stream=None):
    """
    Opens a Word template, replaces placeholders.
    Inserts a table image from 'provided_image_stream' (Direct user upload).
    """
    doc = Document(template_path)
    
    # 1. Iterate through all paragraphs
    paragraphs = list(doc.paragraphs)
    
    for p in paragraphs:
        # --- Table Image Insertion Logic ---
        # Look for the placeholder <<valuation.jpg>>
        if '<<valuation.jpg>>' in p.text and provided_image_stream is not None:
            p.text = "" # Clear the placeholder text
            run = p.add_run()
            
            # Reset stream pointer to ensure we read from start if reused
            provided_image_stream.seek(0)
            
            # Add picture with a set width (6 inches fits standard margins)
            run.add_picture(provided_image_stream, width=Inches(6.0))
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

# --- Legacy Helper Functions (Preserved but not active in simplified UI) ---

def clean_and_trim_df(df):
    df = df.replace(r'^\s*$', pd.NA, regex=True)
    df = df.dropna(how='all', axis=0)
    df = df.dropna(how='all', axis=1)
    df = df.fillna('')
    return df

def extract_valuation_table_data(excel_file):
    try:
        xls = pd.ExcelFile(excel_file)
        sheet_names = xls.sheet_names
        target_df = None
        dcf_sheets = [s for s in sheet_names if 'DCF' in s.upper()]
        if dcf_sheets and 'Financials' in sheet_names:
            target_df = pd.read_excel(xls, 'Financials', header=None)
        if target_df is None:
            nav_sheets = [s for s in sheet_names if 'NAV' in s.upper()]
            if nav_sheets:
                target_df = pd.read_excel(xls, nav_sheets[0], header=None)
        if target_df is not None:
            return clean_and_trim_df(target_df)
        return None
    except Exception as e:
        print(f"Error extracting table data: {e}")
        return None

def generate_financial_table_image(df):
    df = df.fillna('')
    w = max(len(df.columns) * 2.0, 8) 
    h = max((len(df) + 1) * 0.4, 3)
    fig, ax = plt.subplots(figsize=(w, h))
    ax.axis('off')
    tbl = table(ax, df, loc='center', cellLoc='center', colWidths=[1.0/len(df.columns)] * len(df.columns))
    tbl.auto_set_font_size(False)
    tbl.set_fontsize(11)
    tbl.scale(1.2, 1.5)
    for (row, col), cell in tbl.get_celld().items():
        cell.set_edgecolor('white')
        cell.set_linewidth(0)
        if row == 0:
            cell.set_text_props(weight='bold', color='black')
            cell.set_facecolor('white')
            cell.set_edgecolor('black') 
            cell.set_linewidth(1.5)
            cell.visible_edges = "B" 
        else:
            cell.set_facecolor('white')
            cell.set_text_props(color='#333333')
            cell.set_edgecolor('#e0e0e0')
            cell.set_linewidth(0.5)
            cell.visible_edges = "B"
    buf = io.BytesIO()
    plt.savefig(buf, format='png', bbox_inches='tight', dpi=300, pad_inches=0.1)
    buf.seek(0)
    plt.close(fig)
    return buf