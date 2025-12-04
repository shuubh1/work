import pandas as pd
import matplotlib.pyplot as plt
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches
import io
import os

def clean_currency(x):
    """Helper to clean currency strings from Excel if necessary"""
    if isinstance(x, (int, float)):
        return x
    if isinstance(x, str):
        return x.replace('₹', '').replace(',', '').strip()
    return x

def generate_nav_table_image(excel_file):
    """
    Reads the 'NAV Calculation Working' sheet from the uploaded Excel
    and generates a Matplotlib image of the table.
    """
    try:
        # Load the specific sheet
        df = pd.read_excel(excel_file, sheet_name='NAV Calculation Working')
        
        # specific cleanup based on your script's logic
        # Dropping completely empty rows/cols
        df = df.dropna(how='all').dropna(axis=1, how='all')
        
        # Fill NaNs with empty strings for better display
        df = df.fillna('')

        # Create the plot
        fig, ax = plt.subplots(figsize=(10, len(df) * 0.5 + 2)) # Dynamic height
        ax.axis('off')
        
        # Create the table
        table = ax.table(
            cellText=df.values, 
            colLabels=df.columns, 
            cellLoc='center', 
            loc='center',
            colColours=['#f2f2f2']*len(df.columns) # Light gray header
        )
        
        table.auto_set_font_size(False)
        table.set_fontsize(10)
        table.scale(1.2, 1.5) # Adjust scaling for readability

        # Save to buffer
        img_buffer = io.BytesIO()
        plt.savefig(img_buffer, format='png', bbox_inches='tight', dpi=300)
        img_buffer.seek(0)
        plt.close(fig)
        
        return img_buffer

    except Exception as e:
        raise Exception(f"Error generating NAV table: {str(e)}")

def generate_valuation_report(template_path, context, nav_image_buffer):
    """
    Renders the docxtpl template with text context and the generated image.
    """
    doc = DocxTemplate(template_path)
    
    # Insert the image into the context object using docxtpl's InlineImage
    if nav_image_buffer:
        # Note: We use a placeholder key 'nav_image' here. 
        # Ensure your Word doc uses {{ nav_image }} or we map it correctly.
        # Based on your file, it uses {{nav.jpg}}
        
        # We need to write the buffer to a temp file because InlineImage often prefers paths
        # or we can try passing the stream directly depending on version.
        # For safety/stability with docxtpl, saving to a temp file is often most reliable.
        temp_img_name = "temp_nav_table.png"
        with open(temp_img_name, "wb") as f:
            f.write(nav_image_buffer.getbuffer())
            
        context['nav.jpg'] = InlineImage(doc, temp_img_name, width=Inches(6.0))
    
    # Render
    doc.render(context)
    
    # Save to memory
    output_io = io.BytesIO()
    doc.save(output_io)
    output_io.seek(0)
    
    # Cleanup temp image
    if os.path.exists("temp_nav_table.png"):
        os.remove("temp_nav_table.png")
        
    return output_io