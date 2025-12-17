import streamlit as st
import io
import os
import zipfile
import pandas as pd
from utils.auth_manager import require_auth
from utils.doc_utils import process_word_template, extract_valuation_table_data, generate_financial_table_image

st.set_page_config(page_title="Document Gen", page_icon="📝", layout="wide")

# --- AUTHENTICATION CHECK ---
require_auth()
# ----------------------------

st.title("Document Generation Suite")
st.markdown("Fill in the details below to generate standard firm documents.")

# ==================================================
# SECTION A: UPLOAD & PREVIEW (Interactive)
# Moved outside form for instant feedback
# ==================================================
st.subheader("1. Upload Valuation Workings")
st.info("Script looks for a **'DCF'** tab (then finds 'Financials') OR a **'NAV'** tab.")

valuation_excel = st.file_uploader(
    "Upload Excel File", 
    type=['xlsx', 'xls'],
    label_visibility="collapsed"
)

# Variable to hold the extracted data for later use
table_df = None

if valuation_excel:
    with st.spinner("Scanning Excel file and generating preview..."):
        # We perform extraction immediately upon upload
        table_df = extract_valuation_table_data(valuation_excel)
        
    if table_df is not None:
        st.success(f"✅ Data Found! Will generate a table with {table_df.shape[0]} rows and {table_df.shape[1]} columns.")
        
        # --- PREVIEWER ---
        with st.expander("👁️ Preview Generated Image (Click to Expand)", expanded=True):
            st.caption("This is the exact image that will be pasted into the Word document.")
            
            # Generate the image buffer just for preview
            img_buffer = generate_financial_table_image(table_df)
            st.image(img_buffer, caption="Table Preview", use_column_width=False)
            
            # Reset pointer not needed for buffer, but good practice if we re-use logic
    else:
        st.error("❌ Could not find a valid 'DCF' (Financials sheet) or 'NAV' sheet in this file.")

st.divider()

# ==================================================
# SECTION B: DOCUMENT DETAILS FORM
# ==================================================
with st.form("doc_gen_form"):
    st.subheader("2. Document Details")
    
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**Date Details**")
        val_date = st.date_input("Valuation Date (Mgmt Rep Letter)", key="d1")
        br_date = st.date_input("Board Resolution Date", key="d2")
        eng_date = st.date_input("Engagement Letter Date", key="d3")
        
    with col2:
        st.markdown("**Client Details**")
        company = st.text_input("Company Name", "My Client Company Pvt Ltd")
        valuation_statement = st.text_input("Statement for Valuation", "Preferential Allotment of Equity Shares")
        
    st.divider()
    
    st.subheader("3. Addressee Details")
    col_a, col_b = st.columns(2)
    with col_a:
        authority = st.text_input("Addressee Name (Authority)", "Mr. John Doe")
        designation = st.text_input("Addressee Designation", "Director")
    with col_b:
        addressed_to = st.text_input("Salutation (e.g. Sir/Ma'am)", "Sir")
        group_designation = st.text_input("Group Designation (e.g. Board of Directors)", "Board of Directors")

    st.subheader("4. Address Details")
    address1 = st.text_input("Address Line 1", "123 Business Park")
    col3, col4 = st.columns(2)
    with col3:
        address2 = st.text_input("Address Line 2", "Financial District")
    with col4:
        address3 = st.text_input("Address Line 3", "New York, NY 10001")
            
    st.subheader("5. Select Documents to Generate")
    doc_opts = {
        "Board Resolution": "Board Resolution.docx",
        "Engagement Letter": "Engagement Letter.docx",
        "Management Rep Letter": "Management-representation-letter.docx"
    }
    selected_docs = st.multiselect("Choose files:", list(doc_opts.keys()), default=list(doc_opts.keys()))

    submitted = st.form_submit_button("Generate Documents", type="primary")

# ==================================================
# SECTION C: PROCESSING
# ==================================================
if submitted:
    if not selected_docs:
        st.error("Please select at least one document to generate.")
        st.stop()

    full_address = f"{address1}, {address2}, {address3}"
    
    # Base Context
    base_context = {
        "<<company>>": company,
        "<<company_caps>>": company.upper(),
        "<<valuation_type_statement>>": valuation_statement,
        "<<authority>>": authority,
        "<<designation>>": designation,
        "<<addressed_to>>": addressed_to,
        "<<address_caps>>": full_address.upper(),
        "<<address>>": full_address,
        "<<address1>>": address1,
        "<<address2>>": address2,
        "<<address3>>": address3,
        "<<group_designation>>": group_designation,
        "<<authority_designation>>": designation,
        "<<valuation_date>>": val_date.strftime("%d-%b-%Y"),
        "<<br_date>>": br_date.strftime("%d-%b-%Y"),
        "<<engagement_date>>": eng_date.strftime("%d-%b-%Y"),
    }

    template_dir = "templates"
    if not os.path.exists(template_dir):
        st.error(f"Error: '{template_dir}' folder not found.")
        st.stop()

    # NOTE: We use 'table_df' here, which was extracted in Section A (outside the form)
    # This is efficient because we don't need to re-read the Excel file.
    if valuation_excel and table_df is None:
        # If file was uploaded but df is None, it means extraction failed silently or previously
        st.warning("Warning: Excel file provided but no table data was extracted. The table in the document will be empty.")

    generated_files = []

    try:
        for doc_name in selected_docs:
            filename = doc_opts[doc_name]
            file_path = os.path.join(template_dir, filename)
            
            if not os.path.exists(file_path):
                st.error(f"Template not found: {filename}")
                continue
            
            # Context Switching Logic for Dates
            current_context = base_context.copy()
            
            if doc_name == "Management Rep Letter":
                current_context["<<date>>"] = base_context["<<valuation_date>>"]
            elif doc_name == "Board Resolution":
                current_context["<<date>>"] = base_context["<<br_date>>"]
            elif doc_name == "Engagement Letter":
                current_context["<<date>>"] = base_context["<<engagement_date>>"]
                current_context["<<date >>"] = base_context["<<engagement_date>>"]
            else:
                current_context["<<date>>"] = base_context["<<valuation_date>>"] 
            
            # Process Document (Passing DataFrame now)
            # The doc_utils.process_word_template will handle generating the image from the DF
            doc = process_word_template(file_path, current_context, table_df)
            
            doc_io = io.BytesIO()
            doc.save(doc_io)
            doc_io.seek(0)
            generated_files.append((f"Generated_{filename}", doc_io))

        if not generated_files:
            st.error("No files were generated.")
        elif len(generated_files) == 1:
            st.success("Document generated successfully!")
            st.download_button(
                label=f"Download {generated_files[0][0]}",
                data=generated_files[0][1],
                file_name=generated_files[0][0],
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        else:
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zf:
                for name, data in generated_files:
                    zf.writestr(name, data.getvalue())
            zip_buffer.seek(0)
            
            st.success(f"{len(generated_files)} documents generated successfully!")
            st.download_button(
                label="Download All (ZIP)",
                data=zip_buffer,
                file_name="generated_documents.zip",
                mime="application/zip"
            )

    except Exception as e:
        st.error(f"An error occurred during generation: {str(e)}")