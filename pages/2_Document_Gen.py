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
st.markdown("### 1. Valuation Table Input")
st.info("💡 **Pro Tip:** You can copy your Excel table (using Snipping Tool or 'Copy as Picture') and **press Ctrl+V** directly into the uploader below.")

# ==================================================
# SECTION A: DYNAMIC INPUT (Paste Image OR Upload Excel)
# ==================================================
col_input, col_preview = st.columns([1, 1])

# Global variables to store the final assets
final_table_df = None
final_image_buffer = None

with col_input:
    # We use a single uploader that accepts images AND excel
    uploaded_file = st.file_uploader(
        "Paste Image (Ctrl+V) or Upload File", 
        type=['png', 'jpg', 'jpeg', 'xlsx', 'xls'],
        help="Click here and press Ctrl+V to paste an image from your clipboard, or drag and drop a file."
    )

    if uploaded_file:
        file_type = uploaded_file.type
        
        # CASE 1: IMAGE (User pasted or uploaded a screenshot)
        if "image" in file_type:
            final_image_buffer = uploaded_file
            st.success("✅ Image captured!")

        # CASE 2: EXCEL (User uploaded raw data)
        elif "spreadsheet" in file_type or "excel" in file_type:
            with st.spinner("Extracting table from Excel..."):
                final_table_df = extract_valuation_table_data(uploaded_file)
                if final_table_df is not None:
                    st.success(f"✅ Extracted table: {final_table_df.shape[0]} rows x {final_table_df.shape[1]} cols")
                else:
                    st.error("❌ No valid 'Financials' or 'NAV' sheet found.")

with col_preview:
    # Show preview based on what we have
    if final_image_buffer:
        st.image(final_image_buffer, caption="Preview: This image will be used in the document", use_column_width=True)
    elif final_table_df is not None:
        st.caption("Preview: Auto-generated table image from Excel")
        # Generate a preview image on the fly
        preview_img = generate_financial_table_image(final_table_df)
        st.image(preview_img, use_column_width=True)
    else:
        st.info("No content detected yet. Waiting for paste or upload...")

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
    
    # Validation: Do we have *something* to put in the table?
    if final_image_buffer is None and final_table_df is None:
        st.warning("⚠️ No table data provided (Image or Excel). The `<<valuation.jpg>>` section will be empty.")

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
            
            # Process Document
            # We pass BOTH. Logic inside doc_utils decides priority (Image > DF).
            doc = process_word_template(
                file_path, 
                current_context, 
                table_data=final_table_df, 
                provided_image_stream=final_image_buffer
            )
            
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