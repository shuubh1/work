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
# SECTION A: PROVIDE TABLE DATA
# ==================================================
st.subheader("1. Provide Valuation Table")

# Toggle between Manual Image or Excel Auto-Gen
input_method = st.radio(
    "Select Table Source:",
    ["Upload Image (Recommended)", "Upload Excel (Auto-Generate)"],
    horizontal=True
)

# Variables to hold the final assets
table_df = None
uploaded_image_buffer = None

if input_method == "Upload Image (Recommended)":
    st.info("💡 **Tip:** In Excel, select your table > Copy as Picture > Save as Image (or Paste here if supported). This guarantees the exact formatting.")
    
    uploaded_image = st.file_uploader(
        "Upload Table Screenshot", 
        type=['png', 'jpg', 'jpeg'],
        help="Upload a screenshot or saved image of your Excel table."
    )
    
    if uploaded_image:
        uploaded_image_buffer = uploaded_image
        with st.expander("👁️ Preview Uploaded Image", expanded=True):
            st.image(uploaded_image, caption="This image will be inserted into the document.", use_column_width=False)

else:
    # --- OLD EXCEL METHOD ---
    st.info("Script looks for a **'DCF'** tab (then finds 'Financials') OR a **'NAV'** tab.")
    
    valuation_excel = st.file_uploader(
        "Upload Excel File", 
        type=['xlsx', 'xls'],
        label_visibility="collapsed"
    )

    if valuation_excel:
        with st.spinner("Scanning Excel file and generating preview..."):
            table_df = extract_valuation_table_data(valuation_excel)
            
        if table_df is not None:
            st.success(f"✅ Data Found! Will generate a table with {table_df.shape[0]} rows and {table_df.shape[1]} columns.")
            
            with st.expander("👁️ Preview Generated Image (Click to Expand)", expanded=True):
                st.caption("This is the image generated from your Excel data.")
                # Generate the image buffer just for preview
                img_buffer = generate_financial_table_image(table_df)
                st.image(img_buffer, caption="Table Preview", use_column_width=False)
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

    # NOTE: Validation Check
    # If Excel mode was chosen but no DF found, OR Image mode chosen but no Image found
    if input_method == "Upload Excel (Auto-Generate)" and table_df is None:
        st.warning("⚠️ Warning: Excel method selected but no table data was extracted. Document table will be empty.")
    elif input_method == "Upload Image (Recommended)" and uploaded_image_buffer is None:
        st.warning("⚠️ Warning: Image method selected but no image uploaded. Document table will be empty.")

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
            # We pass BOTH possible sources. The function decides which to use based on what is not None.
            # We ensure we only pass the one corresponding to the user's choice.
            
            df_to_pass = table_df if input_method == "Upload Excel (Auto-Generate)" else None
            img_to_pass = uploaded_image_buffer if input_method == "Upload Image (Recommended)" else None
            
            doc = process_word_template(
                file_path, 
                current_context, 
                table_data=df_to_pass, 
                provided_image_stream=img_to_pass
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