import streamlit as st
import io
import os
import zipfile
from utils.auth_manager import require_auth
from utils.doc_utils import process_word_template, generate_valuation_table_image

st.set_page_config(page_title="Document Gen", page_icon="📝", layout="wide")

# --- AUTHENTICATION CHECK ---
require_auth()
# ----------------------------

st.title("Document Generation Suite")
st.markdown("Fill in the details below to generate standard firm documents.")

# --- Input Form ---
with st.form("doc_gen_form"):
    st.subheader("1. General Information")
    col1, col2 = st.columns(2)
    with col1:
        # New Date Logic
        st.markdown("**Date Details**")
        val_date = st.date_input("Valuation Date (Mgmt Rep Letter)", key="d1")
        br_date = st.date_input("Board Resolution Date", key="d2")
        eng_date = st.date_input("Engagement Letter Date", key="d3")
        
    with col2:
        st.markdown("**Client Details**")
        company = st.text_input("Company Name", "My Client Company Pvt Ltd")
        valuation_statement = st.text_input("Statement for Valuation", "Preferential Allotment of Equity Shares")
        
    st.divider()
    
    st.subheader("2. Addressee Details")
    col_a, col_b = st.columns(2)
    with col_a:
        authority = st.text_input("Addressee Name (Authority)", "Mr. John Doe")
        designation = st.text_input("Addressee Designation", "Director")
    with col_b:
        addressed_to = st.text_input("Salutation (e.g. Sir/Ma'am)", "Sir")
        group_designation = st.text_input("Group Designation (e.g. Board of Directors)", "Board of Directors")

    st.subheader("3. Address Details")
    address1 = st.text_input("Address Line 1", "123 Business Park")
    col3, col4 = st.columns(2)
    with col3:
        address2 = st.text_input("Address Line 2", "Financial District")
    with col4:
        address3 = st.text_input("Address Line 3", "New York, NY 10001")
            
    st.subheader("4. Specifics & Uploads")
    # Changed from Image Upload to Excel Upload
    valuation_excel = st.file_uploader(
        "Upload Valuation Workings (Excel) for <<valuation.jpg>>", 
        type=['xlsx', 'xls'],
        help="Script looks for 'DCF' (uses 'Financials' sheet) or 'NAV' (uses that sheet) to create the table image."
    )

    st.subheader("5. Select Documents to Generate")
    doc_opts = {
        "Board Resolution": "Board Resolution.docx",
        "Engagement Letter": "Engagement Letter.docx",
        "Management Rep Letter": "Management-representation-letter.docx"
    }
    selected_docs = st.multiselect("Choose files:", list(doc_opts.keys()), default=list(doc_opts.keys()))

    submitted = st.form_submit_button("Generate Documents", type="primary")

if submitted:
    if not selected_docs:
        st.error("Please select at least one document to generate.")
        st.stop()

    full_address = f"{address1}, {address2}, {address3}"
    
    # Base Context (Global variables)
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
        # New explicit keys if user updates templates
        "<<valuation_date>>": val_date.strftime("%d-%b-%Y"),
        "<<br_date>>": br_date.strftime("%d-%b-%Y"),
        "<<engagement_date>>": eng_date.strftime("%d-%b-%Y"),
    }

    template_dir = "templates"
    if not os.path.exists(template_dir):
        st.error(f"Error: '{template_dir}' folder not found.")
        st.stop()

    # Generate Image from Excel (Once)
    img_stream = None
    if valuation_excel:
        with st.spinner("Extracting table from Excel..."):
            img_stream = generate_valuation_table_image(valuation_excel)
            if img_stream is None:
                st.warning("Could not find a valid 'DCF' (Financials sheet) or 'NAV' sheet in the uploaded Excel. <<valuation.jpg>> will be empty.")

    generated_files = []

    try:
        for doc_name in selected_docs:
            filename = doc_opts[doc_name]
            file_path = os.path.join(template_dir, filename)
            
            if not os.path.exists(file_path):
                st.error(f"Template not found: {filename}")
                continue
            
            # --- Dynamic Context Switching for Dates ---
            # This logic ensures the old <<date>> tag gets the RIGHT date based on the document type
            current_context = base_context.copy()
            
            if doc_name == "Management Rep Letter":
                current_context["<<date>>"] = base_context["<<valuation_date>>"]
            elif doc_name == "Board Resolution":
                current_context["<<date>>"] = base_context["<<br_date>>"]
            elif doc_name == "Engagement Letter":
                # Handles potential typo in your template <<date >>
                current_context["<<date>>"] = base_context["<<engagement_date>>"]
                current_context["<<date >>"] = base_context["<<engagement_date>>"]
            else:
                current_context["<<date>>"] = base_context["<<valuation_date>>"] # Default
            
            # Process Document
            doc = process_word_template(file_path, current_context, img_stream)
            
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