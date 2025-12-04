import streamlit as st
import io
import os
import zipfile
from utils.auth_manager import require_auth
# Import logic from utils
from utils.doc_utils import process_word_template

st.set_page_config(page_title="Document Gen", page_icon="📝", layout="wide")

require_auth()

st.title("Document Generation Suite")
st.markdown("Fill in the details below to generate standard firm documents.")

# --- Input Form ---
with st.form("doc_gen_form"):
    st.subheader("1. General Information")
    col1, col2 = st.columns(2)
    with col1:
        input_date = st.date_input("Date")
        company = st.text_input("Company Name", "My Client Company Pvt Ltd")
        valuation_statement = st.text_input("Statement for Valuation", "Preferential Allotment of Equity Shares")
    with col2:
        authority = st.text_input("Addressee Name (Authority)", "Mr. John Doe")
        designation = st.text_input("Addressee Designation", "Director")
        addressed_to = st.text_input("Salutation (e.g. Sir/Ma'am)", "Sir")

    st.subheader("2. Address Details")
    address1 = st.text_input("Address Line 1", "123 Business Park")
    col3, col4 = st.columns(2)
    with col3:
        address2 = st.text_input("Address Line 2", "Financial District")
    with col4:
        address3 = st.text_input("Address Line 3", "New York, NY 10001")
        
    st.subheader("3. Specifics")
    group_designation = st.text_input("Group Designation (e.g. Board of Directors)", "Board of Directors")
    valuation_image = st.file_uploader("Upload Valuation Image/Graph (for <<valuation.jpg>>)", type=['png', 'jpg', 'jpeg'])

    st.subheader("4. Select Documents to Generate")
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

    # Prepare Context Dictionary
    full_address = f"{address1}, {address2}, {address3}"
    
    context = {
        "<<date>>": input_date.strftime("%d-%b-%Y"),
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
        "<<authority_designation>>": designation 
    }

    # Verify Templates Exist
    template_dir = "templates"
    if not os.path.exists(template_dir):
        st.error(f"Error: '{template_dir}' folder not found. Please create it and add your .docx files.")
        st.stop()

    generated_files = []

    try:
        for doc_name in selected_docs:
            filename = doc_opts[doc_name]
            file_path = os.path.join(template_dir, filename)
            
            if not os.path.exists(file_path):
                st.error(f"Template not found: {filename}")
                continue
            
            # Handle Image Stream
            img_stream = None
            if valuation_image:
                img_stream = io.BytesIO(valuation_image.getvalue())

            # Process using utility function
            doc = process_word_template(file_path, context, img_stream)
            
            # Save to memory
            doc_io = io.BytesIO()
            doc.save(doc_io)
            doc_io.seek(0)
            generated_files.append((f"Generated_{filename}", doc_io))

        # Output Logic
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