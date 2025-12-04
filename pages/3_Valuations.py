import streamlit as st
import io
import os
import pandas as pd
from utils.valuation_utils import generate_nav_table_image, generate_valuation_report

st.set_page_config(page_title="Valuations", page_icon="📊", layout="wide")

st.title("Valuations Dashboard")

tab1, tab2 = st.tabs(["DCF Analysis", "NAV Calculation"])

# --- DCF TAB (Placeholder) ---
with tab1:
    st.header("Discounted Cash Flow (DCF)")
    st.info("This module is currently under development.")

# --- NAV TAB ---
with tab2:
    st.header("Net Asset Value (NAV) Generation")
    st.markdown("Generate a standard NAV Valuation Report by combining **Excel workings** with **Report Metadata**.")

    col_layout_1, col_layout_2 = st.columns([1, 2])

    # 1. Excel Upload (The Workings)
    with col_layout_1:
        st.subheader("1. Upload Workings")
        uploaded_excel = st.file_uploader(
            "Upload 'nav_valuation_workings_template.xlsx'", 
            type=["xlsx", "xls"]
        )
        
        if uploaded_excel:
            st.success("Excel Loaded!")
            # Optional: Preview the data
            try:
                preview_df = pd.read_excel(uploaded_excel, sheet_name='NAV Calculation Working')
                with st.expander("Preview NAV Data"):
                    st.dataframe(preview_df.head())
            except Exception as e:
                st.error(f"Could not read 'NAV Calculation Working' sheet. Check file format.")

    # 2. Input Fields (The Report Details)
    with col_layout_2:
        st.subheader("2. Report Details")
        with st.form("nav_form"):
            col_a, col_b = st.columns(2)
            
            with col_a:
                valuation_date = st.date_input("Valuation Date")
                company_name = st.text_input("Target Company Name", "My Client Company Pvt Ltd")
                directed_to = st.text_input("Directed To", "The Board of Directors")
            
            with col_b:
                appointing_company = st.text_input("Appointing Company/Person", "Shareholder's Firm")
                appointing_address = st.text_input("Appointing Co. Address (Short)", "Kolkata, West Bengal")
            
            st.markdown("**Full Address (for Header/Footer):**")
            appointing_3line = st.text_area("Appointing Co. Full Address (Multi-line)", "123 Street Name,\nDistrict,\nCity - 700001")

            submit_nav = st.form_submit_button("Generate NAV Report", type="primary")

    # 3. Processing Logic
    if submit_nav:
        if not uploaded_excel:
            st.error("Please upload the Excel workings file first.")
            st.stop()
        
        # Verify template exists
        template_path = os.path.join("templates", "valuation_report_template.docx")
        if not os.path.exists(template_path):
            st.error(f"Template not found at: {template_path}")
            st.stop()

        with st.spinner("Analyzing Excel and generating report..."):
            try:
                # A. Generate the Image from Excel
                # We need to reset the file pointer because we might have read it for preview
                uploaded_excel.seek(0)
                nav_image_buffer = generate_nav_table_image(uploaded_excel)
                
                # B. Prepare the Context (Map inputs to {{placeholders}})
                context = {
                    'valuation_date': valuation_date.strftime("%d-%b-%Y"),
                    'company': company_name,
                    'directed_to': directed_to,
                    'appointing_company': appointing_company,
                    'appointing_company_address': appointing_address,
                    'appointing_company_3line_address': appointing_3line,
                    # 'nav.jpg' is handled inside the utility function
                }

                # C. Generate the DOCX
                final_report_io = generate_valuation_report(template_path, context, nav_image_buffer)
                
                # D. Success & Download
                st.success("Report Generated Successfully!")
                
                st.download_button(
                    label="Download NAV Report (.docx)",
                    data=final_report_io,
                    file_name=f"NAV_Report_{company_name.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

            except Exception as e:
                st.error(f"An error occurred: {str(e)}")