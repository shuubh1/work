import streamlit as st
import pandas as pd
import io
from utils.auth_manager import require_auth
# Import logic from our new utils folder
from utils.bank_parsers import (
    parse_bank_of_america, 
    parse_td_visa_card, 
    parse_td_generic
)

st.set_page_config(page_title="Operations", page_icon="🏦", layout="wide")

require_auth()

st.title("Operations & Reconciliation")
st.markdown("### Bank Statement Consolidator")

# Initialize session state for file selections
if 'file_selections' not in st.session_state:
    st.session_state.file_selections = {}

# Initialize session state for the processed dataframe
if 'processed_data' not in st.session_state:
    st.session_state.processed_data = None

out_name = st.text_input("Output Filename", "consolidated_summary.xlsx")
uploaded_files = st.file_uploader("Upload PDFs", type="pdf", accept_multiple_files=True)

bank_opts = [
    "Select...", 
    "Bank of America", 
    "TD Business Convenience Plus", 
    "TD BUSINESS SOLUTIONS VISA", 
    "TD Small Business Premium Money Mar"
]

if uploaded_files:
    st.divider()
    st.subheader("Assign Banks")
    for f in uploaded_files:
        col_a, col_b = st.columns([2,3])
        with col_a: st.write(f"📄 {f.name}")
        with col_b:
            # Use f.name as key
            if f.name not in st.session_state.file_selections:
                st.session_state.file_selections[f.name] = "Select..."
            
            st.session_state.file_selections[f.name] = st.selectbox(
                "Bank Type", 
                bank_opts, 
                key=f.name, 
                label_visibility="collapsed"
            )
    
    # Process Button
    if st.button("Process Files", type="primary"):
        # Clear previous results to avoid confusion
        st.session_state.processed_data = None
        
        # Validation
        if any(st.session_state.file_selections[f.name] == "Select..." for f in uploaded_files):
            st.error("Please select a bank type for all files.")
        else:
            all_txns = []
            with st.spinner("Processing..."):
                for f in uploaded_files:
                    b_type = st.session_state.file_selections[f.name]
                    f.seek(0) 
                    
                    if b_type == "Bank of America": 
                        all_txns.extend(parse_bank_of_america(f))
                    elif b_type == "TD BUSINESS SOLUTIONS VISA": 
                        all_txns.extend(parse_td_visa_card(f))
                    elif b_type == "TD Small Business Premium Money Mar": 
                        all_txns.extend(parse_td_generic(f, b_type, ["Other Credits"], ["Electronic Payments", "Other Withdrawals"]))
                    elif b_type == "TD Business Convenience Plus": 
                        all_txns.extend(parse_td_generic(f, b_type, ["Electronic Deposits"], ["Electronic Payments"]))

            if all_txns:
                st.success(f"Success! Extracted {len(all_txns)} transactions.")
                df = pd.DataFrame(all_txns, columns=['Bank', 'Date', 'Ref', 'Description', 'Amount'])
                df['Date'] = pd.to_datetime(df['Date'])
                df = df.sort_values(by='Date').reset_index(drop=True)
                df['Date'] = df['Date'].dt.strftime('%Y-%m-%d')
                
                # SAVE TO SESSION STATE instead of creating button immediately
                st.session_state.processed_data = df
            else:
                st.warning("No transactions found.")

    # Show Download Button OUTSIDE the process button block
    # This checks if data exists in memory and shows the button persistently
    if st.session_state.processed_data is not None:
        st.divider()
        st.write("### Download Results")
        
        buffer = io.BytesIO()
        st.session_state.processed_data.to_excel(buffer, index=False)
        buffer.seek(0)
        
        st.download_button(
            label="Download Excel File",
            data=buffer,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )