import streamlit as st
from utils.auth_manager import require_auth

st.set_page_config(
    page_title="Firm Internal Tools",
    page_icon="🏢",
    layout="wide"
)

require_auth()

st.title("Firm Internal Tools")
st.markdown("### Centralized Automation Dashboard")
st.divider()

col1, col2, col3 = st.columns(3)

with col1:
    st.info("**Valuations**")
    st.write("DCF Analysis & NAV Calculations.")
    st.page_link("pages/3_Valuations.py", label="Go to Valuations", icon="📊")

with col2:
    st.success("**Document Gen**")
    st.write("Automated report and letter generation.")
    st.page_link("pages/2_Document_Gen.py", label="Go to Generator", icon="📝")

with col3:
    st.success("**Operations**")
    st.write("Bank Statement Consolidation.")
    st.page_link("pages/1_Operations.py", label="Go to Operations", icon="🏦")

st.divider()
st.caption("Select a module from the sidebar or the cards above to begin.")