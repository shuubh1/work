import streamlit as st
import streamlit_authenticator as stauth
import yaml
from yaml.loader import SafeLoader
import os

def load_config():
    """Loads the YAML config file."""
    config_path = 'config.yaml'
    if not os.path.exists(config_path):
        st.error("Config file not found. Please create config.yaml.")
        st.stop()
        
    with open(config_path) as file:
        config = yaml.load(file, Loader=SafeLoader)
    return config

def require_auth():
    """
    This function should be called at the very top of EVERY page.
    It handles the login widget and stops execution if not logged in.
    """
    config = load_config()

    authenticator = stauth.Authenticate(
        config['credentials'],
        config['cookie']['name'],
        config['cookie']['key'],
        config['cookie']['expiry_days'],
    )

    # Render the login widget
    # Note: We put this in the sidebar usually, or main body
    # For a dedicated login page feel, we put it in main, but once logged in, it disappears.
    
    try:
        # Check authentication state
        authenticator.login()
    except Exception as e:
        st.error(e)

    if st.session_state["authentication_status"]:
        # LOGGED IN SUCCESSFULLY
        with st.sidebar:
            st.write(f'Welcome *{st.session_state["name"]}*')
            authenticator.logout('Logout', 'main')
        return True # Allow the page to run
        
    elif st.session_state["authentication_status"] is False:
        st.error('Username/password is incorrect')
        st.stop() # Stop the page from running
    elif st.session_state["authentication_status"] is None:
        st.warning('Please enter your username and password')
        st.stop() # Stop the page from running