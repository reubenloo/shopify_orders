import streamlit as st
import pandas as pd
import os
from convert_orders import convert_shopify_to_singpost
import json
import tempfile

st.set_page_config(
    page_title="Shopify to SingPost Converter",
    page_icon="📦",
    layout="wide"
)

# Setup page
st.title("Shopify to SingPost Converter")
st.write("Upload your Shopify order export CSV to convert it to SingPost ezy2ship format")
st.caption("v2.2.0 - retail-mail-international template with IRREPK service code")

# Function to handle credentials from secrets
def setup_credentials_from_secrets():
    credentials_loaded = False

    # Load Google credentials
    if 'google_credentials' in st.secrets:
        # Create a temporary file for the credentials
        credentials_path = "google_credentials.json"

        # Extract the credentials from secrets and write to file
        creds_dict = dict(st.secrets["google_credentials"])

        # Write to a temporary file
        with open(credentials_path, "w") as f:
            json.dump(creds_dict, f)

        os.environ['GOOGLE_CREDENTIALS_PATH'] = credentials_path

        # If template URL is in secrets, use it
        if 'google_slides_template' in st.secrets and 'url' in st.secrets['google_slides_template']:
            template_url = st.secrets['google_slides_template']['url']
            os.environ['SLIDES_TEMPLATE_URL'] = template_url
            st.sidebar.success(f"Using template from secrets: {template_url}")

        credentials_loaded = True

    # Load Shopify credentials
    if 'shopify' in st.secrets:
        os.environ['SHOPIFY_ACCESS_TOKEN'] = st.secrets['shopify']['access_token']
        os.environ['SHOPIFY_STORE_URL'] = st.secrets['shopify']['store_url']
        st.sidebar.success("✓ Shopify API credentials loaded")
        credentials_loaded = True
    else:
        st.sidebar.warning("⚠️ Shopify API credentials not found - US orders will fail")

    return credentials_loaded

# Check for credentials in secrets first
using_secrets = setup_credentials_from_secrets()

# Sidebar for configuration
with st.sidebar:
    st.header("Configuration")
    
    # Google Slides Configuration
    st.subheader("Google Slides Integration")
    
    # Check if credentials exist in session state
    if 'google_credentials' not in st.session_state:
        st.session_state.google_credentials = None
        st.session_state.credentials_path = None
    
    if using_secrets:
        st.success("Google credentials loaded successfully from Streamlit Secrets!")
    else:
        # Upload Google credentials
        credentials_file = st.file_uploader(
            "Upload Google service account credentials (JSON)", 
            type=['json'],
            help="Upload the service account JSON file from Google Cloud console"
        )
        
        if credentials_file is not None:
            # Save the credentials to a temporary file
            credentials_content = credentials_file.getvalue().decode('utf-8')
            try:
                # Validate it's a proper JSON
                json.loads(credentials_content)
                st.session_state.google_credentials = credentials_content
                
                # Save to a file
                credentials_path = "google_credentials.json"
                with open(credentials_path, "w") as f:
                    f.write(credentials_content)
                
                st.session_state.credentials_path = credentials_path
                os.environ['GOOGLE_CREDENTIALS_PATH'] = credentials_path
                
                st.success("Google credentials loaded successfully!")
            except json.JSONDecodeError:
                st.error("Invalid JSON file. Please upload a valid service account credentials file.")
    
    # Template URL input (only if not in secrets)
    if not (using_secrets and 'google_slides_template' in st.secrets):
        template_url = st.text_input(
            "Google Slides Template URL (optional)",
            help="URL of a Google Slides template that has been shared with your service account"
        )
        
        if template_url and template_url.startswith("https://docs.google.com/presentation"):
            os.environ['SLIDES_TEMPLATE_URL'] = template_url
            st.success("Template URL saved!")
    
    st.divider()
    
    # Show information
    st.info("""
    **How this works:**
    1. Upload your Shopify export CSV
    2. The app generates TWO SingPost CSV files:
       - International orders (ex-SG, ex-US, ex-CA) in SGD
       - US orders in USD with US-specific pricing
    3. For Singapore orders, it creates shipping labels in Google Slides
    """)

# Initialize session state for conversion results
if 'conversion_results' not in st.session_state:
    st.session_state.conversion_results = None
if 'last_uploaded_file' not in st.session_state:
    st.session_state.last_uploaded_file = None

# Main content area
# File uploader for Shopify CSV
uploaded_file = st.file_uploader("Upload Shopify orders_export.csv", type='csv')

if uploaded_file:
    # Check if a new file was uploaded - clear old results
    if st.session_state.last_uploaded_file != uploaded_file.name:
        st.session_state.conversion_results = None
        st.session_state.last_uploaded_file = uploaded_file.name

    # Save the uploaded file
    with open("orders_export.csv", "wb") as f:
        f.write(uploaded_file.getvalue())

    # Add a button to run the conversion
    if st.button("Convert to SingPost Format"):
        with st.spinner("Processing orders..."):
            # Run conversion
            try:
                result, intl_df, us_df, slides_url = convert_shopify_to_singpost('orders_export.csv', 'singpost_orders.csv')

                # Store results in session state
                st.session_state.conversion_results = {
                    'summary': result,
                    'intl_df': intl_df,
                    'us_df': us_df,
                    'slides_url': slides_url
                }

                st.success("Conversion completed!")

            except Exception as e:
                st.error(f"Error processing orders: {str(e)}")
                st.exception(e)
                st.session_state.conversion_results = None

# Display results if they exist in session state
if st.session_state.conversion_results is not None:
    results = st.session_state.conversion_results

    # Create tabs for summary and data preview
    tab1, tab2, tab3 = st.tabs(["Summary", "International Orders Preview", "US Orders Preview"])

    with tab1:
        st.text_area("Conversion Summary", results['summary'], height=400)

    with tab2:
        if results['intl_df'] is not None:
            st.subheader("International Orders (ex-SG, ex-US, ex-CA)")
            st.dataframe(results['intl_df'])
        else:
            st.info("No international orders (ex-SG, ex-US, ex-CA)")

    with tab3:
        if results['us_df'] is not None:
            st.subheader("US Orders")
            st.dataframe(results['us_df'])
        else:
            st.info("No US orders")

    # Download section
    st.divider()
    st.subheader("Download Results")

    col1, col2, col3 = st.columns(3)

    with col1:
        if os.path.exists("singpost_orders.csv"):
            with open("singpost_orders.csv", "rb") as file:
                st.download_button(
                    label="📥 Download International CSV",
                    data=file,
                    file_name="singpost_international_orders.csv",
                    mime="text/csv",
                    key="download_intl"
                )
        else:
            st.info("No international CSV")

    with col2:
        if os.path.exists("singpost_orders_us.csv"):
            with open("singpost_orders_us.csv", "rb") as file:
                st.download_button(
                    label="📥 Download US CSV",
                    data=file,
                    file_name="singpost_us_orders.csv",
                    mime="text/csv",
                    key="download_us"
                )
        else:
            st.info("No US CSV")

    with col3:
        if results['slides_url']:
            st.success("Google Slides labels created!")
            st.markdown(f"[Open Shipping Labels]({results['slides_url']})")
        else:
            if os.environ.get('GOOGLE_CREDENTIALS_PATH') and os.path.exists(os.environ.get('GOOGLE_CREDENTIALS_PATH')):
                if os.environ.get('SLIDES_TEMPLATE_URL'):
                    st.info("No Singapore orders or error generating slides")
                else:
                    st.info("No Slides template URL")
            else:
                st.info("No Google credentials")