import streamlit as st
import pandas as pd
import os
from convert_orders import convert_shopify_to_singpost
import json

st.set_page_config(
    page_title="Shopify to SingPost Converter",
    page_icon="📦",
    layout="wide"
)

# Setup page
st.title("Shopify to SingPost Converter")
st.write("Upload your Shopify order export CSV to convert it to SingPost ezy2ship format")

# Sidebar for configuration
with st.sidebar:
    st.header("Configuration")
    
    # Google Slides Configuration
    st.subheader("Google Slides Integration")
    
    # Check if credentials exist in session state
    if 'google_credentials' not in st.session_state:
        st.session_state.google_credentials = None
        st.session_state.credentials_path = None
    
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
    
    # Template URL input
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
    2. The app generates SingPost CSV for international orders
    3. For Singapore orders, it creates shipping labels in Google Slides
    """)

# Main content area
# File uploader for Shopify CSV
uploaded_file = st.file_uploader("Upload Shopify orders_export.csv", type='csv')

if uploaded_file:
    # Save the uploaded file
    with open("orders_export.csv", "wb") as f:
        f.write(uploaded_file.getvalue())
    
    # Add a button to run the conversion
    if st.button("Convert to SingPost Format"):
        with st.spinner("Processing orders..."):
            # Run conversion
            try:
                result, pdf_path, slides_url = convert_shopify_to_singpost('orders_export.csv', 'singpost_orders.csv')
                
                # Display results
                st.success("Conversion completed!")
                
                # Create tabs for summary and data preview
                tab1, tab2 = st.tabs(["Summary", "Data Preview"])
                
                with tab1:
                    st.text_area("Conversion Summary", result, height=400)
                
                with tab2:
                    if os.path.exists("singpost_orders.csv"):
                        df = pd.read_csv("singpost_orders.csv")
                        st.dataframe(df)
                    else:
                        st.info("No SingPost CSV was generated (no international orders)")
                
                # Download section
                st.divider()
                st.subheader("Download Results")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    if os.path.exists("singpost_orders.csv"):
                        with open("singpost_orders.csv", "rb") as file:
                            st.download_button(
                                label="Download SingPost CSV",
                                data=file,
                                file_name="singpost_orders.csv",
                                mime="text/csv"
                            )
                    else:
                        st.info("No SingPost CSV generated")
                
                with col2:
                    if slides_url:
                        st.markdown(f"[Open Google Slides Shipping Labels]({slides_url})")
                    else:
                        if st.session_state.credentials_path:
                            st.info("No Google Slides generated. Check permissions.")
                        else:
                            st.warning("Upload Google credentials to use Slides integration")
                            
            except Exception as e:
                st.error(f"Error processing orders: {str(e)}")
                st.exception(e)