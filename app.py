import streamlit as st
import pandas as pd
import os
import json
import traceback
from convert_orders import convert_shopify_to_singpost

st.set_page_config(
    page_title="Shopify to SingPost Converter",
    page_icon="📦",
    layout="wide"
)

# Add this line right after the set_page_config
st.sidebar.caption("Version 1.4 - Updated Feb 28, 2025")  # Change the version number each time you update

# Setup page
st.title("Shopify to SingPost Converter")
st.write("Upload your Shopify order export CSV to convert it to SingPost ezy2ship format")

# Debug function for secrets
def debug_secrets():
    debug_info = []
    debug_info.append("DEBUGGING SECRETS INFO:")
    
    # Check if secrets exist
    debug_info.append(f"Has secrets object: {hasattr(st, 'secrets')}")
    
    # Check if google_credentials exists in secrets
    if hasattr(st, 'secrets'):
        debug_info.append(f"Keys in secrets: {list(st.secrets.keys())}")
        has_google_creds = 'google_credentials' in st.secrets
        debug_info.append(f"Has google_credentials in secrets: {has_google_creds}")
        
        if has_google_creds:
            try:
                # Check type of credentials
                cred_type = type(st.secrets['google_credentials']).__name__
                debug_info.append(f"Type of google_credentials: {cred_type}")
                
                # Check if it's valid JSON (if it's a string)
                if isinstance(st.secrets['google_credentials'], str):
                    try:
                        json.loads(st.secrets['google_credentials'])
                        debug_info.append("Credentials string is valid JSON")
                    except json.JSONDecodeError:
                        debug_info.append("WARNING: Credentials string is NOT valid JSON")
                
                # Check credential keys if it's a dict
                if isinstance(st.secrets['google_credentials'], dict):
                    credential_keys = list(st.secrets['google_credentials'].keys())
                    debug_info.append(f"Credential keys: {credential_keys}")
                    
                    # Check for required service account keys
                    required_keys = ['type', 'project_id', 'private_key', 'client_email']
                    missing_keys = [key for key in required_keys if key not in credential_keys]
                    if missing_keys:
                        debug_info.append(f"WARNING: Missing required keys: {missing_keys}")
                    else:
                        debug_info.append("All required credential keys present")
            except Exception as e:
                debug_info.append(f"Error inspecting credentials: {str(e)}")
    
    # Check if the session state has credentials info
    debug_info.append(f"Has google_credentials in session_state: {'google_credentials' in st.session_state}")
    debug_info.append(f"Has credentials_path in session_state: {'credentials_path' in st.session_state}")
    
    # Check environment variables
    debug_info.append(f"GOOGLE_CREDENTIALS_PATH env var set: {'GOOGLE_CREDENTIALS_PATH' in os.environ}")
    debug_info.append(f"SLIDES_TEMPLATE_URL env var set: {'SLIDES_TEMPLATE_URL' in os.environ}")
    
    # Check if credentials file exists
    if 'credentials_path' in st.session_state and st.session_state.credentials_path:
        debug_info.append(f"Credentials file exists: {os.path.exists(st.session_state.credentials_path)}")
        
        if os.path.exists(st.session_state.credentials_path):
            # Check file size
            file_size = os.path.getsize(st.session_state.credentials_path)
            debug_info.append(f"Credentials file size: {file_size} bytes")
            
            # Check if it's valid JSON
            try:
                with open(st.session_state.credentials_path, 'r') as f:
                    json.load(f)
                debug_info.append("Credentials file contains valid JSON")
            except json.JSONDecodeError:
                debug_info.append("WARNING: Credentials file does NOT contain valid JSON")
            except Exception as e:
                debug_info.append(f"Error reading credentials file: {str(e)}")
    
    return "\n".join(debug_info)

# Setup page
st.title("Shopify to SingPost Converter")
st.write("Upload your Shopify order export CSV to convert it to SingPost ezy2ship format")

# Debug section - can be commented out in production
with st.expander("Debug Information"):
    st.code(debug_secrets())

# Sidebar for configuration
with st.sidebar:
    st.header("Configuration")
    
    # Google Slides Configuration
    st.subheader("Google Slides Integration")
    
    # Check if credentials exist in session state
    if 'google_credentials' not in st.session_state:
        st.session_state.google_credentials = None
        st.session_state.credentials_path = None
    
# Try to load credentials from secrets
credentials_from_secrets = False
if hasattr(st, 'secrets'):
    # Check for gcp_service_account (the key shown in your error)
    if 'gcp_service_account' in st.secrets:
        try:
            # Get the credentials directly from secrets
            creds = st.secrets['gcp_service_account']
            
            # Create a credentials file manually without trying to serialize the AttrDict
            credentials_path = "google_credentials.json"
            
            with open(credentials_path, "w") as f:
                # Write a JSON object manually with the required fields
                f.write("{\n")
                f.write(f'  "type": "{creds.get("type", "service_account")}",\n')
                f.write(f'  "project_id": "{creds.get("project_id", "")}",\n')
                f.write(f'  "private_key_id": "{creds.get("private_key_id", "")}",\n')
                f.write(f'  "private_key": {json.dumps(creds.get("private_key", ""))},\n')
                f.write(f'  "client_email": "{creds.get("client_email", "")}",\n')
                f.write(f'  "client_id": "{creds.get("client_id", "")}",\n')
                f.write(f'  "auth_uri": "{creds.get("auth_uri", "https://accounts.google.com/o/oauth2/auth")}",\n')
                f.write(f'  "token_uri": "{creds.get("token_uri", "https://oauth2.googleapis.com/token")}",\n')
                f.write(f'  "auth_provider_x509_cert_url": "{creds.get("auth_provider_x509_cert_url", "https://www.googleapis.com/oauth2/v1/certs")}",\n')
                if "client_x509_cert_url" in creds:
                    f.write(f'  "client_x509_cert_url": "{creds.get("client_x509_cert_url")}",\n')
                if "universe_domain" in creds:
                    f.write(f'  "universe_domain": "{creds.get("universe_domain", "googleapis.com")}"\n')
                else:
                    f.write(f'  "universe_domain": "googleapis.com"\n')
                f.write("}")
            
            # Update session state and environment variable
            st.session_state.credentials_path = credentials_path
            os.environ['GOOGLE_CREDENTIALS_PATH'] = credentials_path
            
            st.success("Google credentials loaded from Streamlit secrets!")
            credentials_from_secrets = True
            
            # Debug: Log to console
            print("Successfully loaded Google credentials from secrets!")
            
            # Verify the file was created correctly
            try:
                with open(credentials_path, 'r') as f:
                    print(f"First 50 chars of credentials file: {f.read(50)}...")
                print(f"Credentials file size: {os.path.getsize(credentials_path)} bytes")
            except Exception as e:
                print(f"Error verifying credentials file: {e}")
                
        except Exception as e:
            st.error(f"Error loading credentials from gcp_service_account: {str(e)}")
            st.code(traceback.format_exc())
            print(f"Error loading credentials from gcp_service_account: {e}")
            print(traceback.format_exc())
    
    # If google_credentials not found, try gcp_service_account as fallback
    elif 'gcp_service_account' in st.secrets:
        try:
            # Get the credentials from secrets
            credentials_json = st.secrets['gcp_service_account']
            
            # Handle different formats of credentials
            if isinstance(credentials_json, dict):
                # If it's already a dict, convert to JSON string
                credentials_content = json.dumps(credentials_json)
            else:
                # If it's a string, use it directly
                credentials_content = credentials_json
                # Validate it's proper JSON
                json.loads(credentials_content)
            
            # Store in session state
            st.session_state.google_credentials = credentials_content
            
            # Save to a file
            credentials_path = "google_credentials.json"
            with open(credentials_path, "w") as f:
                f.write(credentials_content)
            
            # Update session state and environment variable
            st.session_state.credentials_path = credentials_path
            os.environ['GOOGLE_CREDENTIALS_PATH'] = credentials_path
            
            st.success("Google credentials loaded from Streamlit secrets (gcp_service_account)!")
            credentials_from_secrets = True
            
            # Debug: Log to console
            print("Successfully loaded Google credentials from gcp_service_account secrets!")
        except Exception as e:
            st.error(f"Error loading credentials from gcp_service_account: {str(e)}")
            st.code(traceback.format_exc())
            print(f"Error loading credentials from gcp_service_account: {e}")
            print(traceback.format_exc())
    
    # Only show upload if not loaded from secrets
    if not credentials_from_secrets:
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
    
    # Try to load template URL from secrets
    template_url_from_secrets = False
    if hasattr(st, 'secrets') and 'google_slides_template' in st.secrets:
        template_url = st.secrets['google_slides_template']
        if template_url and template_url.startswith("https://docs.google.com/presentation"):
            os.environ['SLIDES_TEMPLATE_URL'] = template_url
            st.success("Template URL loaded from Streamlit secrets!")
            template_url_from_secrets = True
            print(f"Successfully loaded template URL from secrets: {template_url}")
    
    # Only show template URL input if not loaded from secrets
    if not template_url_from_secrets:
        # Template URL input
        template_url = st.text_input(
            "Google Slides Template URL (optional)",
            help="URL of a Google Slides template that has been shared with your service account"
        )
        
        if template_url and template_url.startswith("https://docs.google.com/presentation"):
            os.environ['SLIDES_TEMPLATE_URL'] = template_url
            st.success("Template URL saved!")
    
    st.divider()
    
    # Display credentials status
    if 'credentials_path' in st.session_state and st.session_state.credentials_path:
        if os.path.exists(st.session_state.credentials_path):
            st.success("✅ Google credentials are set")
        else:
            st.error("❌ Credentials file not found")
    else:
        st.warning("⚠️ Google credentials not set")
    
    # Display template URL status
    if 'SLIDES_TEMPLATE_URL' in os.environ:
        st.success(f"✅ Template URL is set")
    else:
        st.info("ℹ️ No template URL set (optional)")
    
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
            # Debug: Check environment just before conversion
            print(f"Before conversion - GOOGLE_CREDENTIALS_PATH in env: {'GOOGLE_CREDENTIALS_PATH' in os.environ}")
            print(f"Before conversion - SLIDES_TEMPLATE_URL in env: {'SLIDES_TEMPLATE_URL' in os.environ}")
            if 'GOOGLE_CREDENTIALS_PATH' in os.environ:
                creds_file = os.environ['GOOGLE_CREDENTIALS_PATH']
                print(f"Credentials file exists: {os.path.exists(creds_file)}")
            
            # Run conversion
            try:
                result, pdf_path, slides_url = convert_shopify_to_singpost('orders_export.csv', 'singpost_orders.csv')
                
                # Debug: Print results
                print(f"Conversion result - pdf_path: {pdf_path}, slides_url: {slides_url}")
                
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
                
                col1, col2, col3 = st.columns(3)
                
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
                    if pdf_path and os.path.exists(pdf_path):
                        with open(pdf_path, "rb") as file:
                            st.download_button(
                                label="Download Shipping Labels PDF",
                                data=file,
                                file_name=os.path.basename(pdf_path),
                                mime="application/pdf"
                            )
                    else:
                        st.info("No PDF shipping labels generated")
                
                with col3:
                    if slides_url:
                        st.markdown(f"[Open Google Slides Shipping Labels]({slides_url})")
                    else:
                        if ('credentials_path' in st.session_state and 
                            st.session_state.credentials_path and 
                            os.path.exists(st.session_state.credentials_path)):
                            st.warning("No Google Slides generated. Check service account permissions.")
                        else:
                            st.warning("Google credentials not found. Upload Google credentials to use Slides integration.")
                            
            except Exception as e:
                st.error(f"Error processing orders: {str(e)}")
                st.code(traceback.format_exc())
                print(f"Conversion error: {e}")
                print(traceback.format_exc())