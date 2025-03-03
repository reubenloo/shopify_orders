import streamlit as st
import pandas as pd
import os
import json
import traceback
import re
from datetime import datetime
from convert_orders import convert_shopify_to_singpost
from google_slides import get_template_id_from_url

# The set_page_config MUST be the first Streamlit command used in your app
st.set_page_config(
    page_title="Shopify to SingPost Converter",
    page_icon="📦",
    layout="wide"
)

# Add version info right after the set_page_config
st.sidebar.caption("Version 1.7 - Updated Mar 1, 2025")

# Define functions
def test_google_api_connection(credentials_path):
    """
    Test Google API connection with provided credentials
    
    Args:
        credentials_path: Path to the service account JSON credentials file
        
    Returns:
        dict: Status of different API connections
    """
    import traceback
    results = {
        "credentials_valid": False,
        "drive_api_working": False,
        "slides_api_working": False,
        "account_email": None,
        "errors": []
    }
    
    try:
        # Verify credentials file exists
        if not os.path.exists(credentials_path):
            results["errors"].append(f"Credentials file not found at {credentials_path}")
            return results
            
        # Set up credentials
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
        
        SCOPES = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/presentations']
        
        try:
            # Validate credentials format
            try:
                with open(credentials_path, 'r') as f:
                    cred_content = f.read()
                    import json
                    cred_json = json.loads(cred_content)
                    required_fields = ['type', 'project_id', 'private_key', 'client_email']
                    missing_fields = [field for field in required_fields if field not in cred_json]
                    if missing_fields:
                        results["errors"].append(f"Credentials missing required fields: {missing_fields}")
                    
                    # Save account email for reference
                    results["account_email"] = cred_json.get('client_email', 'Unknown')
            except Exception as e:
                results["errors"].append(f"Error reading credentials: {str(e)}")
                return results
                
            # Create credentials
            credentials = service_account.Credentials.from_service_account_file(
                credentials_path, scopes=SCOPES)
            
            results["credentials_valid"] = True
            results["account_email"] = credentials.service_account_email
            
            # Test Drive API
            try:
                drive_service = build('drive', 'v3', credentials=credentials)
                # Make a simple API call to verify connection
                files = drive_service.files().list(pageSize=1).execute()
                results["drive_api_working"] = True
            except Exception as e:
                results["errors"].append(f"Drive API Error: {str(e)}")
                
            # Test Slides API
            try:
                slides_service = build('slides', 'v1', credentials=credentials)
                # Create a blank presentation to test
                presentation = {
                    'title': f"API Test - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
                }
                slides_service.presentations().create(body=presentation).execute()
                results["slides_api_working"] = True
            except Exception as e:
                results["errors"].append(f"Slides API Error: {str(e)}")
                
        except Exception as e:
            results["errors"].append(f"Authentication Error: {str(e)}")
            results["errors"].append(traceback.format_exc())
            
    except Exception as e:
        results["errors"].append(f"General Error: {str(e)}")
        results["errors"].append(traceback.format_exc())
        
    return results

def test_create_slides(credentials_path, template_id=None):
    """
    Test creating a simple Google Slides presentation
    
    Args:
        credentials_path: Path to the service account JSON credentials file
        template_id: Optional ID of a template presentation to copy
        
    Returns:
        dict: Results of the test
    """
    import traceback
    results = {
        "success": False,
        "presentation_url": None,
        "errors": [],
        "steps_completed": []
    }
    
    try:
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
        
        # Verify credentials file
        if not os.path.exists(credentials_path):
            results["errors"].append(f"Credentials file not found at {credentials_path}")
            return results
            
        results["steps_completed"].append("Verified credentials file exists")
            
        # Set up credentials
        try:
            SCOPES = ['https://www.googleapis.com/auth/presentations', 
                     'https://www.googleapis.com/auth/drive']
            
            credentials = service_account.Credentials.from_service_account_file(
                credentials_path, scopes=SCOPES)
                
            results["steps_completed"].append("Created credentials object")
        except Exception as e:
            results["errors"].append(f"Error creating credentials: {str(e)}")
            results["errors"].append(traceback.format_exc())
            return results
            
        # Create services
        try:
            slides_service = build('slides', 'v1', credentials=credentials)
            drive_service = build('drive', 'v3', credentials=credentials)
            results["steps_completed"].append("Built API services")
        except Exception as e:
            results["errors"].append(f"Error building services: {str(e)}")
            results["errors"].append(traceback.format_exc())
            return results
            
        # Create presentation
        try:
            presentation_id = None
            
            if template_id:
                try:
                    # Try to use template
                    copy_title = f"Test Presentation - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
                    drive_response = drive_service.files().copy(
                        fileId=template_id,
                        body={"name": copy_title}
                    ).execute()
                    presentation_id = drive_response.get('id')
                    results["steps_completed"].append(f"Copied template: {template_id}")
                except Exception as e:
                    results["errors"].append(f"Error copying template: {str(e)}")
                    # Don't return, fall back to creating new presentation
            
            if not presentation_id:
                # Create new presentation
                presentation = {
                    'title': f"Test Presentation - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
                }
                presentation = slides_service.presentations().create(body=presentation).execute()
                presentation_id = presentation.get('presentationId')
                results["steps_completed"].append("Created new blank presentation")
                
            # Add a text slide
            requests = [
                {
                    'createSlide': {
                        'slideLayoutReference': {
                            'predefinedLayout': 'TITLE_AND_BODY'
                        },
                        'placeholderIdMappings': []
                    }
                }
            ]
            
            response = slides_service.presentations().batchUpdate(
                presentationId=presentation_id,
                body={'requests': requests}
            ).execute()
            
            slide_id = response.get('replies', [{}])[0].get('createSlide', {}).get('objectId')
            
            results["steps_completed"].append("Added a slide to the presentation")
            
            # Get presentation URL
            presentation_url = f"https://docs.google.com/presentation/d/{presentation_id}/edit"
            results["presentation_url"] = presentation_url
            results["steps_completed"].append("Generated presentation URL")
            
            results["success"] = True
            
        except Exception as e:
            results["errors"].append(f"Error creating presentation: {str(e)}")
            results["errors"].append(traceback.format_exc())
            
    except Exception as e:
        results["errors"].append(f"General error: {str(e)}")
        results["errors"].append(traceback.format_exc())
        
    return results

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

def check_template_permissions(credentials_path, template_id):
    """
    Check if the service account has access to the template file
    
    Args:
        credentials_path: Path to the service account JSON credentials file
        template_id: ID of the Google Slides template
        
    Returns:
        bool: True if the service account has access, False otherwise
    """
    if not template_id or not credentials_path:
        print("No template ID or credentials path provided")
        return False
        
    try:
        # Set up credentials
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
        
        SCOPES = ['https://www.googleapis.com/auth/drive']
        
        # Try to create credentials
        try:
            credentials = service_account.Credentials.from_service_account_file(
                credentials_path, scopes=SCOPES)
            print(f"Created credentials for {credentials.service_account_email}")
        except Exception as e:
            print(f"Error creating credentials: {str(e)}")
            return False
            
        # Create drive service
        drive_service = build('drive', 'v3', credentials=credentials)
        
        # Try to get file metadata
        try:
            file = drive_service.files().get(fileId=template_id).execute()
            print(f"Successfully accessed template file: {file.get('name')}")
            return True
        except Exception as e:
            print(f"Error accessing template file: {str(e)}")
            # If the error is about permissions, provide specific guidance
            error_str = str(e).lower()
            if 'permission' in error_str or 'access' in error_str or 'not found' in error_str:
                print("This is likely a permissions issue. Make sure the template has been shared with the service account email.")
                print(f"Service account email: {credentials.service_account_email}")
            return False
            
    except Exception as e:
        print(f"Error checking template permissions: {str(e)}")
        return False

# Setup page - Title and Description
st.title("Shopify to SingPost Converter")
st.write("Upload your Shopify order export CSV to convert it to SingPost ezy2ship format")

# Debug section - can be commented out in production
with st.expander("Debug Information"):
    st.code(debug_secrets())

# API Connection Tester
with st.expander("API Connection Tester"):
    st.write("Test Google API connection with current credentials")
    if st.button("Test Google API Connection"):
        if 'credentials_path' in st.session_state and os.path.exists(st.session_state.credentials_path):
            with st.spinner("Testing API connection..."):
                results = test_google_api_connection(st.session_state.credentials_path)
                
                # Display results
                if results["credentials_valid"]:
                    st.success("✅ Credentials format is valid")
                else:
                    st.error("❌ Credentials format is invalid")
                    
                if results["drive_api_working"]:
                    st.success("✅ Google Drive API is working")
                else:
                    st.error("❌ Google Drive API is not working")
                    
                if results["slides_api_working"]:
                    st.success("✅ Google Slides API is working")
                else:
                    st.error("❌ Google Slides API is not working")
                
                st.markdown(f"**Service Account Email:** {results['account_email']}")
                
                if results["errors"]:
                    st.error("Errors encountered:")
                    for error in results["errors"]:
                        st.code(error)
                        
                # Provide suggestions based on results
                if not results["credentials_valid"]:
                    st.info("💡 Check that your service account JSON file has the correct format")
                    
                if not results["drive_api_working"] or not results["slides_api_working"]:
                    st.info("💡 Possible causes:")
                    st.markdown("""
                    - The Google Cloud project might not have the necessary APIs enabled
                    - The service account might not have the required permissions
                    - The private key might be incorrect
                    """)
                    st.markdown(f"""
                    **Steps to fix:**
                    1. Go to [Google Cloud Console](https://console.cloud.google.com/)
                    2. Make sure the Google Drive API and Google Slides API are enabled
                    3. Check that the service account has the necessary permissions
                    4. Try downloading a new service account JSON key
                    """)
        else:
            st.error("No credentials file found. Please upload or load credentials first.")

# Google Slides Tester
with st.expander("Google Slides Tester"):
    st.write("Test creating a Google Slides presentation")
    
    col1, col2 = st.columns([3, 1])
    
    with col1:
        use_template = st.checkbox("Use template", value=False)
        if use_template:
            template_url = st.text_input(
                "Template URL", 
                value=os.environ.get('SLIDES_TEMPLATE_URL', ''),
                help="URL of a Google Slides template that has been shared with your service account"
            )
        else:
            template_url = None
    
    with col2:
        if st.button("Create Test Slides"):
            if 'credentials_path' in st.session_state and os.path.exists(st.session_state.credentials_path):
                with st.spinner("Creating test presentation..."):
                    # Get template ID from URL if provided
                    template_id = None
                    if template_url:
                        match = re.search(r'/d/([a-zA-Z0-9-_]+)', template_url)
                        if match:
                            template_id = match.group(1)
                    
                    # Run the test
                    results = test_create_slides(st.session_state.credentials_path, template_id)
                    
                    # Display results
                    if results["success"]:
                        st.success("✅ Successfully created presentation!")
                        st.markdown(f"[Open presentation]({results['presentation_url']})")
                    else:
                        st.error("❌ Failed to create presentation")
                    
                    st.subheader("Steps Completed")
                    for i, step in enumerate(results["steps_completed"]):
                        st.write(f"{i+1}. {step}")
                    
                    if results["errors"]:
                        st.subheader("Errors")
                        for error in results["errors"]:
                            st.code(error)
            else:
                st.error("No credentials file found. Please upload or load credentials first.")

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
        # Check for gcp_service_account key
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
            st.info("Starting conversion process...")
            
            debug_msg = []
            debug_msg.append(f"GOOGLE_CREDENTIALS_PATH in env: {'GOOGLE_CREDENTIALS_PATH' in os.environ}")
            if 'GOOGLE_CREDENTIALS_PATH' in os.environ:
                creds_file = os.environ['GOOGLE_CREDENTIALS_PATH']
                debug_msg.append(f"Credentials file path: {creds_file}")
                debug_msg.append(f"Credentials file exists: {os.path.exists(creds_file)}")
                if os.path.exists(creds_file):
                    debug_msg.append(f"Credentials file size: {os.path.getsize(creds_file)} bytes")
                    
            debug_msg.append(f"SLIDES_TEMPLATE_URL in env: {'SLIDES_TEMPLATE_URL' in os.environ}")
            if 'SLIDES_TEMPLATE_URL' in os.environ:
                debug_msg.append(f"Template URL: {os.environ['SLIDES_TEMPLATE_URL']}")
                
            # Debug credentials in session state
            debug_msg.append(f"google_credentials in session_state: {'google_credentials' in st.session_state}")
            debug_msg.append(f"credentials_path in session_state: {'credentials_path' in st.session_state}")
            if 'credentials_path' in st.session_state:
                debug_msg.append(f"Session state path: {st.session_state.credentials_path}")
                debug_msg.append(f"Path exists: {os.path.exists(st.session_state.credentials_path)}")
                
            # Log to console
            print("\n".join(debug_msg))
            
            # Create an expander for debugging info
            with st.expander("Pre-Conversion Debug Info"):
                st.code("\n".join(debug_msg))
            
            # Check template permissions if URL is set
            if 'SLIDES_TEMPLATE_URL' in os.environ and 'credentials_path' in st.session_state:
                template_url = os.environ['SLIDES_TEMPLATE_URL']
                template_id = None
                if template_url:
                    match = re.search(r'/d/([a-zA-Z0-9-_]+)', template_url)
                    if match:
                        template_id = match.group(1)
                        
                if template_id:
                    has_access = check_template_permissions(st.session_state.credentials_path, template_id)
                    if not has_access:
                        st.warning("⚠️ The service account does not have access to the template. Please check sharing permissions.")
            
            # Run conversion
            # Update this part of app.py where the conversion is happening
# Replace this block in your app.py file

# Run conversion
try:
    result, pdf_path, slides_url, debug_log = convert_shopify_to_singpost('orders_export.csv', 'singpost_orders.csv')
    
    # Debug: Print results
    print(f"Conversion result - pdf_path: {pdf_path}, slides_url: {slides_url}")
    
    # Create an expander for post-conversion debugging
    with st.expander("Post-Conversion Debug Info"):
        st.code(f"PDF Path: {pdf_path}\nSlides URL: {slides_url}")
    
    # Display results
    st.success("Conversion completed!")
    
    # Create tabs for summary, data preview, and debug log
    tab1, tab2, tab3 = st.tabs(["Summary", "Data Preview", "Google Slides Debug Log"])
    
    with tab1:
        st.text_area("Conversion Summary", result, height=400)
    
    with tab2:
        if os.path.exists("singpost_orders.csv"):
            df = pd.read_csv("singpost_orders.csv")
            st.dataframe(df)
        else:
            st.info("No SingPost CSV was generated (no international orders)")
    
    with tab3:
        if debug_log:
            st.text_area("Debug Log", debug_log, height=600)
        else:
            st.info("No debug information available")
    
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
    with st.expander("Detailed Error Information"):
        st.code(traceback.format_exc())
    print(f"Conversion error: {e}")
    print(traceback.format_exc())