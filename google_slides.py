import os
from googleapiclient.discovery import build
from google.oauth2 import service_account
from datetime import datetime, timedelta
import re
import pytz

def create_shipping_slides(order_details, credentials_path, template_id=None):
    """
    Create or update a Google Slides presentation with shipping labels for orders
    
    Args:
        order_details: List of dictionaries containing order information
        credentials_path: Path to the service account JSON credentials file
        template_id: Optional ID of a template presentation to copy
        
    Returns:
        tuple: (presentation_url, pdf_path) URLs of the created presentation and path to PDF
    """
    try:
        # Verify credentials file exists
        if not os.path.exists(credentials_path):
            print(f"Error: Credentials file not found at {credentials_path}")
            return None, None
            
        # Print debug info
        print(f"Creating shipping slides with credentials: {credentials_path}")
        print(f"File exists: {os.path.exists(credentials_path)}")
        print(f"Template ID: {template_id}")
        
        try:
            # Validate credentials by loading the file
            with open(credentials_path, 'r') as f:
                cred_content = f.read()
                # Try parsing as JSON to verify format
                json.loads(cred_content)
                print("Credentials file loaded and validated as JSON")
        except Exception as e:
            print(f"Error validating credentials file: {str(e)}")
            return None, None
            
        # Set up credentials
        SCOPES = ['https://www.googleapis.com/auth/presentations', 
                  'https://www.googleapis.com/auth/drive']
        
        # Try to create credentials
        try:
            credentials = service_account.Credentials.from_service_account_file(
                credentials_path, scopes=SCOPES)
            print("Successfully created credentials object")
        except Exception as e:
            print(f"Error creating credentials: {str(e)}")
            return None, None
            
        # Create services
        try:
            slides_service = build('slides', 'v1', credentials=credentials)
            drive_service = build('drive', 'v3', credentials=credentials)
            print("Successfully built Google API services")
        except Exception as e:
            print(f"Error building Google API services: {str(e)}")
            return None, None
        
        # Create a new presentation or use template
        presentation_id = None
        presentation_url = None
        
        try:
            if template_id:
                # Copy the template presentation
                copy_title = f"Shipping Labels - {datetime.now().strftime('%Y-%m-%d %H:%M')}"
                drive_response = drive_service.files().copy(
                    fileId=template_id,
                    body={"name": copy_title}
                ).execute()
                presentation_id = drive_response.get('id')
                print(f"Successfully copied template presentation: {template_id}")
            else:
                # Create a new blank presentation
                presentation = {
                    'title': f"Shipping Labels - {datetime.now().strftime('%Y-%m-%d %H:%M')}"
                }
                presentation = slides_service.presentations().create(body=presentation).execute()
                presentation_id = presentation.get('presentationId')
                print("Successfully created new blank presentation")
                
            presentation_url = f"https://docs.google.com/presentation/d/{presentation_id}/edit"
        except Exception as e:
            print(f"Error creating/copying presentation: {str(e)}")
            return None, None
        
        # The rest of the function remains the same...
        # ... (include all the code for deleting slides, creating date slide, etc.)
        
        return presentation_url, None
        
    except Exception as e:
        print(f"Error creating Google Slides: {str(e)}")
        import traceback
        traceback.print_exc()
        return None, None

def get_template_id_from_url(url):
    """Extract the presentation ID from a Google Slides URL"""
    match = re.search(r'/d/([a-zA-Z0-9-_]+)', url)
    if match:
        return match.group(1)
    return None