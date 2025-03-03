import os
import json
from googleapiclient.discovery import build
from google.oauth2 import service_account
from datetime import datetime
import re
import time

def create_shipping_slides(order_details, credentials_path, template_id=None):
    """
    Edit an existing Google Slides presentation with shipping labels for orders
    
    Args:
        order_details: List of dictionaries containing order information
        credentials_path: Path to the service account JSON credentials file
        template_id: ID of the existing presentation to edit
        
    Returns:
        tuple: (presentation_url, pdf_path) URL of the edited presentation and path to PDF
    """
    try:
        # Debug info
        print(f"Starting create_shipping_slides with {len(order_details)} orders")
        print(f"Credentials path: {credentials_path}")
        print(f"File exists: {os.path.exists(credentials_path)}")
        print(f"Template ID: {template_id or 'None'}")
        
        # Validate credentials file
        try:
            with open(credentials_path, 'r') as f:
                cred_content = f.read()
                masked_content = cred_content[:100].replace('"private_key":', '"private_key": "[MASKED]",')
                print(f"Credentials file content (first 100 chars): {masked_content}...")
                
                try:
                    cred_json = json.loads(cred_content)
                    required_fields = ['type', 'project_id', 'private_key', 'client_email']
                    missing_fields = [field for field in required_fields if field not in cred_json]
                    if missing_fields:
                        print(f"WARNING: Credentials file is missing required fields: {missing_fields}")
                    else:
                        print("Credentials file contains all required fields")
                except json.JSONDecodeError as e:
                    print(f"ERROR: Credentials file is not valid JSON: {str(e)}")
                    return None, None
        except Exception as e:
            print(f"ERROR: Could not read credentials file: {str(e)}")
            return None, None
        
        # Set up credentials
        SCOPES = ['https://www.googleapis.com/auth/presentations',
                  'https://www.googleapis.com/auth/drive']
        
        try:
            print("Creating credentials object...")
            credentials = service_account.Credentials.from_service_account_file(
                credentials_path, scopes=SCOPES)
            print(f"Credentials created successfully: {credentials.service_account_email}")
        except Exception as e:
            print(f"ERROR creating credentials: {str(e)}")
            return None, None
        
        # Create Slides/Drive services
        try:
            print("Building slides service...")
            slides_service = build('slides', 'v1', credentials=credentials)
            print("Building drive service...")
            drive_service = build('drive', 'v3', credentials=credentials)
            print("Services built successfully")
        except Exception as e:
            print(f"ERROR building services: {str(e)}")
            return None, None
        
        # Check for template_id
        if not template_id:
            print("No template ID provided, cannot proceed")
            return None, None
        
        presentation_id = template_id
        presentation_url = f"https://docs.google.com/presentation/d/{presentation_id}/edit"
        print(f"Presentation URL: {presentation_url}")
        
        # Retrieve the presentation to see how many slides exist
        presentation = slides_service.presentations().get(
            presentationId=presentation_id
        ).execute()
        
        slides = presentation.get('slides', [])
        if len(slides) < 2:
            print("ERROR: Presentation must have at least 2 slides (date and template)")
            return presentation_url, None
        
        # -----------------------------
        # 1) Duplicate the first slide (date slide)
        # -----------------------------
        print("Duplicating the first slide for the date page...")
        date_request = [{
            'duplicateObject': {
                'objectId': slides[0].get('objectId'),
            }
        }]
        
        date_response = slides_service.presentations().batchUpdate(
            presentationId=presentation_id,
            body={'requests': date_request}
        ).execute()
        
        new_date_slide_id = date_response.get('replies', [{}])[0].get('duplicateObject', {}).get('objectId')
        print(f"Created new date slide with ID: {new_date_slide_id}")
        
        # -----------------------------
        # 2) Update the date on the new slide
        # -----------------------------
        # Get fresh slide data
        new_slides_data = slides_service.presentations().get(
            presentationId=presentation_id,
            fields="slides"
        ).execute()
        
        new_date_slide = None
        for slide in new_slides_data.get('slides', []):
            if slide.get('objectId') == new_date_slide_id:
                new_date_slide = slide
                break
        
        if new_date_slide:
            # Insert today's date in the first text shape we find
            for element in new_date_slide.get('pageElements', []):
                if 'shape' in element and 'text' in element['shape']:
                    element_id = element['objectId']
                    today = datetime.now().strftime("%B %d, %Y")
                    
                    date_update_request = [
                        {
                            'deleteText': {
                                'objectId': element_id,
                                'textRange': {'type': 'ALL'}
                            }
                        },
                        {
                            'insertText': {
                                'objectId': element_id,
                                'insertionIndex': 0,
                                'text': today
                            }
                        }
                    ]
                    
                    slides_service.presentations().batchUpdate(
                        presentationId=presentation_id,
                        body={'requests': date_update_request}
                    ).execute()
                    print(f"Updated date slide text to: {today}")
                    break
        
        # -----------------------------
        # 3) Move the new date slide to the beginning
        # -----------------------------
        move_date_request = [{
            'updateSlidesPosition': {
                'slideObjectIds': [new_date_slide_id],
                'insertionIndex': 0
            }
        }]
        slides_service.presentations().batchUpdate(
            presentationId=presentation_id,
            body={'requests': move_date_request}
        ).execute()
        print("Moved new date slide to the beginning")
        
        # -----------------------------
        # 4) Duplicate the template slide (slide[1]) for each order
        # -----------------------------
        template_slide_id = slides[1].get('objectId')
        print(f"Using template slide ID: {template_slide_id}")
        
        new_order_slide_ids = []
        
        for i, order in enumerate(order_details):
            print(f"Creating slide for order {i+1}: {order.get('order_number')}")
            order_request = [{
                'duplicateObject': {
                    'objectId': template_slide_id,
                }
            }]
            
            order_response = slides_service.presentations().batchUpdate(
                presentationId=presentation_id,
                body={'requests': order_request}
            ).execute()
            
            new_slide_id = order_response.get('replies', [{}])[0].get('duplicateObject', {}).get('objectId')
            new_order_slide_ids.append(new_slide_id)
            print(f"Created new order slide with ID: {new_slide_id}")
            
            # Update the order details on this new slide
            update_order_details(slides_service, presentation_id, new_slide_id, order)
        
        # -----------------------------
        # 5) Reposition new order slides after the date slide
        # -----------------------------
        for i, slide_id in enumerate(new_order_slide_ids):
            move_request = [{
                'updateSlidesPosition': {
                    'slideObjectIds': [slide_id],
                    'insertionIndex': i + 1  # place them right after the date slide
                }
            }]
            slides_service.presentations().batchUpdate(
                presentationId=presentation_id,
                body={'requests': move_request}
            ).execute()
        
        print(f"Successfully repositioned {len(new_order_slide_ids)} order slides")
        
        # -----------------------------
        # No PDF generation here
        # -----------------------------
        pdf_path = None
        return presentation_url, pdf_path
    
    except Exception as e:
        print(f"ERROR in create_shipping_slides: {str(e)}")
        return None, None


def update_order_details(slides_service, presentation_id, slide_id, order):
    """
    Update a slide (which may contain shapes and/or a table) with order information.
    
    Args:
        slides_service: Google Slides API service
        presentation_id: ID of the presentation
        slide_id: ID of the slide to update
        order: Dictionary containing order information
    """
    print(f"Updating slide {slide_id} with order details for {order.get('order_number')}")
    
    # Wait a moment to ensure the slide duplication is complete
    time.sleep(0.5)
    
    try:
        # Fetch updated slide data
        slide_data = slides_service.presentations().get(
            presentationId=presentation_id
        ).execute()
        
        # Locate the target slide
        target_slide = None
        for slide in slide_data.get('slides', []):
            if slide.get('objectId') == slide_id:
                target_slide = slide
                break
        if not target_slide:
            print(f"WARNING: Could not find slide {slide_id}")
            return
        
        all_requests = []
        
        # -----------------------------
        # HELPER: Build a function to do placeholder replacements
        # -----------------------------
        def replace_placeholders(original_text):
            """
            Given the text in a shape or table cell, replace placeholders with order details.
            Returns the new text if changes were made, or the same text if no placeholders found.
            """
            new_text = original_text
            
            # Simple replacements:
            # 1) Name: # -> Name: #<order_number> <customer_name>
            if "NAME:" in new_text.upper():
                new_text = new_text.replace(
                    "Name: #", 
                    f"Name: #{order.get('order_number')} {order.get('name','')}"
                )
            
            # 2) Contact:
            if "CONTACT:" in new_text.upper() and "ECZEMA" not in new_text.upper():
                new_text = new_text.replace(
                    "Contact:",
                    f"Contact: {order.get('phone','N/A')}"
                )
            
            # 3) Delivery Address:
            if "DELIVERY ADDRESS:" in new_text.upper():
                address_parts = []
                if order.get('address1'):
                    address_parts.append(order.get('address1'))
                if order.get('address2'):
                    address_parts.append(order.get('address2'))
                
                address_str = "\n".join(address_parts) if address_parts else ""
                new_text = new_text.replace(
                    "Delivery Address:",
                    f"Delivery Address:\n{address_str}"
                )
            
            # 4) Postal: (not "Return Address:")
            #    If your template has "Postal:" in the same cell as "Return Address:",
            #    you may need a different approach. But here's the simple version:
            if "POSTAL:" in new_text.upper() and "RETURN ADDRESS:" not in new_text.upper():
                new_text = new_text.replace(
                    "Postal:",
                    f"Postal: {order.get('postal','')}"
                )
            
            # 5) Item:
            if "ITEM:" in new_text.upper():
                quantity = "2" if order.get('is_bundle') else "1"
                size = order.get('size', '')
                material = order.get('material', '')
                new_text = new_text.replace(
                    "Item:",
                    f"Item: {quantity} {size} {material} Eczema Mitten"
                )
            
            return new_text
        
        # -----------------------------
        # 1) Check all shapes
        # -----------------------------
        for element in target_slide.get('pageElements', []):
            if 'shape' in element and 'text' in element['shape']:
                element_id = element['objectId']
                text_elements = element['shape']['text']['textElements']
                
                # Combine all text runs
                original_text = ""
                for t_elem in text_elements:
                    if 'textRun' in t_elem:
                        original_text += t_elem['textRun'].get('content', '')
                
                new_text = replace_placeholders(original_text)
                
                if new_text != original_text:
                    # Build requests
                    all_requests.extend([
                        {
                            'deleteText': {
                                'objectId': element_id,
                                'textRange': {'type': 'ALL'}
                            }
                        },
                        {
                            'insertText': {
                                'objectId': element_id,
                                'insertionIndex': 0,
                                'text': new_text
                            }
                        }
                    ])
        
            # -----------------------------
            # 2) Check if this element is a table
            # -----------------------------
            elif 'table' in element:
                table_id = element['objectId']
                table = element['table']
                
                # Loop through rows/cells
                for row_index, row in enumerate(table.get('tableRows', [])):
                    for cell_index, cell in enumerate(row.get('tableCells', [])):
                        # Collect all paragraphs in this cell
                        original_text = ""
                        for content in cell.get('content', []):
                            paragraph = content.get('paragraph', {})
                            if 'elements' in paragraph:
                                for e in paragraph['elements']:
                                    text_run = e.get('textRun', {})
                                    original_text += text_run.get('content', '')
                        
                        # Replace placeholders
                        new_text = replace_placeholders(original_text)
                        
                        # If changed, build requests to delete and insert text
                        if new_text != original_text:
                            all_requests.extend([
                                {
                                    "deleteText": {
                                        "objectId": table_id,
                                        "cellLocation": {
                                            "rowIndex": row_index,
                                            "columnIndex": cell_index
                                        },
                                        "textRange": {
                                            "type": "ALL"
                                        }
                                    }
                                },
                                {
                                    "insertText": {
                                        "objectId": table_id,
                                        "insertionIndex": 0,
                                        "cellLocation": {
                                            "rowIndex": row_index,
                                            "columnIndex": cell_index
                                        },
                                        "text": new_text
                                    }
                                }
                            ])
        
        # -----------------------------
        # 3) Batch update if needed
        # -----------------------------
        if all_requests:
            print(f"Submitting {len(all_requests)} update requests for slide {slide_id}...")
            slides_service.presentations().batchUpdate(
                presentationId=presentation_id,
                body={'requests': all_requests}
            ).execute()
            print("Successfully updated order details for this slide.")
        else:
            print("WARNING: No fields were identified for update on this slide.")
    
    except Exception as e:
        print(f"ERROR updating order details on slide {slide_id}: {str(e)}")


def get_template_id_from_url(url):
    """Extract the presentation ID from a Google Slides URL"""
    if not url:
        return None
        
    match = re.search(r'/d/([a-zA-Z0-9-_]+)', url)
    if match:
        return match.group(1)
    return None

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
        SCOPES = ['https://www.googleapis.com/auth/drive']
        try:
            credentials = service_account.Credentials.from_service_account_file(
                credentials_path, scopes=SCOPES)
            print(f"Created credentials for {credentials.service_account_email}")
        except Exception as e:
            print(f"Error creating credentials: {str(e)}")
            return False
        
        drive_service = build('drive', 'v3', credentials=credentials)
        try:
            file = drive_service.files().get(fileId=template_id).execute()
            print(f"Successfully accessed template file: {file.get('name')}")
            return True
        except Exception as e:
            print(f"Error accessing template file: {str(e)}")
            error_str = str(e).lower()
            if 'permission' in error_str or 'access' in error_str or 'not found' in error_str:
                print("This is likely a permissions issue. Make sure the template has been shared with the service account email.")
                print(f"Service account email: {credentials.service_account_email}")
            return False
    except Exception as e:
        print(f"Error checking template permissions: {str(e)}")
        return False
