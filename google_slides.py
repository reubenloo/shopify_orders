import os
import json
from googleapiclient.discovery import build
from google.oauth2 import service_account
from datetime import datetime, timedelta
import re
import pytz
import copy

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
        # Print detailed information for debugging
        print(f"Starting create_shipping_slides with {len(order_details)} orders")
        print(f"Credentials path: {credentials_path}")
        print(f"File exists: {os.path.exists(credentials_path)}")
        print(f"Template ID: {template_id or 'None'}")
        
        # Read and print first part of credentials file to verify it's valid
        try:
            with open(credentials_path, 'r') as f:
                cred_content = f.read()
                # Print first 100 chars with sensitive info masked
                masked_content = cred_content[:100].replace('"private_key":', '"private_key": "[MASKED]",')
                print(f"Credentials file content (first 100 chars): {masked_content}...")
                
                # Validate it's proper JSON
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
        
        # Set up credentials with detailed error handling
        SCOPES = ['https://www.googleapis.com/auth/presentations', 
                 'https://www.googleapis.com/auth/drive']
        
        try:
            print("Creating credentials object...")
            credentials = service_account.Credentials.from_service_account_file(
                credentials_path, scopes=SCOPES)
            print(f"Credentials created successfully: {credentials.service_account_email}")
        except Exception as e:
            print(f"ERROR creating credentials: {str(e)}")
            import traceback
            traceback.print_exc()
            return None, None
        
        # Create services with detailed error handling
        try:
            print("Building slides service...")
            slides_service = build('slides', 'v1', credentials=credentials)
            print("Building drive service...")
            drive_service = build('drive', 'v3', credentials=credentials)
            print("Services built successfully")
        except Exception as e:
            print(f"ERROR building services: {str(e)}")
            import traceback
            traceback.print_exc()
            return None, None
        
        # Use the existing presentation if template_id is provided
        presentation_id = None
        presentation_url = None
        
        try:
            if template_id:
                print(f"Using existing presentation: {template_id}")
                presentation_id = template_id
            else:
                print("No template ID provided, cannot proceed")
                return None, None
                
            presentation_url = f"https://docs.google.com/presentation/d/{presentation_id}/edit"
            print(f"Presentation URL: {presentation_url}")
        except Exception as e:
            print(f"ERROR accessing presentation: {str(e)}")
            import traceback
            traceback.print_exc()
            return None, None
        
        # SIMPLIFIED APPROACH: Use duplicateObject API call for copying slides
        try:
            # Step 1: Create a copy of the entire presentation as a temporary working copy
            print("Creating a new batch of slides at the top...")
            
            # First, get the date from the first slide
            presentation = slides_service.presentations().get(
                presentationId=presentation_id
            ).execute()
            
            slides = presentation.get('slides', [])
            if len(slides) < 2:
                print("ERROR: Presentation must have at least 2 slides (date and template)")
                return presentation_url, None
            
            # Create a new date slide by duplicating the first slide
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
            
            # Update the date on the new slide
            # First get text elements on the new date slide
            new_slides_data = slides_service.presentations().get(
                presentationId=presentation_id,
                fields="slides"
            ).execute()
            
            # Find the newly created date slide
            new_date_slide = None
            for slide in new_slides_data.get('slides', []):
                if slide.get('objectId') == new_date_slide_id:
                    new_date_slide = slide
                    break
            
            if new_date_slide:
                # Update the date text
                for element in new_date_slide.get('pageElements', []):
                    if 'shape' in element and 'text' in element.get('shape', {}):
                        element_id = element.get('objectId')
                        today = datetime.now().strftime("%B %d, %Y")
                        
                        date_update_request = [{
                            'deleteText': {
                                'objectId': element_id,
                                'textRange': {
                                    'type': 'ALL'
                                }
                            }
                        }, {
                            'insertText': {
                                'objectId': element_id,
                                'insertionIndex': 0,  # FIX: added insertionIndex
                                'text': today
                            }
                        }]
                        
                        slides_service.presentations().batchUpdate(
                            presentationId=presentation_id,
                            body={'requests': date_update_request}
                        ).execute()
                        print(f"Updated date to: {today}")
                        break
            
            # Move the new date slide to the beginning
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
            print("Moved date slide to the beginning")
            
            # Now create a slide for each order by duplicating the template slide (slide 2)
            template_slide_id = slides[1].get('objectId')
            print(f"Using template slide ID: {template_slide_id}")
            
            # Track the new slide IDs
            new_order_slide_ids = []
            
            # Create a slide for each order
            for i, order in enumerate(order_details):
                print(f"Creating slide for order {i+1}: {order.get('order_number')}")
                
                # Duplicate the template slide
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
                
                # Update the order details on this slide
                update_order_details(slides_service, presentation_id, new_slide_id, order)
            
            # Reposition all the new order slides after the date slide
            for i, slide_id in enumerate(new_order_slide_ids):
                move_request = [{
                    'updateSlidesPosition': {
                        'slideObjectIds': [slide_id],
                        'insertionIndex': i + 1  # Position after date slide
                    }
                }]
                
                slides_service.presentations().batchUpdate(
                    presentationId=presentation_id,
                    body={'requests': move_request}
                ).execute()
                
            print(f"Successfully repositioned {len(new_order_slide_ids)} order slides")
            
            # No PDF generation in this version
            pdf_path = None
            
            return presentation_url, pdf_path
            
        except Exception as e:
            print(f"ERROR in slide creation: {str(e)}")
            import traceback
            traceback.print_exc()
            return presentation_url, None
        
    except Exception as e:
        print(f"ERROR in create_shipping_slides: {str(e)}")
        import traceback
        traceback.print_exc()
        return None, None

def update_order_details(slides_service, presentation_id, slide_id, order):
    """
    Update a slide with order information
    
    Args:
        slides_service: Google Slides API service
        presentation_id: ID of the presentation
        slide_id: ID of the slide to update
        order: Dictionary containing order information
    """
    print(f"Updating slide {slide_id} with order details for {order.get('order_number')}")
    
    # Wait a moment to ensure the slide duplication is complete
    import time
    time.sleep(0.5)
    
    try:
        # Get fresh slide data
        slide_data = slides_service.presentations().get(
            presentationId=presentation_id
        ).execute()
        
        # Find our target slide
        target_slide = None
        for slide in slide_data.get('slides', []):
            if slide.get('objectId') == slide_id:
                target_slide = slide
                break
                
        if not target_slide:
            print(f"WARNING: Could not find slide {slide_id}")
            return
        
        # Keep track of all update requests
        all_requests = []
        
        # Look for text elements on the slide
        for element in target_slide.get('pageElements', []):
            if 'shape' in element and 'text' in element.get('shape', {}):
                element_id = element.get('objectId')
                
                # Get all text elements inside this shape
                text_elements = element.get('shape', {}).get('text', {}).get('textElements', [])
                
                # Extract the full text of this element
                full_text = ""
                for text_element in text_elements:
                    if 'textRun' in text_element:
                        full_text += text_element.get('textRun', {}).get('content', '')
                
                # Clean up the text and print for debugging
                clean_text = full_text.strip()
                print(f"Found text element: '{clean_text}'")
                
                # MUCH SIMPLER FIELD DETECTION
                # Instead of complex pattern matching, look for common text patterns
                clean_upper = clean_text.upper()
                
                if "NAME:" in clean_upper and "COMPANY:" not in clean_upper:
                    # This is the name field for the customer
                    new_text = f"Name: #{order.get('order_number')} {order.get('name', '')}"
                    
                    all_requests.extend([{
                        'deleteText': {
                            'objectId': element_id,
                            'textRange': {
                                'type': 'ALL'
                            }
                        }
                    }, {
                        'insertText': {
                            'objectId': element_id,
                            'insertionIndex': 0,  # FIX: added insertionIndex
                            'text': new_text
                        }
                    }])
                    print(f"Will update name to: {new_text}")
                    
                elif "CONTACT:" in clean_upper and "ECZEMA" not in clean_upper:
                    # This is the customer contact field
                    new_text = f"Contact: {order.get('phone', 'N/A')}"
                    
                    all_requests.extend([{
                        'deleteText': {
                            'objectId': element_id,
                            'textRange': {
                                'type': 'ALL'
                            }
                        }
                    }, {
                        'insertText': {
                            'objectId': element_id,
                            'insertionIndex': 0,  # FIX: added insertionIndex
                            'text': new_text
                        }
                    }])
                    print(f"Will update contact to: {new_text}")
                    
                elif "DELIVERY ADDRESS:" in clean_upper:
                    # This is the delivery address field
                    address_parts = []
                    if order.get('address1'):
                        address_parts.append(order.get('address1'))
                    if order.get('address2'):
                        address_parts.append(order.get('address2'))
                    
                    new_text = "Delivery Address:"
                    if address_parts:
                        new_text += "\n" + "\n".join(address_parts)
                    
                    all_requests.extend([{
                        'deleteText': {
                            'objectId': element_id,
                            'textRange': {
                                'type': 'ALL'
                            }
                        }
                    }, {
                        'insertText': {
                            'objectId': element_id,
                            'insertionIndex': 0,  # FIX: added insertionIndex
                            'text': new_text
                        }
                    }])
                    print(f"Will update delivery address to: {new_text}")
                    
                elif "POSTAL:" in clean_upper and "RETURN ADDRESS:" not in clean_upper:
                    # This is the customer postal code field
                    new_text = f"Postal: {order.get('postal', '')}"
                    
                    all_requests.extend([{
                        'deleteText': {
                            'objectId': element_id,
                            'textRange': {
                                'type': 'ALL'
                            }
                        }
                    }, {
                        'insertText': {
                            'objectId': element_id,
                            'insertionIndex': 0,  # FIX: added insertionIndex
                            'text': new_text
                        }
                    }])
                    print(f"Will update postal code to: {new_text}")
                    
                elif "ITEM:" in clean_upper:
                    # This is the product details field
                    quantity = "2" if order.get('is_bundle') else "1"
                    size = order.get('size', '')
                    material = order.get('material', '')
                    
                    new_text = f"Item: {quantity} {size} {material} Eczema Mitten"
                    
                    all_requests.extend([{
                        'deleteText': {
                            'objectId': element_id,
                            'textRange': {
                                'type': 'ALL'
                            }
                        }
                    }, {
                        'insertText': {
                            'objectId': element_id,
                            'insertionIndex': 0,  # FIX: added insertionIndex
                            'text': new_text
                        }
                    }])
                    print(f"Will update item to: {new_text}")
        
        # Submit all update requests at once
        if all_requests:
            print(f"Submitting {len(all_requests)} update requests...")
            slides_service.presentations().batchUpdate(
                presentationId=presentation_id,
                body={'requests': all_requests}
            ).execute()
            print("Successfully updated order details")
        else:
            print("WARNING: No fields were identified for update")
    
    except Exception as e:
        print(f"ERROR updating order details: {str(e)}")
        import traceback
        traceback.print_exc()

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
        # Set up credentials
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
