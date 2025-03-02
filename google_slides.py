import os
import json
from googleapiclient.discovery import build
from google.oauth2 import service_account
from datetime import datetime, timedelta
import re
import pytz

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
        
        # Get current presentation details to understand slide layout
        try:
            presentation = slides_service.presentations().get(
                presentationId=presentation_id
            ).execute()
            print(f"Fetched presentation details, title: {presentation.get('title')}")
            
            # Get existing slides
            slides = presentation.get('slides', [])
            print(f"Presentation has {len(slides)} existing slides")
            
            # Ensure we have at least 2 slides (first for date, second for template)
            if len(slides) < 2:
                print("ERROR: Template presentation should have at least 2 slides")
                return presentation_url, None
                
            # Save the date slide and template slide
            date_slide_id = slides[0].get('objectId')
            template_slide_id = slides[1].get('objectId')
            
            print(f"Found date slide with ID: {date_slide_id}")
            print(f"Found template slide with ID: {template_slide_id}")
            
            # Store all slide IDs
            all_slide_ids = [slide.get('objectId') for slide in slides]
        except Exception as e:
            print(f"ERROR getting presentation details: {str(e)}")
            import traceback
            traceback.print_exc()
            return None, None
        
        # Update the date on the first slide
        try:
            print("Updating date on first slide...")
            
            # Get the current date
            today = datetime.now().strftime("%B %d, %Y")
            
            # Find text elements on the first slide
            date_slide = slides[0]
            date_updated = False
            
            for element in date_slide.get('pageElements', []):
                if 'shape' in element and 'text' in element.get('shape', {}):
                    element_id = element.get('objectId')
                    
                    # Update the text to today's date
                    requests = [{
                        'deleteText': {
                            'objectId': element_id,
                            'textRange': {
                                'type': 'ALL'
                            }
                        }
                    }, {
                        'insertText': {
                            'objectId': element_id,
                            'text': today
                        }
                    }]
                    
                    slides_service.presentations().batchUpdate(
                        presentationId=presentation_id,
                        body={'requests': requests}
                    ).execute()
                    
                    date_updated = True
                    print(f"Updated date to: {today}")
                    break
            
            if not date_updated:
                print("WARNING: Could not find text element on the date slide")
        except Exception as e:
            print(f"ERROR updating date slide: {str(e)}")
            import traceback
            traceback.print_exc()
            # Continue anyway
        
        # Clear all slides except the first two (date slide and template slide)
        try:
            if len(all_slide_ids) > 2:
                print(f"Removing {len(all_slide_ids) - 2} existing order slides...")
                requests = []
                
                for slide_id in all_slide_ids[2:]:
                    requests.append({
                        'deleteObject': {
                            'objectId': slide_id
                        }
                    })
                
                if requests:
                    slides_service.presentations().batchUpdate(
                        presentationId=presentation_id,
                        body={'requests': requests}
                    ).execute()
                    print("Removed old order slides")
        except Exception as e:
            print(f"ERROR removing old slides: {str(e)}")
            import traceback
            traceback.print_exc()
            # Continue anyway
        
        # Create shipping label slides for each order by duplicating the template slide
        try:
            print(f"Creating shipping label slides for {len(order_details)} orders...")
            
            # Store the IDs of the newly created slides
            new_slide_ids = []
            
            # Duplicate the template slide for each order
            for i, order in enumerate(order_details):
                print(f"Creating slide for order {i+1}: {order.get('order_number')}")
                
                # Duplicate the template slide
                requests = [{
                    'duplicateObject': {
                        'objectId': template_slide_id,
                    }
                }]
                
                response = slides_service.presentations().batchUpdate(
                    presentationId=presentation_id,
                    body={'requests': requests}
                ).execute()
                
                # Get the ID of the newly created slide
                new_slide_id = response.get('replies', [{}])[0].get('duplicateObject', {}).get('objectId')
                if not new_slide_id:
                    print(f"WARNING: Could not get ID for duplicated slide for order {order.get('order_number')}")
                    continue
                    
                new_slide_ids.append(new_slide_id)
                print(f"Created new slide with ID: {new_slide_id}")
                
                # Update the text on the new slide with order details
                try:
                    update_order_slide(slides_service, presentation_id, new_slide_id, order)
                except Exception as e:
                    print(f"ERROR updating slide content for order {order.get('order_number')}: {str(e)}")
                    import traceback
                    traceback.print_exc()
            
            # Move all new slides after the date slide
            print("Rearranging slides...")
            requests = []
            
            # Position new slides after the date slide (keep template slide at the end)
            for i, slide_id in enumerate(new_slide_ids):
                requests.append({
                    'updateSlidesPosition': {
                        'slideObjectIds': [slide_id],
                        'insertionIndex': i + 1  # Position after date slide
                    }
                })
            
            # Move template slide to the end
            requests.append({
                'updateSlidesPosition': {
                    'slideObjectIds': [template_slide_id],
                    'insertionIndex': len(new_slide_ids) + 1  # Position after all new slides
                }
            })
            
            if requests:
                slides_service.presentations().batchUpdate(
                    presentationId=presentation_id,
                    body={'requests': requests}
                ).execute()
                print("Slides rearranged successfully")
            
            print("All slides created and positioned successfully!")
            
            # Attempt to create a PDF export
            # Note: Direct PDF export is limited with the Drive API, so we'll just
            # return None for pdf_path, and users will need to download PDF from Google Slides UI
            pdf_path = None
            
            return presentation_url, pdf_path
            
        except Exception as e:
            print(f"ERROR creating slides: {str(e)}")
            import traceback
            traceback.print_exc()
            # Return the URL if we have it, even if there was an error adding content
            return presentation_url, None
        
    except Exception as e:
        print(f"ERROR in create_shipping_slides: {str(e)}")
        import traceback
        traceback.print_exc()
        return None, None

def update_order_slide(slides_service, presentation_id, slide_id, order):
    """
    Update a slide with order information
    
    Args:
        slides_service: Google Slides API service
        presentation_id: ID of the presentation
        slide_id: ID of the slide to update
        order: Dictionary containing order information
    """
    print(f"Updating slide {slide_id} with order details for {order.get('order_number')}")
    
    # Get the slide details
    slide = slides_service.presentations().get(
        presentationId=presentation_id,
        fields=f"slides(objectId,pageElements)"
    ).execute()
    
    # Find the slide
    target_slide = None
    for s in slide.get('slides', []):
        if s.get('objectId') == slide_id:
            target_slide = s
            break
    
    if not target_slide:
        print(f"WARNING: Could not find slide {slide_id} in presentation")
        return
    
    # Look for text elements that contain specific keywords
    requests = []
    
    for element in target_slide.get('pageElements', []):
        if 'shape' in element and 'text' in element.get('shape', {}):
            element_id = element.get('objectId')
            text_content = element.get('shape', {}).get('text', {}).get('textElements', [])
            
            if not text_content:
                continue
            
            full_text = ""
            for text_element in text_content:
                if 'textRun' in text_element:
                    full_text += text_element.get('textRun', {}).get('content', '')
            
            # Clean up text and convert to uppercase for matching
            clean_text = full_text.strip().upper()
            
            # Find fields to update based on content
            if "NAME:" in clean_text:
                # This is the order number/name field
                new_text = f"Name: #{order.get('order_number')} {order.get('name', '')}"
                requests.append(create_text_update_request(element_id, new_text))
                print(f"Updated customer name: {new_text}")
                
            elif "CONTACT: +65 9" in clean_text:
                # This is the customer contact field
                new_text = f"Contact: {order.get('phone', 'N/A')}"
                requests.append(create_text_update_request(element_id, new_text))
                print(f"Updated customer phone: {new_text}")
                
            elif "DELIVERY ADDRESS:" in clean_text:
                # This is the delivery address field
                address_parts = []
                if order.get('address1'):
                    address_parts.append(order.get('address1'))
                if order.get('address2'):
                    address_parts.append(order.get('address2'))
                
                address_text = f"Delivery Address:\n{' '.join(address_parts)}"
                requests.append(create_text_update_request(element_id, address_text))
                print(f"Updated delivery address: {address_text}")
                
            elif "POSTAL: 5" in clean_text:
                # This is the customer postal code field
                new_text = f"Postal: {order.get('postal', '')}"
                requests.append(create_text_update_request(element_id, new_text))
                print(f"Updated postal code: {new_text}")
                
            elif "ITEM: " in clean_text:
                # This is the item field
                quantity = "2" if order.get('is_bundle') else "1"
                size = order.get('size', '')
                material = order.get('material', '')
                
                item_text = f"Item: {quantity} {size} {material} Eczema Mitten"
                requests.append(create_text_update_request(element_id, item_text))
                print(f"Updated item details: {item_text}")
    
    # Submit the updates
    if requests:
        print(f"Submitting {len(requests)} text updates for slide...")
        try:
            response = slides_service.presentations().batchUpdate(
                presentationId=presentation_id,
                body={'requests': requests}
            ).execute()
            print(f"Successfully updated slide with {len(response.get('replies', []))} changes")
        except Exception as e:
            print(f"ERROR updating slide: {str(e)}")
            import traceback
            traceback.print_exc()
    else:
        print("WARNING: No text elements were updated on the slide")

def create_text_update_request(element_id, new_text):
    """
    Create a request to update text on a slide
    
    Args:
        element_id: ID of the text element
        new_text: New text content
        
    Returns:
        dict: Request object for batchUpdate
    """
    return [{
        'deleteText': {
            'objectId': element_id,
            'textRange': {
                'type': 'ALL'
            }
        }
    }, {
        'insertText': {
            'objectId': element_id,
            'text': new_text
        }
    }]

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