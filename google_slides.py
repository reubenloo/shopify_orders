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
        except Exception as e:
            print(f"ERROR getting presentation details: {str(e)}")
            import traceback
            traceback.print_exc()
            return None, None
        
        # IMPLEMENTATION OF THE NEW APPROACH
        try:
            # Step 1: Insert a new date slide at the beginning
            print("Creating new date slide at the beginning...")
            
            date_requests = [{
                'createSlide': {
                    'insertionIndex': 0,  # Place at the beginning
                    'slideLayoutReference': {
                        'predefinedLayout': 'BLANK'
                    }
                }
            }]
            
            date_response = slides_service.presentations().batchUpdate(
                presentationId=presentation_id,
                body={'requests': date_requests}
            ).execute()
            
            new_date_slide_id = date_response.get('replies', [{}])[0].get('createSlide', {}).get('objectId')
            if not new_date_slide_id:
                print("WARNING: Could not get ID for the new date slide")
            else:
                print(f"Created new date slide with ID: {new_date_slide_id}")
                
                # Copy content from the original date slide to the new date slide
                copy_slide_content(slides_service, presentation_id, date_slide_id, new_date_slide_id)
                
                # Update the date on the new date slide
                update_date_slide(slides_service, presentation_id, new_date_slide_id)
            
            # Step 2: Create order detail slides, one for each order
            print(f"Creating {len(order_details)} order slides...")
            new_slide_ids = []
            insertion_index = 1  # Start after the new date slide
            
            for i, order in enumerate(order_details):
                print(f"Processing order {i+1}: {order.get('order_number')}")
                
                # Create a new slide and insert it after the date slide
                order_slide_requests = [{
                    'createSlide': {
                        'insertionIndex': insertion_index + i,  # Position after date slide and previous order slides
                        'slideLayoutReference': {
                            'predefinedLayout': 'BLANK'
                        }
                    }
                }]
                
                order_slide_response = slides_service.presentations().batchUpdate(
                    presentationId=presentation_id,
                    body={'requests': order_slide_requests}
                ).execute()
                
                new_slide_id = order_slide_response.get('replies', [{}])[0].get('createSlide', {}).get('objectId')
                if not new_slide_id:
                    print(f"WARNING: Could not get ID for the new order slide {i+1}")
                    continue
                    
                print(f"Created new order slide with ID: {new_slide_id}")
                new_slide_ids.append(new_slide_id)
                
                # Copy content from the template slide to the new order slide
                copy_slide_content(slides_service, presentation_id, template_slide_id, new_slide_id)
                
                # Update the order details on this slide
                update_order_details(slides_service, presentation_id, new_slide_id, order)
            
            # Success!
            print(f"Successfully created {len(new_slide_ids)} new order slides")
            
            # Attempt to create a PDF export (not implemented in this version)
            pdf_path = None
            
            return presentation_url, pdf_path
            
        except Exception as e:
            print(f"ERROR in main slide creation: {str(e)}")
            import traceback
            traceback.print_exc()
            return presentation_url, None
        
    except Exception as e:
        print(f"ERROR in create_shipping_slides: {str(e)}")
        import traceback
        traceback.print_exc()
        return None, None

def copy_slide_content(slides_service, presentation_id, source_slide_id, target_slide_id):
    """
    Copy content (currently only shape elements with text) from a source slide to a target slide.
    """
    print(f"Copying content from slide {source_slide_id} to slide {target_slide_id}...")
    
    try:
        # Get the source slide's elements
        presentation = slides_service.presentations().get(
            presentationId=presentation_id,
            fields="slides(objectId,pageElements)"
        ).execute()
        
        source_slide = None
        for slide in presentation.get('slides', []):
            if slide.get('objectId') == source_slide_id:
                source_slide = slide
                break
                
        if not source_slide:
            print(f"WARNING: Could not find source slide {source_slide_id}")
            return
            
        create_requests = []
        element_counter = 0
        
        for element in source_slide.get('pageElements', []):
            element_counter += 1
            # Process only shape elements (extend later for images, tables, etc.)
            if 'shape' in element:
                shape = element['shape']
                shape_type = shape.get('shapeType', 'RECTANGLE')
                
                # Build elementProperties explicitly
                element_properties = {}
                if 'transform' in element:
                    element_properties['transform'] = element['transform']
                if 'size' in element:
                    element_properties['size'] = element['size']
                # Set target slide for the new element
                element_properties['pageObjectId'] = target_slide_id
                
                new_object_id = f"{target_slide_id}_{element_counter}"
                
                # Create the shape on the target slide
                shape_request = {
                    'createShape': {
                        'objectId': new_object_id,
                        'shapeType': shape_type,
                        'elementProperties': element_properties
                    }
                }
                create_requests.append(shape_request)
                
                # Check if the shape had text and extract it
                text = ""
                if 'text' in shape:
                    text_elements = shape.get('text', {}).get('textElements', [])
                    for text_element in text_elements:
                        if 'textRun' in text_element:
                            text += text_element['textRun'].get('content', '')
                
                if text.strip():
                    # Insert the text into the newly created shape
                    insert_text_request = {
                        'insertText': {
                            'objectId': new_object_id,
                            'insertionIndex': 0,
                            'text': text
                        }
                    }
                    create_requests.append(insert_text_request)
                    
            # You could add additional handling for images, tables, etc.
        
        if create_requests:
            slides_service.presentations().batchUpdate(
                presentationId=presentation_id,
                body={'requests': create_requests}
            ).execute()
            print(f"Successfully copied {len(create_requests)} element requests to target slide")
        else:
            print("WARNING: No elements were found to copy")
        
    except Exception as e:
        print(f"ERROR copying slide content: {str(e)}")
        import traceback
        traceback.print_exc()


def update_date_slide(slides_service, presentation_id, slide_id):
    """
    Update the date on a slide to today's date
    
    Args:
        slides_service: Google Slides API service
        presentation_id: ID of the presentation
        slide_id: ID of the date slide
    """
    print(f"Updating date on slide {slide_id}...")
    
    try:
        # Get today's date
        today = datetime.now().strftime("%B %d, %Y")
        
        # Get the slide details
        slide_data = slides_service.presentations().get(
            presentationId=presentation_id,
            fields="slides"
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
            
        # Find and update the text elements
        for element in target_slide.get('pageElements', []):
            if 'shape' in element and 'text' in element.get('shape', {}):
                # Found a text element, update it with today's date
                element_id = element.get('objectId')
                
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
                
                print(f"Updated date to: {today}")
                break  # Only update the first text element
                
        # If no text element exists, create one
        if not target_slide.get('pageElements'):
            print("Creating new text element for date")
            create_requests = [{
                'createShape': {
                    'objectId': f"date_text_{slide_id}",
                    'shapeType': 'TEXT_BOX',
                    'elementProperties': {
                        'pageObjectId': slide_id,
                        'size': {
                            'width': {'magnitude': 300, 'unit': 'PT'},
                            'height': {'magnitude': 50, 'unit': 'PT'}
                        },
                        'transform': {
                            'scaleX': 1,
                            'scaleY': 1,
                            'translateX': 100,
                            'translateY': 100,
                            'unit': 'PT'
                        }
                    }
                }
            }, {
                'insertText': {
                    'objectId': f"date_text_{slide_id}",
                    'text': today
                }
            }]
            
            slides_service.presentations().batchUpdate(
                presentationId=presentation_id,
                body={'requests': create_requests}
            ).execute()
            
            print(f"Created and updated date to: {today}")
        
    except Exception as e:
        print(f"ERROR updating date slide: {str(e)}")
        import traceback
        traceback.print_exc()

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
    
    try:
        # Get the slide details
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
        
        # Create a map of fields to look for and their updated values
        field_map = {
            'NAME': {
                'pattern': ['NAME:', 'NAME :', 'NAME #'],
                'new_text': f"Name: #{order.get('order_number')} {order.get('name', '')}"
            },
            'CONTACT': {
                'pattern': ['CONTACT:', 'CONTACT :', '+65 9'],
                'excludes': ['ECZEMA', 'COMPANY'],
                'new_text': f"Contact: {order.get('phone', 'N/A')}"
            },
            'ADDRESS': {
                'pattern': ['DELIVERY ADDRESS:', 'DELIVERY ADDRESS :', 'ADDRESS:'],
                'excludes': ['RETURN'],
                'new_text': f"Delivery Address:\n{order.get('address1', '')}\n{order.get('address2', '')}"
            },
            'POSTAL': {
                'pattern': ['POSTAL:', 'POSTAL :', 'POSTAL CODE:'],
                'excludes': ['680', 'RETURN'],
                'new_text': f"Postal: {order.get('postal', '')}"
            },
            'ITEM': {
                'pattern': ['ITEM:', 'ITEM :', 'PRODUCT:'],
                'new_text': f"Item: {2 if order.get('is_bundle') else 1} {order.get('size', '')} {order.get('material', '')} Eczema Mitten"
            }
        }
        
        # Keep track of all update requests
        all_requests = []
        
        # Look for text elements on the slide
        for element in target_slide.get('pageElements', []):
            if 'shape' in element and 'text' in element.get('shape', {}):
                element_id = element.get('objectId')
                text_content = element.get('shape', {}).get('text', {}).get('textElements', [])
                
                # Extract the full text of this element
                full_text = ""
                for text_element in text_content:
                    if 'textRun' in text_element:
                        full_text += text_element.get('textRun', {}).get('content', '')
                
                # Clean up the text and convert to uppercase for matching
                clean_text = full_text.strip().upper()
                
                # Print detected text for debugging
                print(f"Found text element: '{clean_text}'")
                
                # Check each field to see if this element matches
                for field, config in field_map.items():
                    patterns = config['pattern']
                    excludes = config.get('excludes', [])
                    
                    # Check if any pattern matches and no excludes match
                    pattern_match = any(p.upper() in clean_text for p in patterns)
                    exclude_match = any(e.upper() in clean_text for e in excludes)
                    
                    if pattern_match and not exclude_match:
                        print(f"Identified {field} field: '{clean_text}'")
                        
                        # Create requests to update this field
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
                                'text': config['new_text']
                            }
                        }])
                        
                        print(f"Will update {field} to: {config['new_text']}")
                        break  # Move to next element
        
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