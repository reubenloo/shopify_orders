import os
import json
from googleapiclient.discovery import build
from google.oauth2 import service_account
from datetime import datetime
import re
import pytz
import time

def create_shipping_slides(order_details, credentials_path, template_id=None):
    """
    Edit an existing Google Slides presentation with shipping labels for orders
    
    Args:
        order_details: List of dictionaries containing order information
        credentials_path: Path to the service account JSON credentials file
        template_id: ID of the existing presentation to edit
        
    Returns:
        presentation_url: URL of the edited presentation
    """
    try:
        # Print detailed information for debugging
        print(f"Starting create_shipping_slides with {len(order_details)} orders")
        print(f"Credentials path: {credentials_path}")
        print(f"File exists: {os.path.exists(credentials_path)}")
        print(f"Template ID: {template_id or 'None'}")
        
        # Validate credentials file
        try:
            with open(credentials_path, 'r') as f:
                cred_content = f.read()
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
                    return None
        except Exception as e:
            print(f"ERROR: Could not read credentials file: {str(e)}")
            return None
        
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
            import traceback
            traceback.print_exc()
            return None
        
        # Create services
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
            return None
        
        # Use the existing presentation
        presentation_id = None
        presentation_url = None
        
        try:
            if template_id:
                print(f"Using existing presentation: {template_id}")
                presentation_id = template_id
            else:
                print("No template ID provided, cannot proceed")
                return None
                
            presentation_url = f"https://docs.google.com/presentation/d/{presentation_id}/edit"
            print(f"Presentation URL: {presentation_url}")
        except Exception as e:
            print(f"ERROR accessing presentation: {str(e)}")
            import traceback
            traceback.print_exc()
            return None
        
        # Get current presentation details
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
                return presentation_url
                
            # Save the date slide and template slide
            date_slide_id = slides[0].get('objectId')
            template_slide_id = slides[1].get('objectId')
            
            print(f"Found date slide with ID: {date_slide_id}")
            print(f"Found template slide with ID: {template_slide_id}")
        except Exception as e:
            print(f"ERROR getting presentation details: {str(e)}")
            import traceback
            traceback.print_exc()
            return None
        
        # IMPLEMENTATION OF THE SLIDE CREATION
        try:
            # Step 1: Create a new date slide by duplicating the existing date slide
            print("Creating new date slide at the beginning...")
            
            duplicate_request = {
                'duplicateObject': {
                    'objectId': date_slide_id
                }
            }
            
            response = slides_service.presentations().batchUpdate(
                presentationId=presentation_id,
                body={'requests': [duplicate_request]}
            ).execute()
            
            new_date_slide_id = response.get('replies', [{}])[0].get('duplicateObject', {}).get('objectId')
            if not new_date_slide_id:
                print("WARNING: Could not get ID for the new date slide")
            else:
                print(f"Created new date slide with ID: {new_date_slide_id}")
                
                # Move the new date slide to position 0
                move_request = {
                    'updateSlidesPosition': {
                        'slideObjectIds': [new_date_slide_id],
                        'insertionIndex': 0  # Put at the beginning
                    }
                }
                
                slides_service.presentations().batchUpdate(
                    presentationId=presentation_id,
                    body={'requests': [move_request]}
                ).execute()
                print("Moved date slide to the beginning")
                
                # Update the date on the new date slide
                update_date_slide(slides_service, presentation_id, new_date_slide_id)
            
            # Step 1b: Create a new template slide by duplicating the existing template slide
            # and position it after the new date slide
            print("Creating new template slide at position 1...")
            
            duplicate_template_request = {
                'duplicateObject': {
                    'objectId': template_slide_id
                }
            }
            
            template_response = slides_service.presentations().batchUpdate(
                presentationId=presentation_id,
                body={'requests': [duplicate_template_request]}
            ).execute()
            
            new_template_slide_id = template_response.get('replies', [{}])[0].get('duplicateObject', {}).get('objectId')
            if not new_template_slide_id:
                print("WARNING: Could not get ID for the new template slide")
            else:
                print(f"Created new template slide with ID: {new_template_slide_id}")
                
                # Move the new template slide to position 1 (right after the date slide)
                template_move_request = {
                    'updateSlidesPosition': {
                        'slideObjectIds': [new_template_slide_id],
                        'insertionIndex': 1  # Put right after the date slide
                    }
                }
                
                slides_service.presentations().batchUpdate(
                    presentationId=presentation_id,
                    body={'requests': [template_move_request]}
                ).execute()
                print("Moved template slide to position 1")
                
                # Now use this new template slide as our actual template
                template_slide_id = new_template_slide_id
            
            # Step 2: Create order detail slides, one for each order
            print(f"Creating {len(order_details)} order slides...")
            insert_index = 2  # Start inserting after the template slide (now at position 1)
            
            for i, order in enumerate(order_details):
                print(f"Processing order {i+1}: {order.get('order_number', 'unknown')}")
                
                # Create a copy of the template slide
                duplicate_request = {
                    'duplicateObject': {
                        'objectId': template_slide_id
                    }
                }
                
                response = slides_service.presentations().batchUpdate(
                    presentationId=presentation_id,
                    body={'requests': [duplicate_request]}
                ).execute()
                
                new_slide_id = response.get('replies', [{}])[0].get('duplicateObject', {}).get('objectId')
                if not new_slide_id:
                    print(f"WARNING: Could not get ID for the new order slide {i+1}")
                    continue
                    
                print(f"Created new order slide with ID: {new_slide_id}")
                
                # Position this slide after the template slide and before other order slides
                position_request = {
                    'updateSlidesPosition': {
                        'slideObjectIds': [new_slide_id],
                        'insertionIndex': insert_index
                    }
                }
                
                slides_service.presentations().batchUpdate(
                    presentationId=presentation_id,
                    body={'requests': [position_request]}
                ).execute()
                print(f"Positioned order slide at index {insert_index}")
                
                # Wait briefly to ensure the slide is fully created
                time.sleep(0.5)
                
                # Now update the order details on this slide using placeholder replacement
                update_slide_with_placeholders(slides_service, presentation_id, new_slide_id, order)
                
                # Increment the insertion index for the next slide
                insert_index += 1
            
            # Success!
            print(f"Successfully created slides for {len(order_details)} orders")
            return presentation_url
            
        except Exception as e:
            print(f"ERROR in main slide creation: {str(e)}")
            import traceback
            traceback.print_exc()
            return presentation_url
        
    except Exception as e:
        print(f"ERROR in create_shipping_slides: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

def update_date_slide(slides_service, presentation_id, slide_id):
    """
    Update the date on a slide to today's date
    
    Args:
        slides_service: Google Slides API service
        presentation_id: ID of the presentation
        slide_id: ID of the date slide
    """
    try:
        print(f"Updating date on slide {slide_id}...")
        
        # Get today's date
        today = datetime.now().strftime("%B %d, %Y")
        
        # Get the slide details
        slide = slides_service.presentations().pages().get(
            presentationId=presentation_id,
            pageObjectId=slide_id
        ).execute()
        
        # Find text elements on the slide
        text_elements = []
        for element in slide.get('pageElements', []):
            if 'shape' in element and 'text' in element.get('shape', {}):
                text_elements.append(element.get('objectId'))
                print(f"Found text element on date slide: {element.get('objectId')}")
        
        if not text_elements:
            print("WARNING: No text elements found on date slide")
            return
        
        # Update the first text element with today's date
        update_requests = []
        
        # Clear the existing text
        update_requests.append({
            'deleteText': {
                'objectId': text_elements[0],
                'textRange': {
                    'type': 'ALL'
                }
            }
        })
        
        # Insert the new date
        update_requests.append({
            'insertText': {
                'objectId': text_elements[0],
                'insertionIndex': 0,
                'text': today
            }
        })
        
        # Apply text style to match the template
        update_requests.append({
            'updateTextStyle': {
                'objectId': text_elements[0],
                'textRange': {
                    'type': 'ALL'
                },
                'style': {
                    'bold': True,
                    'fontSize': {
                        'magnitude': 24,
                        'unit': 'PT'
                    },
                    'foregroundColor': {
                        'opaqueColor': {
                            'rgbColor': {
                                'red': 0,
                                'green': 0,
                                'blue': 0
                            }
                        }
                    }
                },
                'fields': 'bold,fontSize,foregroundColor'
            }
        })
        
        # Apply paragraph style to center the text
        update_requests.append({
            'updateParagraphStyle': {
                'objectId': text_elements[0],
                'textRange': {
                    'type': 'ALL'
                },
                'style': {
                    'alignment': 'CENTER'
                },
                'fields': 'alignment'
            }
        })
        
        # Execute all updates
        slides_service.presentations().batchUpdate(
            presentationId=presentation_id,
            body={'requests': update_requests}
        ).execute()
        
        print(f"Successfully updated date to: {today}")
        
    except Exception as e:
        print(f"ERROR updating date slide: {str(e)}")
        import traceback
        traceback.print_exc()

def update_slide_with_placeholders(slides_service, presentation_id, slide_id, order):
    """
    Update a slide using placeholder text replacement
    
    Args:
        slides_service: Google Slides API service
        presentation_id: ID of the presentation
        slide_id: ID of the slide to update
        order: Dictionary containing order information
    """
    try:
        print(f"Updating slide {slide_id} with placeholder replacements for order: {order.get('order_number', 'unknown')}")
        
        # Prepare order information
        quantity = "2" if order.get('is_bundle', False) else "1"
        size = order.get('size', '')
        material = order.get('material', '')
        
        # Format the size
        if '(' in size and 'cm' in size:
            size_display = size.split('(')[1].replace(')', '').split('-')[0] + 'cm'
        else:
            size_display = size
            
        # Format the phone number to consistent +65 format
        phone = order.get('phone', '')
        if phone:
            # Remove any leading apostrophes
            if phone.startswith("'"):
                phone = phone[1:]
            
            # If it doesn't start with '+65', add it
            if not phone.startswith('+65'):
                # If it starts with a '6' or '65', remove it to avoid doubling the country code
                if phone.startswith('65'):
                    phone = phone[2:]
                elif phone.startswith('6'):
                    phone = phone[1:]
                
                # Add the +65 prefix
                phone = f"+65 {phone}"
            
            # Format with space after +65 and between groups of digits
            # First ensure there's a space after +65
            if '+65' in phone and not phone.startswith('+65 '):
                phone = phone.replace('+65', '+65 ')
            
            # If phone is just digits with no spaces, add a space between the 4th and 5th digits
            if len(phone.replace('+65 ', '').replace(' ', '')) == 8:
                digits = phone.replace('+65 ', '').replace(' ', '')
                phone = f"+65 {digits[:4]} {digits[4:]}"
            
        # Combine address lines
        address1 = order.get('address1', '')
        address2 = order.get('address2', '')
        address = f"{address1}\n{address2}" if address2 and address2.strip() else address1
        
        # Remove any leading apostrophe from postal code
        postal_code = order.get('postal', '')
        if postal_code and postal_code.startswith("'"):
            postal_code = postal_code[1:]
        
        # Create replacements for placeholders
        replacements = [
            {
                'find': '#ORDERNUM#',
                'replace': f"#{order.get('order_number', '').replace('#', '')}"
            },
            {
                'find': '#CUSTOMERNAME#',
                'replace': order.get('name', '')
            },
            {
                'find': '#PHONE#',
                'replace': phone
            },
            {
                'find': '#ADDRESS#',
                'replace': address
            },
            {
                'find': '#POSTALCODE#',
                'replace': postal_code
            },
            {
                'find': '#QUANTITY#',
                'replace': quantity
            },
            {
                'find': '#SIZE#',
                'replace': size_display
            },
            {
                'find': '#MATERIAL#',
                'replace': material
            }
        ]
        
        # Create replacement requests
        replace_requests = []
        for replacement in replacements:
            print(f"Creating replacement: '{replacement['find']}' -> '{replacement['replace']}'")
            replace_requests.append({
                'replaceAllText': {
                    'containsText': {
                        'text': replacement['find'],
                        'matchCase': True
                    },
                    'replaceText': replacement['replace'],
                    'pageObjectIds': [slide_id]
                }
            })
        
        # Execute replacements
        if replace_requests:
            print(f"Executing {len(replace_requests)} replacement requests")
            try:
                slides_service.presentations().batchUpdate(
                    presentationId=presentation_id,
                    body={'requests': replace_requests}
                ).execute()
                print("Successfully executed replacements")
            except Exception as e:
                print(f"WARNING: Error executing batch replacements: {str(e)}")
                
                # Try replacing one at a time
                print("Trying individual replacements...")
                for i, req in enumerate(replace_requests):
                    try:
                        slides_service.presentations().batchUpdate(
                            presentationId=presentation_id,
                            body={'requests': [req]}
                        ).execute()
                        print(f"Successfully executed replacement {i+1}")
                    except Exception as e2:
                        print(f"Failed replacement {i+1}: {str(e2)}")
        else:
            print("WARNING: No replacement requests were created")
    
    except Exception as e:
        print(f"ERROR updating slide with placeholders: {str(e)}")
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

# Legacy functions for backward compatibility
def update_order_details(slides_service, presentation_id, slide_id, order):
    """Legacy function that redirects to the new placeholder replacement approach"""
    return update_slide_with_placeholders(slides_service, presentation_id, slide_id, order)

def find_table_cells(slides_service, presentation_id, slide_id):
    """Legacy function - kept for backward compatibility but no longer used"""
    print("Warning: find_table_cells is deprecated and will not work properly")
    return {}

def update_text_fields(slides_service, presentation_id, text_fields, order):
    """Legacy function - kept for backward compatibility but no longer used"""
    print("Warning: update_text_fields is deprecated and will not work properly")
    return

def direct_update_text_on_slide(slides_service, presentation_id, slide_id, order):
    """Legacy function that redirects to the new placeholder replacement approach"""
    return update_slide_with_placeholders(slides_service, presentation_id, slide_id, order)

def update_table_based_slide(slides_service, presentation_id, slide_id, order):
    """Legacy function that redirects to the new placeholder replacement approach"""
    return update_slide_with_placeholders(slides_service, presentation_id, slide_id, order)