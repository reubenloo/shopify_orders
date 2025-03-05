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
            
            # Step 2: Create order detail slides, one for each order
            print(f"Creating {len(order_details)} order slides...")
            insert_index = 1  # Start inserting after the date slide
            
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
                
                # Position this slide after the date slide and before other order slides
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
                
                # Wait briefly for the slide to be fully created
                time.sleep(0.5)
                
                # Now update the order details on this slide
                direct_update_text_on_slide(slides_service, presentation_id, new_slide_id, order)
                
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

def direct_update_text_on_slide(slides_service, presentation_id, slide_id, order):
    """
    A simplified approach to update text on a slide without relying on table cells
    
    Args:
        slides_service: Google Slides API service
        presentation_id: ID of the presentation
        slide_id: ID of the slide to update
        order: Dictionary containing order information
    """
    try:
        print(f"Directly updating text on slide {slide_id} for order: {order.get('order_number', 'unknown')}")
        
        # Get the slide structure
        slide = slides_service.presentations().pages().get(
            presentationId=presentation_id,
            pageObjectId=slide_id
        ).execute()
        
        # Extract all shape text elements from the slide
        elements_info = []
        for element in slide.get('pageElements', []):
            if 'shape' in element and 'text' in element.get('shape', {}):
                element_id = element.get('objectId')
                
                # Extract text content for this shape
                text_content = ""
                for text_element in element.get('shape', {}).get('text', {}).get('textElements', []):
                    if 'textRun' in text_element:
                        text_content += text_element.get('textRun', {}).get('content', '')
                
                # Get position for sorting
                y_pos = element.get('transform', {}).get('translateY', 0)
                x_pos = element.get('transform', {}).get('translateX', 0)
                
                # Save element info
                elements_info.append({
                    'id': element_id,
                    'text': text_content,
                    'y': y_pos,
                    'x': x_pos
                })
                
                print(f"Found text element: ID={element_id}, Content=\"{text_content.strip()}\", X={x_pos}, Y={y_pos}")
        
        # Prepare order information
        quantity = "2" if order.get('is_bundle', False) else "1"
        size = order.get('size', '')
        material = order.get('material', '')
        
        # Format the size
        if '(' in size and 'cm' in size:
            size_display = size.split('(')[1].replace(')', '').split('-')[0] + 'cm'
        else:
            size_display = size
            
        # Combine address lines
        address1 = order.get('address1', '')
        address2 = order.get('address2', '')
        address = f"{address1}\n{address2}" if address2 and address2.strip() else address1
        
        # Define fields to look for
        fields_to_update = [
            {
                'patterns': ['NAME:', 'NAME #'],  # Text to match for name field
                'content': f"Name: #{order.get('order_number', '').replace('#', '')} {order.get('name', '')}"
            },
            {
                'patterns': ['CONTACT:'],  # Text to match for contact field
                'content': f"Contact: {order.get('phone', '')}"
            },
            {
                'patterns': ['DELIVERY ADDRESS:', 'DELIVERY ADDRESS'],  # Text to match for address field
                'content': f"Delivery Address: {address}"
            },
            {
                'patterns': ['POSTAL:'],  # Text to match for postal field
                'content': f"Postal: {order.get('postal', '')}"
            },
            {
                'patterns': ['ITEM:'],  # Text to match for item field
                'content': f"Item: {quantity} {size_display} {material} Eczema Mitten"
            }
        ]
        
        # Strategy 1: Try to match fields by their text content
        update_requests = []
        matched_elements = set()
        
        for field in fields_to_update:
            for pattern in field['patterns']:
                for element in elements_info:
                    if element['id'] not in matched_elements and pattern in element['text'].upper():
                        print(f"Found match for '{pattern}' in element: {element['id']}")
                        matched_elements.add(element['id'])
                        
                        # Create update request
                        update_requests.extend([
                            {
                                'deleteText': {
                                    'objectId': element['id'],
                                    'textRange': {
                                        'type': 'ALL'
                                    }
                                }
                            },
                            {
                                'insertText': {
                                    'objectId': element['id'],
                                    'insertionIndex': 0,
                                    'text': field['content']
                                }
                            }
                        ])
                        break  # Move to next field
        
        # Strategy 2: If we didn't match all fields, try by position (left side of slide)
        if len(matched_elements) < len(fields_to_update):
            print(f"Only matched {len(matched_elements)} fields using text matching. Trying position-based matching.")
            
            # Filter to only include elements on the left side (customer info side)
            left_elements = [e for e in elements_info if e['id'] not in matched_elements and e['x'] < 200]
            
            # Sort by vertical position
            left_elements.sort(key=lambda e: e['y'])
            
            # Try to match remaining fields
            remaining_fields = [f for f in fields_to_update if not any(p in [e['text'].upper() for e in elements_info if e['id'] in matched_elements] for p in f['patterns'])]
            
            print(f"Found {len(left_elements)} left-side elements for {len(remaining_fields)} remaining fields")
            
            for i, field in enumerate(remaining_fields):
                if i < len(left_elements):
                    element = left_elements[i]
                    print(f"Position-based match: Field '{field['patterns'][0]}' -> Element {element['id']}")
                    matched_elements.add(element['id'])
                    
                    update_requests.extend([
                        {
                            'deleteText': {
                                'objectId': element['id'],
                                'textRange': {
                                    'type': 'ALL'
                                }
                            }
                        },
                        {
                            'insertText': {
                                'objectId': element['id'],
                                'insertionIndex': 0,
                                'text': field['content']
                            }
                        }
                    ])
        
        # Strategy 3: Last resort - find all editable text elements
        if len(matched_elements) < len(fields_to_update):
            print("Some fields still not matched. Attempting direct field creation.")
            
            # Create direct field updates
            create_text_requests = []
            
            # Fixed positions for new text boxes
            positions = [
                {'x': 50, 'y': 100},  # Name
                {'x': 50, 'y': 130},  # Contact
                {'x': 50, 'y': 160},  # Address
                {'x': 50, 'y': 190},  # Postal
                {'x': 50, 'y': 220}   # Item
            ]
            
            for i, field in enumerate(fields_to_update):
                if i < len(positions) and not any(p in [e['text'].upper() for e in elements_info if e['id'] in matched_elements] for p in field['patterns']):
                    print(f"Creating new text box for field: {field['patterns'][0]}")
                    
                    # Generate a unique element ID
                    element_id = f'customText_{slide_id}_{i}'
                    
                    # Create a new text box
                    create_text_requests.append({
                        'createShape': {
                            'objectId': element_id,
                            'shapeType': 'TEXT_BOX',
                            'elementProperties': {
                                'pageObjectId': slide_id,
                                'size': {
                                    'width': {'magnitude': 350, 'unit': 'PT'},
                                    'height': {'magnitude': 30, 'unit': 'PT'}
                                },
                                'transform': {
                                    'scaleX': 1,
                                    'scaleY': 1,
                                    'translateX': positions[i]['x'],
                                    'translateY': positions[i]['y'],
                                    'unit': 'PT'
                                }
                            }
                        }
                    })
                    
                    # Add text to the text box
                    create_text_requests.append({
                        'insertText': {
                            'objectId': element_id,
                            'insertionIndex': 0,
                            'text': field['content']
                        }
                    })
                    
                    # Style the text
                    create_text_requests.append({
                        'updateTextStyle': {
                            'objectId': element_id,
                            'textRange': {
                                'type': 'ALL'
                            },
                            'style': {
                                'bold': True,
                                'fontSize': {
                                    'magnitude': 12,
                                    'unit': 'PT'
                                }
                            },
                            'fields': 'bold,fontSize'
                        }
                    })
            
            # Execute create requests if needed
            if create_text_requests:
                try:
                    print(f"Executing {len(create_text_requests)} create text requests")
                    slides_service.presentations().batchUpdate(
                        presentationId=presentation_id,
                        body={'requests': create_text_requests}
                    ).execute()
                    print("Successfully created new text elements")
                except Exception as e:
                    print(f"Warning: Failed to create new text elements: {str(e)}")
        
        # Execute update requests
        if update_requests:
            try:
                print(f"Executing {len(update_requests)} update requests")
                slides_service.presentations().batchUpdate(
                    presentationId=presentation_id,
                    body={'requests': update_requests}
                ).execute()
                print("Successfully executed text updates")
            except Exception as e:
                print(f"Warning: Failed to update text: {str(e)}")
                import traceback
                traceback.print_exc()
                
                # If the update failed, try a simpler approach with one request at a time
                print("Trying individual updates as fallback...")
                for i in range(0, len(update_requests), 2):
                    if i+1 < len(update_requests):
                        try:
                            # Execute a delete+insert pair
                            slides_service.presentations().batchUpdate(
                                presentationId=presentation_id,
                                body={'requests': [update_requests[i], update_requests[i+1]]}
                            ).execute()
                            print(f"Successfully updated field {i//2 + 1}")
                        except Exception as e2:
                            print(f"Failed to update field {i//2 + 1}: {str(e2)}")
        else:
            print("WARNING: No update requests were generated")
    
    except Exception as e:
        print(f"ERROR in direct_update_text_on_slide: {str(e)}")
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

# Legacy function for backward compatibility - no longer used
def update_order_details(slides_service, presentation_id, slide_id, order):
    """Legacy function that's kept for backward compatibility but replaced by direct_update_text_on_slide"""
    # This is now just a wrapper around the new function
    return direct_update_text_on_slide(slides_service, presentation_id, slide_id, order)

# Legacy function for backward compatibility - no longer used
def find_table_cells(slides_service, presentation_id, slide_id):
    """Legacy function - kept for backward compatibility but no longer used"""
    print("Warning: find_table_cells is deprecated and will not work properly")
    return {}

# Legacy function for backward compatibility - no longer used
def update_text_fields(slides_service, presentation_id, text_fields, order):
    """Legacy function - kept for backward compatibility but no longer used"""
    print("Warning: update_text_fields is deprecated and will not work properly")
    return