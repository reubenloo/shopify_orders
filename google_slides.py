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
                
                # Wait briefly to ensure the slide is fully created
                time.sleep(0.5)
                
                # Now update the order details on this slide with a specialized approach for table-based slides
                update_table_based_slide(slides_service, presentation_id, new_slide_id, order)
                
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

def update_table_based_slide(slides_service, presentation_id, slide_id, order):
    """
    Specialized function to update text in a table-based slide format
    
    Args:
        slides_service: Google Slides API service
        presentation_id: ID of the presentation
        slide_id: ID of the slide to update
        order: Dictionary containing order information
    """
    try:
        print(f"Updating table-based slide {slide_id} for order: {order.get('order_number', 'unknown')}")
        
        # Get the slide structure
        slide = slides_service.presentations().pages().get(
            presentationId=presentation_id,
            pageObjectId=slide_id
        ).execute()
        
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
        
        # Get all text elements on the slide
        all_text_elements = []
        for element in slide.get('pageElements', []):
            # Check for shapes with text
            if 'shape' in element and 'text' in element.get('shape', {}):
                shape_id = element.get('objectId')
                text_content = ""
                for text_element in element.get('shape', {}).get('text', {}).get('textElements', []):
                    if 'textRun' in text_element:
                        text_content += text_element.get('textRun', {}).get('content', '')
                
                all_text_elements.append({
                    'id': shape_id,
                    'type': 'shape',
                    'text': text_content.strip(),
                    'x': element.get('transform', {}).get('translateX', 0),
                    'y': element.get('transform', {}).get('translateY', 0)
                })
                print(f"Found shape text: {shape_id} at ({element.get('transform', {}).get('translateX', 0)}, {element.get('transform', {}).get('translateY', 0)}): \"{text_content.strip()}\"")
            
            # Check for table cells with text
            elif 'table' in element:
                table = element.get('table')
                table_id = element.get('objectId')
                print(f"Found table: {table_id} with {len(table.get('tableRows', []))} rows, {len(table.get('tableRows', [])[0].get('tableCells', []) if table.get('tableRows') else [])} columns")
                
                for row_idx, row in enumerate(table.get('tableRows', [])):
                    for col_idx, cell in enumerate(row.get('tableCells', [])):
                        cell_text = ""
                        cell_object_id = cell.get('objectId')
                        
                        # Extract text from this cell
                        if 'text' in cell:
                            for text_element in cell.get('text', {}).get('textElements', []):
                                if 'textRun' in text_element:
                                    cell_text += text_element.get('textRun', {}).get('content', '')
                        
                        all_text_elements.append({
                            'id': cell_object_id,
                            'type': 'table_cell',
                            'table_id': table_id,
                            'row': row_idx,
                            'col': col_idx,
                            'text': cell_text.strip()
                        })
                        print(f"Found table cell: {cell_object_id} at row={row_idx}, col={col_idx}: \"{cell_text.strip()}\"")
        
        # First try to find elements based on exact row/column positions
        left_side_elements = []
        name_element = next((e for e in all_text_elements if e['type'] == 'table_cell' and e['row'] == 1 and e['col'] == 0), None)
        contact_element = next((e for e in all_text_elements if e['type'] == 'table_cell' and e['row'] == 2 and e['col'] == 0), None)
        address_element = next((e for e in all_text_elements if e['type'] == 'table_cell' and e['row'] == 3 and e['col'] == 0), None)
        postal_element = next((e for e in all_text_elements if e['type'] == 'table_cell' and e['row'] == 4 and e['col'] == 0), None)
        item_element = next((e for e in all_text_elements if e['type'] == 'table_cell' and e['row'] == 5 and e['col'] == 0), None)
        
        # If we couldn't find elements by position, try by content
        if not name_element:
            name_element = next((e for e in all_text_elements if "NAME:" in e['text'].upper() or "NAME #" in e['text'].upper()), None)
        if not contact_element:
            contact_element = next((e for e in all_text_elements if "CONTACT:" in e['text'].upper() and "ECZEMA" not in e['text'].upper()), None)
        if not address_element:
            address_element = next((e for e in all_text_elements if "DELIVERY ADDRESS:" in e['text'].upper()), None)
        if not postal_element:
            postal_element = next((e for e in all_text_elements if "POSTAL:" in e['text'].upper() and "680235" not in e['text'].upper()), None)
        if not item_element:
            item_element = next((e for e in all_text_elements if "ITEM:" in e['text'].upper()), None)
        
        # Try a third approach - identify all text elements on the left side
        left_side_threshold = 200  # Adjust based on your template
        left_side_elements = [e for e in all_text_elements if e['type'] == 'shape' and e['x'] < left_side_threshold]
        left_side_elements.sort(key=lambda e: e['y'])  # Sort by vertical position
        
        # Prepare update requests
        update_requests = []
        
        # Function to add update request for an element if it exists
        def add_update_for_element(element, content):
            if element and element.get('id'):
                print(f"Adding update for element {element['id']} with content: \"{content}\"")
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
                            'text': content
                        }
                    }
                ])
                return True
            return False
        
        # Try to update elements with specific content
        updated_name = add_update_for_element(name_element, f"Name: #{order.get('order_number', '').replace('#', '')} {order.get('name', '')}")
        updated_contact = add_update_for_element(contact_element, f"Contact: {order.get('phone', '')}")
        updated_address = add_update_for_element(address_element, f"Delivery Address: {address}")
        updated_postal = add_update_for_element(postal_element, f"Postal: {order.get('postal', '')}")
        updated_item = add_update_for_element(item_element, f"Item: {quantity} {size_display} {material} Eczema Mitten")
        
        # If we couldn't find specific elements, try using left side elements by position
        if len(left_side_elements) >= 5:
            if not updated_name and len(left_side_elements) > 0:
                add_update_for_element(left_side_elements[0], f"Name: #{order.get('order_number', '').replace('#', '')} {order.get('name', '')}")
            if not updated_contact and len(left_side_elements) > 1:
                add_update_for_element(left_side_elements[1], f"Contact: {order.get('phone', '')}")
            if not updated_address and len(left_side_elements) > 2:
                add_update_for_element(left_side_elements[2], f"Delivery Address: {address}")
            if not updated_postal and len(left_side_elements) > 3:
                add_update_for_element(left_side_elements[3], f"Postal: {order.get('postal', '')}")
            if not updated_item and len(left_side_elements) > 4:
                add_update_for_element(left_side_elements[4], f"Item: {quantity} {size_display} {material} Eczema Mitten")
        
        # Execute updates
        if update_requests:
            print(f"Executing {len(update_requests)} update requests")
            try:
                slides_service.presentations().batchUpdate(
                    presentationId=presentation_id,
                    body={'requests': update_requests}
                ).execute()
                print("Successfully executed updates")
            except Exception as e:
                print(f"WARNING: Error executing batch update: {str(e)}")
                
                # Try updating one pair at a time
                print("Trying individual updates...")
                for i in range(0, len(update_requests), 2):
                    if i+1 < len(update_requests):
                        try:
                            slides_service.presentations().batchUpdate(
                                presentationId=presentation_id,
                                body={'requests': [update_requests[i], update_requests[i+1]]}
                            ).execute()
                            print(f"Successfully updated pair {i//2 + 1}")
                        except Exception as e2:
                            print(f"Failed to update pair {i//2 + 1}: {str(e2)}")
        else:
            print("WARNING: No updates prepared for this slide - no suitable elements found")
            
            # If we couldn't find any elements to update, try creating a text note about it
            # This creates a visible notice on the slide that something went wrong
            try:
                note_requests = [{
                    'createShape': {
                        'objectId': f'note_{slide_id}',
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
                                'translateX': 50,
                                'translateY': 50,
                                'unit': 'PT'
                            }
                        }
                    }
                }, {
                    'insertText': {
                        'objectId': f'note_{slide_id}',
                        'insertionIndex': 0,
                        'text': f"Order #{order.get('order_number', '').replace('#', '')}: {order.get('name', '')}\nContact: {order.get('phone', '')}\nItem: {quantity} {size_display} {material}"
                    }
                }]
                
                slides_service.presentations().batchUpdate(
                    presentationId=presentation_id,
                    body={'requests': note_requests}
                ).execute()
                print("Created fallback note with order information")
            except Exception as note_err:
                print(f"Failed to create fallback note: {str(note_err)}")
    
    except Exception as e:
        print(f"ERROR updating table-based slide: {str(e)}")
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
    """Legacy function that redirects to the new table-based approach"""
    return update_table_based_slide(slides_service, presentation_id, slide_id, order)

def find_table_cells(slides_service, presentation_id, slide_id):
    """Legacy function - kept for backward compatibility but no longer used"""
    print("Warning: find_table_cells is deprecated and will not work properly")
    return {}

def update_text_fields(slides_service, presentation_id, text_fields, order):
    """Legacy function - kept for backward compatibility but no longer used"""
    print("Warning: update_text_fields is deprecated and will not work properly")
    return

def direct_update_text_on_slide(slides_service, presentation_id, slide_id, order):
    """Legacy function that redirects to the new table-based approach"""
    return update_table_based_slide(slides_service, presentation_id, slide_id, order)