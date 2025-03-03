import os
import json
from googleapiclient.discovery import build
from google.oauth2 import service_account
from datetime import datetime, timedelta
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
        tuple: (presentation_url, pdf_path) URL of the edited presentation and path to PDF
    """
    debug_log = []
    
    def log_debug(message):
        print(message)
        debug_log.append(message)
    
    try:
        # Print detailed information for debugging
        log_debug(f"Starting create_shipping_slides with {len(order_details)} orders")
        log_debug(f"Credentials path: {credentials_path}")
        log_debug(f"File exists: {os.path.exists(credentials_path)}")
        log_debug(f"Template ID: {template_id or 'None'}")
        
        # Read and print first part of credentials file to verify it's valid
        try:
            with open(credentials_path, 'r') as f:
                cred_content = f.read()
                # Print first 100 chars with sensitive info masked
                masked_content = cred_content[:100].replace('"private_key":', '"private_key": "[MASKED]",')
                log_debug(f"Credentials file content (first 100 chars): {masked_content}...")
                
                # Validate it's proper JSON
                try:
                    cred_json = json.loads(cred_content)
                    required_fields = ['type', 'project_id', 'private_key', 'client_email']
                    missing_fields = [field for field in required_fields if field not in cred_json]
                    if missing_fields:
                        log_debug(f"WARNING: Credentials file is missing required fields: {missing_fields}")
                    else:
                        log_debug("Credentials file contains all required fields")
                except json.JSONDecodeError as e:
                    log_debug(f"ERROR: Credentials file is not valid JSON: {str(e)}")
                    return None, None
        except Exception as e:
            log_debug(f"ERROR: Could not read credentials file: {str(e)}")
            return None, None
        
        # Set up credentials with detailed error handling
        SCOPES = ['https://www.googleapis.com/auth/presentations', 
                 'https://www.googleapis.com/auth/drive']
        
        try:
            log_debug("Creating credentials object...")
            credentials = service_account.Credentials.from_service_account_file(
                credentials_path, scopes=SCOPES)
            log_debug(f"Credentials created successfully: {credentials.service_account_email}")
        except Exception as e:
            log_debug(f"ERROR creating credentials: {str(e)}")
            import traceback
            log_debug(traceback.format_exc())
            return None, None
        
        # Create services with detailed error handling
        try:
            log_debug("Building slides service...")
            slides_service = build('slides', 'v1', credentials=credentials)
            log_debug("Building drive service...")
            drive_service = build('drive', 'v3', credentials=credentials)
            log_debug("Services built successfully")
        except Exception as e:
            log_debug(f"ERROR building services: {str(e)}")
            import traceback
            log_debug(traceback.format_exc())
            return None, None
        
        # Use the existing presentation if template_id is provided
        presentation_id = None
        presentation_url = None
        
        try:
            if template_id:
                log_debug(f"Using existing presentation: {template_id}")
                
                # Make a copy of the template for this session
                copy_title = f"Shipping Labels - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
                drive_response = drive_service.files().copy(
                    fileId=template_id,
                    body={"name": copy_title}
                ).execute()
                
                presentation_id = drive_response.get('id')
                log_debug(f"Created copy of template with ID: {presentation_id}")
            else:
                log_debug("No template ID provided, cannot proceed")
                return None, None
                
            presentation_url = f"https://docs.google.com/presentation/d/{presentation_id}/edit"
            log_debug(f"Presentation URL: {presentation_url}")
        except Exception as e:
            log_debug(f"ERROR accessing presentation: {str(e)}")
            import traceback
            log_debug(traceback.format_exc())
            return None, None
        
        # Get current presentation details to understand slide layout
        try:
            presentation = slides_service.presentations().get(
                presentationId=presentation_id
            ).execute()
            log_debug(f"Fetched presentation details, title: {presentation.get('title')}")
            
            # Get existing slides
            slides = presentation.get('slides', [])
            log_debug(f"Presentation has {len(slides)} existing slides")
            
            # Ensure we have at least 2 slides (first for date, second for template)
            if len(slides) < 2:
                log_debug("ERROR: Template presentation should have at least 2 slides")
                return presentation_url, None
                
            # Save the date slide and template slide
            date_slide_id = slides[0].get('objectId')
            template_slide_id = slides[1].get('objectId')
            
            log_debug(f"Found date slide with ID: {date_slide_id}")
            log_debug(f"Found template slide with ID: {template_slide_id}")
            
            # Analyze the template slide structure
            template_slide = slides[1]
            analyze_slide_structure(slides_service, presentation_id, template_slide, log_debug)
            
            # Create a new date slide by duplicating the first slide
            date_request = [{
                'duplicateObject': {
                    'objectId': date_slide_id,
                }
            }]
            
            date_response = slides_service.presentations().batchUpdate(
                presentationId=presentation_id,
                body={'requests': date_request}
            ).execute()
            
            new_date_slide_id = date_response.get('replies', [{}])[0].get('duplicateObject', {}).get('objectId')
            log_debug(f"Created new date slide with ID: {new_date_slide_id}")
            
            # Update the date on the new slide
            update_date_slide(slides_service, presentation_id, new_date_slide_id, log_debug)
            
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
            log_debug("Moved date slide to the beginning")
            
            # Create order slides
            new_order_slide_ids = []
            
            for i, order in enumerate(order_details):
                log_debug(f"Creating slide for order {i+1}: {order.get('order_number')}")
                
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
                if not new_slide_id:
                    log_debug(f"WARNING: Could not get ID for the new order slide {i+1}")
                    continue
                    
                new_order_slide_ids.append(new_slide_id)
                log_debug(f"Created new order slide with ID: {new_slide_id}")
                
                # Update the order details on this slide
                update_order_slide(slides_service, presentation_id, new_slide_id, order, log_debug)
            
            # Move all the new order slides after the date slide
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
                
            log_debug(f"Successfully repositioned {len(new_order_slide_ids)} order slides")
            
            # Save debug log to a file
            with open("google_slides_debug.log", "w") as f:
                f.write("\n".join(debug_log))
            
            # No PDF generation in this version
            pdf_path = None
            
            return presentation_url, pdf_path
            
        except Exception as e:
            log_debug(f"ERROR in slide creation: {str(e)}")
            import traceback
            log_debug(traceback.format_exc())
            
            # Save debug log to a file even on error
            with open("google_slides_debug.log", "w") as f:
                f.write("\n".join(debug_log))
                
            return presentation_url, None
        
    except Exception as e:
        print(f"ERROR in create_shipping_slides: {str(e)}")
        import traceback
        traceback.print_exc()
        
        # Try to save debug log if possible
        try:
            with open("google_slides_debug.log", "w") as f:
                f.write("\n".join(debug_log))
        except:
            pass
            
        return None, None

def analyze_slide_structure(slides_service, presentation_id, slide, log_debug):
    """
    Analyze the structure of a slide to understand its elements
    
    Args:
        slides_service: Google Slides API service
        presentation_id: ID of the presentation
        slide: The slide object to analyze
        log_debug: Function for logging debug information
    """
    log_debug("SLIDE STRUCTURE ANALYSIS:")
    
    # Log slide ID and title if available
    slide_id = slide.get('objectId')
    log_debug(f"Slide ID: {slide_id}")
    
    # Count and categorize elements
    elements = slide.get('pageElements', [])
    log_debug(f"Total page elements: {len(elements)}")
    
    element_types = {}
    for i, element in enumerate(elements):
        element_id = element.get('objectId')
        element_type = next(iter([k for k in element.keys() if k not in ['objectId', 'transform']]), 'unknown')
        
        if element_type not in element_types:
            element_types[element_type] = []
        element_types[element_type].append(element_id)
        
        log_debug(f"Element {i+1}: ID={element_id}, Type={element_type}")
        
        # If it's a table, analyze its structure
        if element_type == 'table':
            table = element.get('table', {})
            rows = table.get('tableRows', [])
            columns = table.get('tableColumns', [])
            
            log_debug(f"  Table has {len(rows)} rows and {len(columns)} columns")
            
            # Check each cell of the table
            for row_idx, row in enumerate(rows):
                cells = row.get('tableCells', [])
                log_debug(f"  Row {row_idx+1}: {len(cells)} cells")
                
                for col_idx, cell in enumerate(cells):
                    text_elements = cell.get('text', {}).get('textElements', [])
                    
                    has_text = False
                    for text_el in text_elements:
                        if 'textRun' in text_el:
                            content = text_el.get('textRun', {}).get('content', '')
                            if content.strip():
                                has_text = True
                                log_debug(f"    Cell [{row_idx},{col_idx}]: '{content.strip()}'")
                    
                    # Create a placeholder for empty cells too
                    if not has_text:
                        log_debug(f"    Cell [{row_idx},{col_idx}]: [EMPTY]")
        
        # If it's a shape with text, show the text
        elif element_type == 'shape' and 'text' in element.get('shape', {}):
            text_elements = element.get('shape', {}).get('text', {}).get('textElements', [])
            for text_el in text_elements:
                if 'textRun' in text_el:
                    content = text_el.get('textRun', {}).get('content', '')
                    if content.strip():
                        log_debug(f"  Shape text: '{content.strip()}'")
    
    # Summary
    log_debug("Element Type Summary:")
    for element_type, ids in element_types.items():
        log_debug(f"  {element_type}: {len(ids)} elements")

def update_date_slide(slides_service, presentation_id, slide_id, log_debug):
    """
    Update the date on a slide to today's date
    
    Args:
        slides_service: Google Slides API service
        presentation_id: ID of the presentation
        slide_id: ID of the date slide
        log_debug: Function for logging debug information
    """
    log_debug(f"Updating date on slide {slide_id}...")
    
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
            log_debug(f"WARNING: Could not find slide {slide_id}")
            return
            
        # Find and update the text elements
        found_text_element = False
        for element in target_slide.get('pageElements', []):
            if 'shape' in element and 'text' in element.get('shape', {}):
                # Found a text element, update it with today's date
                element_id = element.get('objectId')
                log_debug(f"Found text element with ID: {element_id}")
                
                # Check if it contains any text elements
                text_elements = element.get('shape', {}).get('text', {}).get('textElements', [])
                log_debug(f"Text element contains {len(text_elements)} text sub-elements")
                
                for text_el in text_elements:
                    if 'textRun' in text_el:
                        content = text_el.get('textRun', {}).get('content', '')
                        log_debug(f"Current text content: '{content}'")
                
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
                
                log_debug(f"Updated date to: {today}")
                found_text_element = True
                break  # Only update the first text element
        
        if not found_text_element:
            log_debug("WARNING: No suitable text elements found on date slide")
        
    except Exception as e:
        log_debug(f"ERROR updating date slide: {str(e)}")
        import traceback
        log_debug(traceback.format_exc())

def update_order_slide(slides_service, presentation_id, slide_id, order, log_debug):
    """
    Update a slide with order information (version that adapts based on slide structure)
    
    Args:
        slides_service: Google Slides API service
        presentation_id: ID of the presentation
        slide_id: ID of the slide to update
        order: Dictionary containing order information
        log_debug: Function for logging debug information
    """
    log_debug(f"Updating slide {slide_id} with order details for {order.get('order_number')}")
    
    # Add a small delay to ensure the slide is fully created
    time.sleep(0.5)
    
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
            log_debug(f"WARNING: Could not find slide {slide_id}")
            return
        
        # First try to update using tables
        tables_updated = update_slide_tables(slides_service, presentation_id, slide_id, target_slide, order, log_debug)
        
        # If no tables were updated, try updating shapes
        if not tables_updated:
            log_debug("No tables updated, trying to update shapes instead...")
            update_slide_shapes(slides_service, presentation_id, slide_id, target_slide, order, log_debug)
        
    except Exception as e:
        log_debug(f"ERROR updating order slide: {str(e)}")
        import traceback
        log_debug(traceback.format_exc())

def update_slide_tables(slides_service, presentation_id, slide_id, slide, order, log_debug):
    """
    Update tables on a slide with order information
    
    Args:
        slides_service: Google Slides API service
        presentation_id: ID of the presentation
        slide_id: ID of the slide
        slide: The slide object to update
        order: Dictionary containing order information
        log_debug: Function for logging debug information
        
    Returns:
        bool: True if at least one table was updated
    """
    log_debug("Attempting to update tables on the slide...")
    
    tables_updated = False
    
    for element in slide.get('pageElements', []):
        if 'table' in element:
            table_id = element.get('objectId')
            table = element.get('table', {})
            rows = table.get('tableRows', [])
            
            log_debug(f"Found table with ID {table_id} containing {len(rows)} rows")
            
            # First approach: Try direct cell-by-cell update based on row/column indices
            update_requests = []
            
            # Map expected content to specific cells (row, column)
            # Order: row_index, column_index, content
            cell_content_map = [
                # Direct approach that doesn't rely on existing text
                # Row 1 (second row, index 1): Customer name (assuming header is first row, index 0)
                [1, 0, f"Name: #{order.get('order_number')} {order.get('name', '')}"],
                
                # Row 2 (third row): Customer contact
                [2, 0, f"Contact: {order.get('phone', 'N/A')}"],
                
                # Row 3 (fourth row): Delivery address
                [3, 0, f"Delivery Address:\n{order.get('address1', '')}\n{order.get('address2', '')}"],
                
                # Row 4 (fifth row): Postal code
                [4, 0, f"Postal: {order.get('postal', '')}"],
                
                # Row 5 (sixth row): Item details
                [5, 0, f"Item: {2 if order.get('is_bundle') else 1} {order.get('size', '')} {order.get('material', '')} Eczema Mitten"]
            ]
            
            # Add each cell update request
            for row_idx, col_idx, content in cell_content_map:
                # Make sure the row and column exist
                if row_idx < len(rows):
                    row = rows[row_idx]
                    cells = row.get('tableCells', [])
                    
                    if col_idx < len(cells):
                        log_debug(f"Creating request to update cell [{row_idx},{col_idx}] with: {content}")
                        
                        # Create a cell location
                        cell_location = {
                            'tableObjectId': table_id,
                            'rowIndex': row_idx,
                            'columnIndex': col_idx
                        }
                        
                        # Delete existing text and insert new text
                        update_requests.extend([
                            {
                                'deleteText': {
                                    'objectId': table_id,
                                    'cellLocation': cell_location,
                                    'textRange': {
                                        'type': 'ALL'
                                    }
                                }
                            },
                            {
                                'insertText': {
                                    'objectId': table_id,
                                    'cellLocation': cell_location,
                                    'text': content
                                }
                            }
                        ])
            
            # Submit all updates
            if update_requests:
                log_debug(f"Submitting {len(update_requests)} table cell update requests...")
                try:
                    slides_service.presentations().batchUpdate(
                        presentationId=presentation_id,
                        body={'requests': update_requests}
                    ).execute()
                    log_debug("Successfully submitted table cell updates")
                    tables_updated = True
                except Exception as e:
                    log_debug(f"Error updating table cells: {str(e)}")
            else:
                log_debug("No table cell update requests were created")
            
    return tables_updated

def update_slide_shapes(slides_service, presentation_id, slide_id, slide, order, log_debug):
    """
    Update shapes on a slide with order information
    
    Args:
        slides_service: Google Slides API service
        presentation_id: ID of the presentation
        slide_id: ID of the slide
        slide: The slide object to update
        order: Dictionary containing order information
        log_debug: Function for logging debug information
    """
    log_debug("Attempting to update shapes on the slide...")
    
    shapes_updated = 0
    
    for element in slide.get('pageElements', []):
        if 'shape' in element and 'text' in element.get('shape', {}):
            element_id = element.get('objectId')
            
            # Get existing text to determine what to replace
            text_elements = element.get('shape', {}).get('text', {}).get('textElements', [])
            current_text = ""
            for text_el in text_elements:
                if 'textRun' in text_el:
                    current_text += text_el.get('textRun', {}).get('content', '')
            
            log_debug(f"Found shape with ID {element_id} containing text: '{current_text.strip()}'")
            
            # Determine what content to put in this shape based on the text it contains
            new_text = None
            
            # Simple keyword matching approach
            current_text_lower = current_text.lower()
            if "name" in current_text_lower:
                new_text = f"Name: #{order.get('order_number')} {order.get('name', '')}"
            elif "contact" in current_text_lower or "phone" in current_text_lower:
                new_text = f"Contact: {order.get('phone', 'N/A')}"
            elif "address" in current_text_lower:
                new_text = f"Delivery Address:\n{order.get('address1', '')}\n{order.get('address2', '')}"
            elif "postal" in current_text_lower:
                new_text = f"Postal: {order.get('postal', '')}"
            elif "item" in current_text_lower:
                new_text = f"Item: {2 if order.get('is_bundle') else 1} {order.get('size', '')} {order.get('material', '')} Eczema Mitten"
            
            # If we determined what to put in this shape
            if new_text:
                log_debug(f"Updating shape with: {new_text}")
                
                try:
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
                            'text': new_text
                        }
                    }]
                    
                    slides_service.presentations().batchUpdate(
                        presentationId=presentation_id,
                        body={'requests': requests}
                    ).execute()
                    
                    log_debug(f"Successfully updated shape with text: {new_text}")
                    shapes_updated += 1
                except Exception as e:
                    log_debug(f"Error updating shape: {str(e)}")
    
    log_debug(f"Updated {shapes_updated} shapes on the slide")

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