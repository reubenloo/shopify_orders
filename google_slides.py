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
            print(f"Created new date slide with ID: {new_date_slide_id}")
            
            # Update the date on the new slide
            update_date_slide(slides_service, presentation_id, new_date_slide_id)
            
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
            
            # Create order slides
            new_order_slide_ids = []
            
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
                if not new_slide_id:
                    print(f"WARNING: Could not get ID for the new order slide {i+1}")
                    continue
                    
                new_order_slide_ids.append(new_slide_id)
                print(f"Created new order slide with ID: {new_slide_id}")
                
                # Update the order details on this slide
                update_order_slide_with_table(slides_service, presentation_id, new_slide_id, order)
            
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
                return  # Only update the first text element
                
        print("WARNING: No text elements found on date slide")
        
    except Exception as e:
        print(f"ERROR updating date slide: {str(e)}")
        import traceback
        traceback.print_exc()

def update_order_slide_with_table(slides_service, presentation_id, slide_id, order):
    """
    Update a slide's table with order information
    
    Args:
        slides_service: Google Slides API service
        presentation_id: ID of the presentation
        slide_id: ID of the slide to update
        order: Dictionary containing order information
    """
    print(f"Updating slide {slide_id} with order details for {order.get('order_number')}")
    
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
            print(f"WARNING: Could not find slide {slide_id}")
            return
        
        # Find the table on the slide
        for element in target_slide.get('pageElements', []):
            if 'table' in element:
                table_id = element.get('objectId')
                table = element.get('table', {})
                rows = table.get('tableRows', [])
                
                print(f"Found table with {len(rows)} rows")
                
                # Prepare the update requests
                update_requests = []
                
                # Table cells we want to update by row and column index (0-based)
                # Format: [row, column, field_name, new_value]
                cells_to_update = [
                    # Name: row 1, column 0
                    [1, 0, "Name field", f"Name: #{order.get('order_number')} {order.get('name', '')}"],
                    
                    # Contact: row 2, column 0
                    [2, 0, "Contact field", f"Contact: {order.get('phone', 'N/A')}"],
                    
                    # Delivery Address: row 3, column 0
                    [3, 0, "Address field", f"Delivery Address:\n{order.get('address1', '')}\n{order.get('address2', '')}"],
                    
                    # Postal: row 4, column 0
                    [4, 0, "Postal field", f"Postal: {order.get('postal', '')}"],
                    
                    # Item: row 5, column 0 and column 1 (merged cell)
                    [5, 0, "Item field", f"Item: {2 if order.get('is_bundle') else 1} {order.get('size', '')} {order.get('material', '')} Eczema Mitten"]
                ]
                
                # Process each cell
                for row_idx, col_idx, field_name, new_value in cells_to_update:
                    # Make sure the row and column exist
                    if row_idx < len(rows):
                        row = rows[row_idx]
                        cells = row.get('tableCells', [])
                        
                        if col_idx < len(cells):
                            cell = cells[col_idx]
                            
                            # Find the text elements in the cell
                            for text_element in cell.get('text', {}).get('textElements', []):
                                if 'textRun' in text_element:
                                    # Get the cellLocation for this cell
                                    cell_location = {
                                        'tableObjectId': table_id,
                                        'rowIndex': row_idx,
                                        'columnIndex': col_idx
                                    }
                                    
                                    # Create requests to delete existing text and insert new text
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
                                                'text': new_value
                                            }
                                        }
                                    ])
                                    
                                    print(f"Will update {field_name} to: {new_value}")
                                    break  # Only process the first text element in the cell
                
                # Submit all updates
                if update_requests:
                    print(f"Submitting {len(update_requests)} update requests...")
                    slides_service.presentations().batchUpdate(
                        presentationId=presentation_id,
                        body={'requests': update_requests}
                    ).execute()
                    print("Successfully updated table cells with order details")
                else:
                    print("WARNING: No table cells were identified for update")
                
                return  # Only process the first table found
        
        print("WARNING: No table found on the slide")
        
    except Exception as e:
        print(f"ERROR updating order slide with table: {str(e)}")
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