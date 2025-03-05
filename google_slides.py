import os
from googleapiclient.discovery import build
from google.oauth2 import service_account
from datetime import datetime
import re

def create_shipping_slides(order_details, credentials_path, template_id=None):
    """
    Create or update a Google Slides presentation with shipping labels for orders
    
    Args:
        order_details: List of dictionaries containing order information
        credentials_path: Path to the service account JSON credentials file
        template_id: Optional ID of a template presentation to use
        
    Returns:
        presentation_url: URL of the created/updated presentation
    """
    try:
        # Set up credentials
        SCOPES = ['https://www.googleapis.com/auth/presentations', 
                  'https://www.googleapis.com/auth/drive']
        
        credentials = service_account.Credentials.from_service_account_file(
            credentials_path, scopes=SCOPES)
        
        # Create services
        slides_service = build('slides', 'v1', credentials=credentials)
        drive_service = build('drive', 'v3', credentials=credentials)
        
        # If template_id is provided, we'll directly modify that presentation
        if template_id:
            presentation_id = template_id
            presentation_url = f"https://docs.google.com/presentation/d/{presentation_id}/edit"
            
            # Get the presentation details to check existing slides
            presentation = slides_service.presentations().get(
                presentationId=presentation_id
            ).execute()
        else:
            # Create a new blank presentation if no template is provided
            presentation = {
                'title': f"Shipping Labels - {datetime.now().strftime('%Y-%m-%d %H:%M')}"
            }
            presentation = slides_service.presentations().create(body=presentation).execute()
            presentation_id = presentation.get('presentationId')
            presentation_url = f"https://docs.google.com/presentation/d/{presentation_id}/edit"
        
        # Get the current slides
        presentation = slides_service.presentations().get(
            presentationId=presentation_id
        ).execute()
        
        slides = presentation.get('slides', [])
        
        # Check if we have at least 2 slides (date slide and template slide)
        if len(slides) < 2:
            return None  # Not enough slides in the template
        
        # Get the template slide (2nd slide)
        template_slide_id = slides[1].get('objectId')
        
        # Create a new date slide for this batch
        current_date = datetime.now().strftime('%d %b, %Y')
        
        # Create a new date slide at the beginning
        date_slide_request = {
            'createSlide': {
                'insertionIndex': 0,  # Insert at the beginning
                'slideLayoutReference': {
                    'predefinedLayout': 'TITLE_ONLY'
                },
                'placeholderIdMappings': [{
                    'layoutPlaceholder': {
                        'type': 'TITLE',
                        'index': 0
                    },
                    'objectId': 'dateTitle'
                }]
            }
        }
        
        # Create the date slide
        date_slide_response = slides_service.presentations().batchUpdate(
            presentationId=presentation_id,
            body={'requests': [date_slide_request]}
        ).execute()
        
        # Get the ID of the new date slide
        new_date_slide_id = date_slide_response.get('replies', [{}])[0].get('createSlide', {}).get('objectId')
        
        # Add the date to the title placeholder
        if new_date_slide_id:
            title_text_request = {
                'insertText': {
                    'objectId': 'dateTitle',
                    'insertionIndex': 0,
                    'text': current_date
                }
            }
            
            slides_service.presentations().batchUpdate(
                presentationId=presentation_id,
                body={'requests': [title_text_request]}
            ).execute()
        
        # Process each order - create new slides based on the template
        insert_index = 1  # Start inserting after the date slide
        
        for order in order_details:
            # Duplicate the template slide
            duplicate_request = {
                'duplicateObject': {
                    'objectId': template_slide_id,
                    'insertionIndex': insert_index
                }
            }
            
            duplicate_response = slides_service.presentations().batchUpdate(
                presentationId=presentation_id,
                body={'requests': [duplicate_request]}
            ).execute()
            
            # Get the ID of the duplicated slide
            new_slide_id = duplicate_response.get('replies', [{}])[0].get('duplicateObject', {}).get('objectId')
            
            # Get the elements of the new slide, particularly the table
            slide = slides_service.presentations().pages().get(
                presentationId=presentation_id,
                pageObjectId=new_slide_id
            ).execute()
            
            # Find the table in the new slide
            table_id = None
            cell_ids = []
            
            for element in slide.get('pageElements', []):
                if 'table' in element:
                    table_id = element.get('objectId')
                    
                    rows = element.get('table', {}).get('tableRows', [])
                    for row_idx, row in enumerate(rows):
                        for col_idx, cell in enumerate(row.get('tableCells', [])):
                            cell_ids.append({
                                'row': row_idx,
                                'column': col_idx,
                                'objectId': cell.get('objectId')
                            })
            
            if not table_id:
                print(f"Warning: No table found in slide for order {order.get('order_number', '')}")
                continue
            
            # Prepare the content for each cell
            quantity = "2" if order.get('is_bundle', False) else "1"
            size = order.get('size', '')
            material = order.get('material', '')
            
            # Format the size for kid sizes
            if '(' in size and 'cm' in size:
                size_display = size.split('(')[1].replace(')', '').split('-')[0] + 'cm'
            else:
                size_display = size
                
            # Combine address lines
            address1 = order.get('address1', '')
            address2 = order.get('address2', '')
            address = f"{address1}\n{address2}" if address2 and address2.strip() else address1
            
            # For the items cell (merged cell)
            item_description = f"{quantity} {size_display} {material} Eczema Mitten"
            
            # Update each cell in the table
            text_updates = []
            
            for cell in cell_ids:
                row = cell.get('row')
                col = cell.get('column')
                
                # Define the content based on row and column
                content = ""
                
                # Left column (col=0)
                if col == 0:
                    if row == 0:  # Name field
                        content = f"Name: #{order.get('order_number', '').replace('#', '')} {order.get('name', '')}"
                    elif row == 1:  # Contact field
                        content = f"Contact: {order.get('phone', '')}"
                    elif row == 2:  # Delivery Address field
                        content = f"Delivery Address: {address}"
                    elif row == 3:  # Postal field
                        content = f"Postal: {order.get('postal', '')}"
                    elif row == 4:  # Item field (merged cell)
                        content = f"Item: {item_description}"
                
                # Right column (col=1) - Company info is already in the template
                
                # Only update if we have content to set
                if content:
                    # Clear existing text in the cell
                    text_updates.append({
                        'deleteText': {
                            'objectId': cell.get('objectId'),
                            'textRange': {
                                'type': 'ALL'
                            }
                        }
                    })
                    
                    # Insert new text
                    text_updates.append({
                        'insertText': {
                            'objectId': cell.get('objectId'),
                            'insertionIndex': 0,
                            'text': content
                        }
                    })
            
            # Apply all text updates
            if text_updates:
                slides_service.presentations().batchUpdate(
                    presentationId=presentation_id,
                    body={'requests': text_updates}
                ).execute()
            
            # Increment insert index for next slide
            insert_index += 1
        
        return presentation_url
        
    except Exception as e:
        print(f"Error creating Google Slides: {str(e)}")
        return None

def get_template_id_from_url(url):
    """Extract the presentation ID from a Google Slides URL"""
    match = re.search(r'/d/([a-zA-Z0-9-_]+)', url)
    if match:
        return match.group(1)
    return None