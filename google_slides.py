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
        template_id: Optional ID of a template presentation to copy
        
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
        
        # Create a new presentation or use template
        presentation_id = None
        presentation_url = None
        
        if template_id:
            # Copy the template presentation
            copy_title = f"Shipping Labels - {datetime.now().strftime('%Y-%m-%d %H:%M')}"
            drive_response = drive_service.files().copy(
                fileId=template_id,
                body={"name": copy_title}
            ).execute()
            presentation_id = drive_response.get('id')
            
            # Get the presentation details
            presentation = slides_service.presentations().get(
                presentationId=presentation_id
            ).execute()
            
            # Check if template has at least one slide
            if len(presentation.get('slides', [])) < 1:
                # Create a blank slide if template is empty
                slides_service.presentations().batchUpdate(
                    presentationId=presentation_id,
                    body={
                        'requests': [{
                            'createSlide': {
                                'insertionIndex': 0,
                                'slideLayoutReference': {
                                    'predefinedLayout': 'BLANK'
                                }
                            }
                        }]
                    }
                ).execute()
        else:
            # Create a new blank presentation
            presentation = {
                'title': f"Shipping Labels - {datetime.now().strftime('%Y-%m-%d %H:%M')}"
            }
            presentation = slides_service.presentations().create(body=presentation).execute()
            presentation_id = presentation.get('presentationId')
            
            # Create a blank slide
            slides_service.presentations().batchUpdate(
                presentationId=presentation_id,
                body={
                    'requests': [{
                        'createSlide': {
                            'slideLayoutReference': {
                                'predefinedLayout': 'BLANK'
                            }
                        }
                    }]
                }
            ).execute()
            
        presentation_url = f"https://docs.google.com/presentation/d/{presentation_id}/edit"
        
        # Get the current slides to determine deletion/updates
        presentation = slides_service.presentations().get(
            presentationId=presentation_id
        ).execute()
        
        # Clear existing slides except the first one (template slide)
        slides = presentation.get('slides', [])
        delete_requests = []
        
        if template_id and len(slides) > 1:
            # Keep the first slide if using a template, delete the rest
            for i in range(1, len(slides)):
                delete_requests.append({
                    'deleteObject': {
                        'objectId': slides[i].get('objectId')
                    }
                })
        elif len(slides) > 0 and not template_id:
            # Delete all slides if not using a template
            for slide in slides:
                delete_requests.append({
                    'deleteObject': {
                        'objectId': slide.get('objectId')
                    }
                })
            
        if delete_requests:
            slides_service.presentations().batchUpdate(
                presentationId=presentation_id,
                body={'requests': delete_requests}
            ).execute()
            
            # If we deleted all slides, create a new blank one
            if not template_id or len(slides) <= 1:
                slides_service.presentations().batchUpdate(
                    presentationId=presentation_id,
                    body={
                        'requests': [{
                            'createSlide': {
                                'slideLayoutReference': {
                                    'predefinedLayout': 'BLANK'
                                }
                            }
                        }]
                    }
                ).execute()
                
        # Get updated list of slides
        presentation = slides_service.presentations().get(
            presentationId=presentation_id
        ).execute()
        slides = presentation.get('slides', [])
        
        # Create slides from template
        template_slide_id = slides[0].get('objectId') if slides else None
        
        # Process each order - either updating template or creating new slides
        requests = []
        
        # Create a table layout on the template slide if needed
        if len(slides) > 0 and not template_id:
            # Create table structure on a blank slide
            table_width = 550  # Width in points
            table_height = 350  # Height in points
            
            # Create a 6x2 table for the shipping label
            requests.append({
                'createTable': {
                    'rows': 6,
                    'columns': 2,
                    'elementProperties': {
                        'pageObjectId': template_slide_id,
                        'size': {
                            'width': {'magnitude': table_width, 'unit': 'PT'},
                            'height': {'magnitude': table_height, 'unit': 'PT'}
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
            })
        
        # Get the updated slides to find the created table
        if len(requests) > 0:
            slides_service.presentations().batchUpdate(
                presentationId=presentation_id,
                body={'requests': requests}
            ).execute()
            
            presentation = slides_service.presentations().get(
                presentationId=presentation_id
            ).execute()
            slides = presentation.get('slides', [])
            
        # Find the table and cells on the template slide
        template_table_id = None
        cell_ids = []
        
        if slides:
            for element in slides[0].get('pageElements', []):
                if 'table' in element:
                    template_table_id = element.get('objectId')
                    
                    rows = element.get('table', {}).get('tableRows', [])
                    for row_idx, row in enumerate(rows):
                        for col_idx, cell in enumerate(row.get('tableCells', [])):
                            cell_ids.append({
                                'row': row_idx,
                                'column': col_idx,
                                'objectId': cell.get('objectId')
                            })
        
        # Create copy of template slide for each order
        for i, order in enumerate(order_details):
            # Determine if we need to duplicate the template or use it
            if i == 0 and template_slide_id:
                slide_id = template_slide_id
            else:
                # Duplicate the template slide
                duplicate_response = slides_service.presentations().batchUpdate(
                    presentationId=presentation_id,
                    body={
                        'requests': [{
                            'duplicateObject': {
                                'objectId': template_slide_id
                            }
                        }]
                    }
                ).execute()
                
                # Get the ID of the duplicated slide
                slide_id = duplicate_response.get('replies', [{}])[0].get('duplicateObject', {}).get('objectId')
            
            # Fill in the shipping label content
            # Format data for the table cells
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
            address = f"{address1}\n{address2}" if address2 else address1
            
            # Format for table cells
            cell_content = [
                ["To:", "From:"],
                [f"Name: {order.get('order_number', '')} {order.get('name', '')}", "Company: Eczema Mitten Private Limited"],
                [f"Contact: {order.get('phone', '')}", "Contact: +65 8889 5607"],
                [f"Delivery Address: {address}", "Return Address: #04-23, Block 235, Choa Chu Kang Central"],
                [f"Postal: {order.get('postal', '')}", "Postal: 680235"],
                [f"Item: {quantity} {size_display} {material} Eczema Mitten", ""]
            ]
            
            # Apply text to each cell if we're using a template with a table
            if cell_ids:
                text_updates = []
                
                for cell in cell_ids:
                    row = cell.get('row')
                    col = cell.get('column')
                    
                    if row < len(cell_content) and col < len(cell_content[row]):
                        content = cell_content[row][col]
                        
                        text_updates.append({
                            'insertText': {
                                'objectId': cell.get('objectId'),
                                'insertionIndex': 0,
                                'text': content
                            }
                        })
                
                if text_updates:
                    slides_service.presentations().batchUpdate(
                        presentationId=presentation_id,
                        body={'requests': text_updates}
                    ).execute()
        
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