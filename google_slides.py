import os
from googleapiclient.discovery import build
from google.oauth2 import service_account
from datetime import datetime, timedelta
import re
import pytz

def create_shipping_slides(order_details, credentials_path, template_id=None):
    """
    Create or update a Google Slides presentation with shipping labels for orders
    
    Args:
        order_details: List of dictionaries containing order information
        credentials_path: Path to the service account JSON credentials file
        template_id: Optional ID of a template presentation to copy
        
    Returns:
        tuple: (presentation_url, pdf_path) URLs of the created presentation and path to PDF
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
        else:
            # Create a new blank presentation
            presentation = {
                'title': f"Shipping Labels - {datetime.now().strftime('%Y-%m-%d %H:%M')}"
            }
            presentation = slides_service.presentations().create(body=presentation).execute()
            presentation_id = presentation.get('presentationId')
            
        presentation_url = f"https://docs.google.com/presentation/d/{presentation_id}/edit"
        
        # Get the current slides to determine deletion/updates
        presentation = slides_service.presentations().get(
            presentationId=presentation_id
        ).execute()
        
        # Delete all existing slides to start fresh
        slides = presentation.get('slides', [])
        delete_requests = []
        
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
        
        # Group orders by date (assuming they're already sorted)
        # Get current date in GMT+8
        sg_timezone = pytz.timezone('Asia/Singapore')
        current_date = datetime.now(sg_timezone).strftime('%B %d, %Y')
        
        # Create date slide first
        date_slide_requests = [{
            'createSlide': {
                'slideLayoutReference': {
                    'predefinedLayout': 'BLANK'
                },
                'placeholderIdMappings': []
            }
        }]
        
        date_slide_response = slides_service.presentations().batchUpdate(
            presentationId=presentation_id,
            body={'requests': date_slide_requests}
        ).execute()
        
        date_slide_id = date_slide_response.get('replies', [{}])[0].get('createSlide', {}).get('objectId')
        
        # Add date text to the slide
        date_text_requests = [{
            'createShape': {
                'objectId': 'dateText',
                'shapeType': 'TEXT_BOX',
                'elementProperties': {
                    'pageObjectId': date_slide_id,
                    'size': {
                        'width': {'magnitude': 500, 'unit': 'PT'},
                        'height': {'magnitude': 100, 'unit': 'PT'}
                    },
                    'transform': {
                        'scaleX': 1,
                        'scaleY': 1,
                        'translateX': 100,
                        'translateY': 150,
                        'unit': 'PT'
                    }
                }
            }
        },
        {
            'insertText': {
                'objectId': 'dateText',
                'insertionIndex': 0,
                'text': current_date
            }
        },
        {
            'updateTextStyle': {
                'objectId': 'dateText',
                'textRange': {
                    'type': 'ALL'
                },
                'style': {
                    'fontSize': {
                        'magnitude': 48,
                        'unit': 'PT'
                    },
                    'fontWeight': 400,
                    'textAlign': 'CENTER'
                },
                'fields': 'fontSize,fontWeight,textAlign'
            }
        }]
        
        slides_service.presentations().batchUpdate(
            presentationId=presentation_id,
            body={'requests': date_text_requests}
        ).execute()
        
        # Create shipping label slides
        for i, order in enumerate(order_details):
            # Create a new slide for each shipping label
            shipping_slide_request = [{
                'createSlide': {
                    'slideLayoutReference': {
                        'predefinedLayout': 'BLANK'
                    },
                    'placeholderIdMappings': []
                }
            }]
            
            shipping_slide_response = slides_service.presentations().batchUpdate(
                presentationId=presentation_id,
                body={'requests': shipping_slide_request}
            ).execute()
            
            shipping_slide_id = shipping_slide_response.get('replies', [{}])[0].get('createSlide', {}).get('objectId')
            
            # Create the table for shipping label
            table_request = [{
                'createTable': {
                    'objectId': f'table_{i}',
                    'rows': 6,
                    'columns': 2,
                    'elementProperties': {
                        'pageObjectId': shipping_slide_id,
                        'size': {
                            'width': {'magnitude': 650, 'unit': 'PT'},
                            'height': {'magnitude': 400, 'unit': 'PT'}
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
            }]
            
            # Create table first
            table_response = slides_service.presentations().batchUpdate(
                presentationId=presentation_id,
                body={'requests': table_request}
            ).execute()
            
            # Get updated slide to find table cells
            slide = slides_service.presentations().get(
                presentationId=presentation_id,
                fields='slides'
            ).execute()
            
            current_slide = None
            for s in slide.get('slides', []):
                if s.get('objectId') == shipping_slide_id:
                    current_slide = s
                    break
            
            if not current_slide:
                continue
                
            # Find table in the slide
            table_element = None
            for element in current_slide.get('pageElements', []):
                if 'table' in element:
                    table_element = element
                    break
            
            if not table_element:
                continue
                
            table_id = table_element.get('objectId')
                
            # Format data for the table cells
            quantity = "2" if order.get('is_bundle', False) else "1"
            size = order.get('size', '')
            material = order.get('material', '')
            
            # Format the size for kid sizes
            if '(' in size and 'cm' in size:
                size_display = size.split('(')[1].replace(')', '').split('-')[0] + 'cm'
            else:
                size_display = size
                
            # Extract address info
            name = order.get('name', '')
            order_number = order.get('order_number', '')
            address1 = order.get('address1', '')
            address2 = order.get('address2', '')
            full_address = f"{address1}"
            if address2 and address2.strip():
                full_address += f"\n{address2}"
            postal = order.get('postal', '')
            phone = order.get('phone', '')
            
            # Create text for cells
            cell_content = [
                ["To:", "From:"],
                [f"Name: #{order_number} {name}", "Company: Eczema Mitten Private Limited"],
                [f"Contact: {phone}", "Contact: +65 8889 5607"],
                [f"Delivery Address:\n{full_address}", "Return Address: #04-23, Block 235, Choa Chu Kang Central"],
                [f"Postal: {postal}", "Postal: 680235"],
                [f"Item: {quantity} {size_display} {material} Eczema Mitten", ""]
            ]
            
            # Style definitions
            header_style = {
                'fontSize': {'magnitude': 24, 'unit': 'PT'},
                'bold': True,
                'underline': True
            }
            
            normal_style = {
                'fontSize': {'magnitude': 14, 'unit': 'PT'}
            }
            
            company_style = {
                'fontSize': {'magnitude': 14, 'unit': 'PT'},
                'bold': True
            }
            
            strikethrough_style = {
                'fontSize': {'magnitude': 14, 'unit': 'PT'},
                'strikethrough': True
            }
            
            # Create a comprehensive request for cell formatting and content
            cell_requests = []
            
            # First get the table cells
            table_cells = table_element.get('table', {}).get('tableRows', [])
            
            for row_idx, row in enumerate(table_cells):
                for col_idx, cell in enumerate(row.get('tableCells', [])):
                    cell_id = cell.get('objectId')
                    
                    if row_idx < len(cell_content) and col_idx < len(cell_content[row_idx]):
                        content = cell_content[row_idx][col_idx]
                        
                        # Text insertion request
                        cell_requests.append({
                            'insertText': {
                                'objectId': cell_id,
                                'insertionIndex': 0,
                                'text': content
                            }
                        })
                        
                        # Apply header style to first row
                        if row_idx == 0:
                            cell_requests.append({
                                'updateTextStyle': {
                                    'objectId': cell_id,
                                    'textRange': {'type': 'ALL'},
                                    'style': header_style,
                                    'fields': 'fontSize,bold,underline'
                                }
                            })
                        # Special formatting for company name in "From" column
                        elif row_idx == 1 and col_idx == 1:
                            cell_requests.append({
                                'updateTextStyle': {
                                    'objectId': cell_id,
                                    'textRange': {'type': 'ALL'},
                                    'style': company_style,
                                    'fields': 'fontSize,bold'
                                }
                            })
                        # Special formatting for strikethrough text
                        elif (row_idx == 3 and col_idx == 1) or (row_idx == 4 and col_idx == 1):
                            cell_requests.append({
                                'updateTextStyle': {
                                    'objectId': cell_id,
                                    'textRange': {'type': 'ALL'},
                                    'style': strikethrough_style,
                                    'fields': 'fontSize,strikethrough'
                                }
                            })
                        # Normal formatting for all other cells
                        else:
                            cell_requests.append({
                                'updateTextStyle': {
                                    'objectId': cell_id,
                                    'textRange': {'type': 'ALL'},
                                    'style': normal_style,
                                    'fields': 'fontSize'
                                }
                            })
                            
                        # Add bold formatting for header parts (like "Name:", "Contact:", etc.)
                        if row_idx > 0 and ":" in content:
                            header_end = content.index(":") + 1
                            cell_requests.append({
                                'updateTextStyle': {
                                    'objectId': cell_id,
                                    'textRange': {
                                        'startIndex': 0,
                                        'endIndex': header_end
                                    },
                                    'style': {'bold': True},
                                    'fields': 'bold'
                                }
                            })
            
            # Apply all cell formatting
            if cell_requests:
                slides_service.presentations().batchUpdate(
                    presentationId=presentation_id,
                    body={'requests': cell_requests}
                ).execute()
        
        return presentation_url, None
        
    except Exception as e:
        print(f"Error creating Google Slides: {str(e)}")
        return None, None

def get_template_id_from_url(url):
    """Extract the presentation ID from a Google Slides URL"""
    match = re.search(r'/d/([a-zA-Z0-9-_]+)', url)
    if match:
        return match.group(1)
    return None