import os
import json
from googleapiclient.discovery import build
from google.oauth2 import service_account
from datetime import datetime
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
                
                # Now update the order details on this slide
                update_order_details(slides_service, presentation_id, new_slide_id, order)
                
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

def update_order_details(slides_service, presentation_id, slide_id, order):
    """
    Update a slide with order information
    
    Args:
        slides_service: Google Slides API service
        presentation_id: ID of the presentation
        slide_id: ID of the slide to update
        order: Dictionary containing order information
    """
    try:
        print(f"Updating slide {slide_id} with order details for order: {order.get('order_number', 'unknown')}")
        
        # Get the slide details
        slide = slides_service.presentations().pages().get(
            presentationId=presentation_id,
            pageObjectId=slide_id
        ).execute()
        
        # Find all text elements on the slide (shapes with text)
        text_elements = []
        for element in slide.get('pageElements', []):
            if 'shape' in element and 'text' in element.get('shape', {}):
                # Get the element ID
                element_id = element.get('objectId')
                
                # Extract text content
                text_content = ""
                for text_element in element.get('shape', {}).get('text', {}).get('textElements', []):
                    if 'textRun' in text_element:
                        text_content += text_element.get('textRun', {}).get('content', '')
                
                # Store element ID and text content
                text_elements.append({
                    'id': element_id,
                    'text': text_content.strip()
                })
                
                print(f"Found text element: ID={element_id}, Content=\"{text_content.strip()}\"")
        
        # Find table element if exists
        table_id = None
        table_cells = []
        
        for element in slide.get('pageElements', []):
            if 'table' in element:
                table_id = element.get('objectId')
                table = element.get('table')
                
                # Process each cell in the table
                for row_idx, row in enumerate(table.get('tableRows', [])):
                    for col_idx, cell in enumerate(row.get('tableCells', [])):
                        # Get the cell's object ID
                        cell_id = cell.get('objectId')
                        
                        # Extract cell content
                        cell_content = ""
                        if 'text' in cell:
                            for text_element in cell.get('text', {}).get('textElements', []):
                                if 'textRun' in text_element:
                                    cell_content += text_element.get('textRun', {}).get('content', '')
                        
                        table_cells.append({
                            'id': cell_id,
                            'row': row_idx,
                            'col': col_idx,
                            'text': cell_content.strip()
                        })
                        
                        print(f"Found table cell: Row={row_idx}, Col={col_idx}, ID={cell_id}, Content=\"{cell_content.strip()}\"")
                break  # We only need the first table
        
        # Prepare the order information
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
        
        # Prepare update requests
        update_requests = []
        
        # APPROACH 1: Try updating cells by row/column position if table exists
        if table_id and table_cells:
            print("Using table-based approach for updates")
            
            # Create a mapping of (row, col) to cell ID
            cell_mapping = {(cell['row'], cell['col']): cell['id'] for cell in table_cells}
            
            # Update cells based on position (typical layout)
            field_updates = [
                # (row, col, content)
                (0, 0, f"Name: #{order.get('order_number', '').replace('#', '')} {order.get('name', '')}"),
                (1, 0, f"Contact: {order.get('phone', '')}"),
                (2, 0, f"Delivery Address: {address}"),
                (3, 0, f"Postal: {order.get('postal', '')}"),
                (4, 0, f"Item: {quantity} {size_display} {material} Eczema Mitten")
            ]
            
            for row, col, content in field_updates:
                if (row, col) in cell_mapping:
                    cell_id = cell_mapping[(row, col)]
                    print(f"Updating table cell ({row}, {col}) with content: \"{content}\"")
                    
                    update_requests.extend([
                        {
                            'deleteText': {
                                'objectId': cell_id,
                                'textRange': {
                                    'type': 'ALL'
                                }
                            }
                        },
                        {
                            'insertText': {
                                'objectId': cell_id,
                                'insertionIndex': 0,
                                'text': content
                            }
                        }
                    ])
        
        # APPROACH 2: Try updating text elements based on content matching
        else:
            print("Using text-based approach for updates")
            
            # Define update pattern for text fields
            field_patterns = [
                ("NAME:", f"Name: #{order.get('order_number', '').replace('#', '')} {order.get('name', '')}"),
                ("CONTACT:", f"Contact: {order.get('phone', '')}"),
                ("DELIVERY ADDRESS:", f"Delivery Address: {address}"),
                ("POSTAL:", f"Postal: {order.get('postal', '')}"),
                ("ITEM:", f"Item: {quantity} {size_display} {material} Eczema Mitten")
            ]
            
            # Try to match fields by content
            for element in text_elements:
                element_id = element['id']
                element_text = element['text'].upper()
                
                for pattern, replacement in field_patterns:
                    if pattern in element_text:
                        print(f"Matched field pattern '{pattern}' in element: \"{element_text}\"")
                        print(f"Updating with: \"{replacement}\"")
                        
                        update_requests.extend([
                            {
                                'deleteText': {
                                    'objectId': element_id,
                                    'textRange': {
                                        'type': 'ALL'
                                    }
                                }
                            },
                            {
                                'insertText': {
                                    'objectId': element_id,
                                    'insertionIndex': 0,
                                    'text': replacement
                                }
                            }
                        ])
                        break  # Only apply one replacement per element
        
        # APPROACH 3: Last resort - try each text element and update it with relevant content
        if not update_requests and text_elements:
            print("Using position-based approach for text elements")
            
            # Try using position-based updates if no matches were found
            field_values = [
                f"Name: #{order.get('order_number', '').replace('#', '')} {order.get('name', '')}",
                f"Contact: {order.get('phone', '')}",
                f"Delivery Address: {address}",
                f"Postal: {order.get('postal', '')}",
                f"Item: {quantity} {size_display} {material} Eczema Mitten"
            ]
            
            # Get text elements sorted by vertical position (top to bottom)
            element_positions = []
            for element in slide.get('pageElements', []):
                if 'shape' in element and 'text' in element.get('shape', {}):
                    element_positions.append({
                        'id': element.get('objectId'),
                        'y': element.get('transform', {}).get('translateY', 0)
                    })
            
            sorted_elements = sorted(element_positions, key=lambda x: x['y'])
            
            # Update elements in order
            for i, element in enumerate(sorted_elements):
                if i < len(field_values):
                    print(f"Positional update: element at position {i} with content: \"{field_values[i]}\"")
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
                                'text': field_values[i]
                            }
                        }
                    ])
        
        # Execute all updates
        if update_requests:
            print(f"Executing {len(update_requests)} update requests")
            slides_service.presentations().batchUpdate(
                presentationId=presentation_id,
                body={'requests': update_requests}
            ).execute()
            print("Successfully executed updates for slide")
        else:
            print("WARNING: No updates were prepared for this slide")
        
    except Exception as e:
        print(f"ERROR updating order details: {str(e)}")
        import traceback
        traceback.print_exc()

def find_table_cells(slides_service, presentation_id, slide_id):
    """
    Find all cells in a table and map them to their row/column positions
    
    Args:
        slides_service: Google Slides API service
        presentation_id: ID of the presentation
        slide_id: ID of the slide containing the table
        
    Returns:
        dict: Mapping of (row, column) tuples to cell IDs
    """
    try:
        # Get the slide with detailed structure
        slide = slides_service.presentations().pages().get(
            presentationId=presentation_id,
            pageObjectId=slide_id
        ).execute()
        
        # Find the table and its cells
        cell_mapping = {}
        
        for element in slide.get('pageElements', []):
            if 'table' in element:
                table = element.get('table')
                
                for row_idx, row in enumerate(table.get('tableRows', [])):
                    for col_idx, cell in enumerate(row.get('tableCells', [])):
                        cell_id = None
                        
                        # Find the text content object ID for this cell
                        if 'text' in cell:
                            for text_element in cell.get('text', {}).get('textElements', []):
                                if 'paragraphMarker' in text_element:
                                    cell_id = cell.get('objectId')
                                    break
                        
                        if cell_id:
                            cell_mapping[(row_idx, col_idx)] = cell_id
        
        return cell_mapping
        
    except Exception as e:
        print(f"ERROR finding table cells: {str(e)}")
        import traceback
        traceback.print_exc()
        return {}

def update_text_fields(slides_service, presentation_id, text_fields, order):
    """
    Update text fields in a slide when no table is found
    
    Args:
        slides_service: Google Slides API service
        presentation_id: ID of the presentation
        text_fields: List of text field objects with ID and content
        order: Dictionary containing order information
    """
    try:
        update_requests = []
        
        # Prepare the order information
        quantity = "2" if order.get('is_bundle', False) else "1"
        size = order.get('size', '')
        material = order.get('material', '')
        
        if '(' in size and 'cm' in size:
            size_display = size.split('(')[1].replace(')', '').split('-')[0] + 'cm'
        else:
            size_display = size
            
        address1 = order.get('address1', '')
        address2 = order.get('address2', '')
        address = f"{address1}\n{address2}" if address2 and address2.strip() else address1
        
        # Match text fields to their purpose based on content
        for field in text_fields:
            text = field.get('text', '').upper()
            field_id = field.get('id')
            
            if "NAME:" in text:
                update_requests.extend([
                    {
                        'deleteText': {
                            'objectId': field_id,
                            'textRange': {
                                'type': 'ALL'
                            }
                        }
                    },
                    {
                        'insertText': {
                            'objectId': field_id,
                            'insertionIndex': 0,
                            'text': f"Name: #{order.get('order_number', '').replace('#', '')} {order.get('name', '')}"
                        }
                    }
                ])
            elif "CONTACT:" in text and "ECZEMA" not in text:
                update_requests.extend([
                    {
                        'deleteText': {
                            'objectId': field_id,
                            'textRange': {
                                'type': 'ALL'
                            }
                        }
                    },
                    {
                        'insertText': {
                            'objectId': field_id,
                            'insertionIndex': 0,
                            'text': f"Contact: {order.get('phone', '')}"
                        }
                    }
                ])
            elif "DELIVERY ADDRESS:" in text:
                update_requests.extend([
                    {
                        'deleteText': {
                            'objectId': field_id,
                            'textRange': {
                                'type': 'ALL'
                            }
                        }
                    },
                    {
                        'insertText': {
                            'objectId': field_id,
                            'insertionIndex': 0,
                            'text': f"Delivery Address: {address}"
                        }
                    }
                ])
            elif "POSTAL:" in text and "680" not in text:
                update_requests.extend([
                    {
                        'deleteText': {
                            'objectId': field_id,
                            'textRange': {
                                'type': 'ALL'
                            }
                        }
                    },
                    {
                        'insertText': {
                            'objectId': field_id,
                            'insertionIndex': 0,
                            'text': f"Postal: {order.get('postal', '')}"
                        }
                    }
                ])
            elif "ITEM:" in text:
                update_requests.extend([
                    {
                        'deleteText': {
                            'objectId': field_id,
                            'textRange': {
                                'type': 'ALL'
                            }
                        }
                    },
                    {
                        'insertText': {
                            'objectId': field_id,
                            'insertionIndex': 0,
                            'text': f"Item: {quantity} {size_display} {material} Eczema Mitten"
                        }
                    }
                ])
        
        # Execute all updates
        if update_requests:
            slides_service.presentations().batchUpdate(
                presentationId=presentation_id,
                body={'requests': update_requests}
            ).execute()
            print("Successfully updated order details in text fields")
        else:
            print("WARNING: No updates were made to the text fields")
            
    except Exception as e:
        print(f"ERROR updating text fields: {str(e)}")
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