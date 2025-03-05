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