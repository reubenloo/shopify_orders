import os
import json
from googleapiclient.discovery import build
from google.oauth2 import service_account
from datetime import datetime, timedelta
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
                
                # Update the presentation title to include the current date/time
                try:
                    current_time = datetime.now().strftime('%Y-%m-%d %H:%M')
                    requests = [{
                        'updatePresentationProperties': {
                            'fields': 'title',
                            'presentationProperties': {
                                'title': f"Shipping Labels - {current_time}"
                            }
                        }
                    }]
                    slides_service.presentations().batchUpdate(
                        presentationId=presentation_id,
                        body={'requests': requests}
                    ).execute()
                    print(f"Updated presentation title to include timestamp: {current_time}")
                except Exception as e:
                    print(f"Warning: Could not update presentation title: {str(e)}")
                    # Continue anyway
            else:
                print("No template ID provided, creating new presentation...")
                presentation = {
                    'title': f"Shipping Labels - {datetime.now().strftime('%Y-%m-%d %H:%M')}"
                }
                presentation = slides_service.presentations().create(body=presentation).execute()
                presentation_id = presentation.get('presentationId')
                print(f"Successfully created new presentation: {presentation_id}")
                
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
            
            # Save the ID of the first slide if it exists (template slide)
            template_slide_id = None
            if slides:
                template_slide_id = slides[0].get('objectId')
                print(f"Found template slide with ID: {template_slide_id}")
        except Exception as e:
            print(f"ERROR getting presentation details: {str(e)}")
            import traceback
            traceback.print_exc()
            # Continue with empty presentation
        
        # Remove all existing slides except the first one (template)
        try:
            print("Cleaning up existing slides...")
            requests = []
            
            if slides:
                # Start from index 1 to keep the first slide as template
                for slide in slides[1:]:
                    slide_id = slide.get('objectId')
                    print(f"Deleting slide: {slide_id}")
                    requests.append({
                        'deleteObject': {
                            'objectId': slide_id
                        }
                    })
                
                if requests:
                    # Submit deletion requests
                    print(f"Submitting batch request to delete {len(requests)} slides...")
                    response = slides_service.presentations().batchUpdate(
                        presentationId=presentation_id,
                        body={'requests': requests}
                    ).execute()
                    print(f"Deletion completed: {len(response.get('replies', []))} replies")
        except Exception as e:
            print(f"WARNING: Could not clean up existing slides: {str(e)}")
            # Continue anyway
        
        # Create shipping label slides for each order
        try:
            print(f"Creating shipping label slides for {len(order_details)} orders...")
            
            # Process each order and create a slide
            for i, order in enumerate(order_details):
                print(f"Processing order {i+1}: {order.get('order_number')}")
                
                # Create a new slide, duplicating the template if available
                requests = []
                if template_slide_id:
                    # Duplicate the template slide
                    requests.append({
                        'duplicateObject': {
                            'objectId': template_slide_id,
                        }
                    })
                else:
                    # Create a blank slide if no template
                    requests.append({
                        'createSlide': {
                            'slideLayoutReference': {
                                'predefinedLayout': 'BLANK'
                            }
                        }
                    })
                
                # Submit the request to create the slide
                print(f"Submitting request to create slide for order {i+1}...")
                response = slides_service.presentations().batchUpdate(
                    presentationId=presentation_id,
                    body={'requests': requests}
                ).execute()
                
                # Get the newly created slide ID
                slide_id = None
                if 'duplicateObject' in response.get('replies', [{}])[0]:
                    slide_id = response.get('replies', [{}])[0].get('duplicateObject', {}).get('objectId')
                else:
                    slide_id = response.get('replies', [{}])[0].get('createSlide', {}).get('objectId')
                
                if not slide_id:
                    print(f"WARNING: Could not get slide ID for order {order.get('order_number')}")
                    continue
                
                print(f"Created slide with ID: {slide_id}")
                
                # Now find and update text placeholders on the slide
                # First get the current slide details to find text elements
                slide_details = slides_service.presentations().get(
                    presentationId=presentation_id,
                    fields="slides"
                ).execute()
                
                # Find the right slide
                current_slide = None
                for slide in slide_details.get('slides', []):
                    if slide.get('objectId') == slide_id:
                        current_slide = slide
                        break
                
                if not current_slide:
                    print(f"WARNING: Cannot find slide {slide_id} in presentation")
                    continue
                
                # Extract text elements from the slide
                text_elements = []
                for element in current_slide.get('pageElements', []):
                    if 'shape' in element and 'text' in element.get('shape', {}):
                        element_id = element.get('objectId')
                        text_content = element.get('shape', {}).get('text', {}).get('textElements', [])
                        
                        # Look for placeholder text to identify what to replace
                        placeholder_text = ''
                        for text_element in text_content:
                            if 'textRun' in text_element:
                                text = text_element.get('textRun', {}).get('content', '')
                                placeholder_text += text
                        
                        text_elements.append({
                            'id': element_id,
                            'text': placeholder_text.strip()
                        })
                
                print(f"Found {len(text_elements)} text elements on slide")
                
                # Prepare replacements based on placeholders
                content_requests = []
                
                for elem in text_elements:
                    elem_id = elem.get('id')
                    placeholder = elem.get('text', '').upper()
                    
                    # Skip elements with no text
                    if not placeholder:
                        continue
                    
                    print(f"Processing text element: '{placeholder}'")
                    
                    # Determine what to replace based on placeholder text
                    if "ORDER" in placeholder:
                        # This is the order number field
                        content_requests.append({
                            'deleteText': {
                                'objectId': elem_id,
                                'textRange': {
                                    'type': 'ALL'
                                }
                            }
                        })
                        content_requests.append({
                            'insertText': {
                                'objectId': elem_id,
                                'text': f"Order #{order.get('order_number')}"
                            }
                        })
                    elif "NAME" in placeholder or "CUSTOMER" in placeholder:
                        # Customer name field
                        content_requests.append({
                            'deleteText': {
                                'objectId': elem_id,
                                'textRange': {
                                    'type': 'ALL'
                                }
                            }
                        })
                        content_requests.append({
                            'insertText': {
                                'objectId': elem_id,
                                'text': order.get('name', '')
                            }
                        })
                    elif "ADDRESS" in placeholder:
                        # Address field
                        address_parts = [
                            order.get('address1', ''),
                            order.get('address2', '') if order.get('address2') else '',
                            f"Singapore {order.get('postal', '')}"
                        ]
                        # Filter out empty parts
                        address_parts = [part for part in address_parts if part]
                        
                        content_requests.append({
                            'deleteText': {
                                'objectId': elem_id,
                                'textRange': {
                                    'type': 'ALL'
                                }
                            }
                        })
                        content_requests.append({
                            'insertText': {
                                'objectId': elem_id,
                                'text': "\n".join(address_parts)
                            }
                        })
                    elif "PHONE" in placeholder:
                        # Phone number field
                        content_requests.append({
                            'deleteText': {
                                'objectId': elem_id,
                                'textRange': {
                                    'type': 'ALL'
                                }
                            }
                        })
                        content_requests.append({
                            'insertText': {
                                'objectId': elem_id,
                                'text': order.get('phone', 'N/A')
                            }
                        })
                    elif "PRODUCT" in placeholder or "ITEM" in placeholder:
                        # Product details field
                        quantity = "2 Pairs" if order.get('is_bundle') else "1 Pair"
                        product_info = f"Eczema Mittens - {quantity}\nSize: {order.get('size', 'Unknown')}\nMaterial: {order.get('material', 'Unknown')}"
                        
                        content_requests.append({
                            'deleteText': {
                                'objectId': elem_id,
                                'textRange': {
                                    'type': 'ALL'
                                }
                            }
                        })
                        content_requests.append({
                            'insertText': {
                                'objectId': elem_id,
                                'text': product_info
                            }
                        })
                
                # If no placeholders were found, add generic content boxes
                if not text_elements:
                    print("No text placeholders found. Adding generic content...")
                    
                    # Add order number text box
                    content_requests.append({
                        'createShape': {
                            'objectId': f'title_{i}',
                            'shapeType': 'TEXT_BOX',
                            'elementProperties': {
                                'pageObjectId': slide_id,
                                'size': {
                                    'width': {'magnitude': 500, 'unit': 'PT'},
                                    'height': {'magnitude': 50, 'unit': 'PT'}
                                },
                                'transform': {
                                    'scaleX': 1,
                                    'scaleY': 1,
                                    'translateX': 50,
                                    'translateY': 30,
                                    'unit': 'PT'
                                }
                            }
                        }
                    })
                    
                    # Add text to the title box
                    content_requests.append({
                        'insertText': {
                            'objectId': f'title_{i}',
                            'text': f"Order #{order.get('order_number')}"
                        }
                    })
                    
                    # Add customer info box
                    content_requests.append({
                        'createShape': {
                            'objectId': f'customer_{i}',
                            'shapeType': 'TEXT_BOX',
                            'elementProperties': {
                                'pageObjectId': slide_id,
                                'size': {
                                    'width': {'magnitude': 300, 'unit': 'PT'},
                                    'height': {'magnitude': 150, 'unit': 'PT'}
                                },
                                'transform': {
                                    'scaleX': 1,
                                    'scaleY': 1,
                                    'translateX': 50,
                                    'translateY': 100,
                                    'unit': 'PT'
                                }
                            }
                        }
                    })
                    
                    # Compile address
                    address_parts = [
                        order.get('name', ''),
                        order.get('address1', ''),
                        order.get('address2', '')
                    ]
                    # Filter out empty parts
                    address_parts = [part for part in address_parts if part]
                    # Add postal code if available
                    if order.get('postal'):
                        address_parts.append(f"Singapore {order.get('postal')}")
                    # Add phone if available
                    if order.get('phone'):
                        address_parts.append(f"Phone: {order.get('phone')}")
                    
                    # Insert address text
                    content_requests.append({
                        'insertText': {
                            'objectId': f'customer_{i}',
                            'text': "Ship To:\n" + "\n".join(address_parts)
                        }
                    })
                    
                    # Add product info box
                    content_requests.append({
                        'createShape': {
                            'objectId': f'product_{i}',
                            'shapeType': 'TEXT_BOX',
                            'elementProperties': {
                                'pageObjectId': slide_id,
                                'size': {
                                    'width': {'magnitude': 300, 'unit': 'PT'},
                                    'height': {'magnitude': 150, 'unit': 'PT'}
                                },
                                'transform': {
                                    'scaleX': 1,
                                    'scaleY': 1,
                                    'translateX': 400,
                                    'translateY': 100,
                                    'unit': 'PT'
                                }
                            }
                        }
                    })
                    
                    # Product details
                    quantity = "2 Pairs" if order.get('is_bundle') else "1 Pair"
                    product_details = [
                        "Product: Eczema Mittens",
                        f"Quantity: {quantity}",
                        f"Size: {order.get('size', 'Unknown')}",
                        f"Material: {order.get('material', 'Unknown')}"
                    ]
                    
                    # Insert product text
                    content_requests.append({
                        'insertText': {
                            'objectId': f'product_{i}',
                            'text': "\n".join(product_details)
                        }
                    })
                
                # Submit the content requests
                if content_requests:
                    print(f"Submitting {len(content_requests)} content updates for slide {i+1}...")
                    try:
                        content_response = slides_service.presentations().batchUpdate(
                            presentationId=presentation_id,
                            body={'requests': content_requests}
                        ).execute()
                        print(f"Content updates completed: {len(content_response.get('replies', []))} replies")
                    except Exception as e:
                        print(f"ERROR updating slide content: {str(e)}")
                        import traceback
                        traceback.print_exc()
            
            print("All slides created successfully!")
            
            # Attempt to create a PDF export
            # Note: Direct PDF export is limited with the Drive API, so we'll just
            # return None for pdf_path, and users will need to download PDF from Google Slides UI
            pdf_path = None
            
            return presentation_url, pdf_path
            
        except Exception as e:
            print(f"ERROR creating slides: {str(e)}")
            import traceback
            traceback.print_exc()
            # Return the URL if we have it, even if there was an error adding content
            return presentation_url, None
        
    except Exception as e:
        print(f"ERROR in create_shipping_slides: {str(e)}")
        import traceback
        traceback.print_exc()
        return None, None

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