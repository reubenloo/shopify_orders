import os
import json
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
        
        # Create a new presentation or use template with detailed error handling
        presentation_id = None
        presentation_url = None
        
        try:
            if template_id:
                print(f"Copying template presentation: {template_id}")
                copy_title = f"Shipping Labels - {datetime.now().strftime('%Y-%m-%d %H:%M')}"
                try:
                    drive_response = drive_service.files().copy(
                        fileId=template_id,
                        body={"name": copy_title}
                    ).execute()
                    presentation_id = drive_response.get('id')
                    print(f"Successfully copied template to new presentation: {presentation_id}")
                except Exception as e:
                    print(f"ERROR copying template: {str(e)}")
                    # Fallback to creating new presentation
                    print("Falling back to creating a new presentation...")
                    template_id = None
            
            if not template_id:
                print("Creating new blank presentation...")
                presentation = {
                    'title': f"Shipping Labels - {datetime.now().strftime('%Y-%m-%d %H:%M')}"
                }
                presentation = slides_service.presentations().create(body=presentation).execute()
                presentation_id = presentation.get('presentationId')
                print(f"Successfully created new presentation: {presentation_id}")
                
            presentation_url = f"https://docs.google.com/presentation/d/{presentation_id}/edit"
            print(f"Presentation URL: {presentation_url}")
        except Exception as e:
            print(f"ERROR creating presentation: {str(e)}")
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
            
            # If using template, we may want to manipulate existing slides
            # For now, we'll create new slides for each order
        except Exception as e:
            print(f"ERROR getting presentation details: {str(e)}")
            import traceback
            traceback.print_exc()
            # Continue with empty presentation
        
        # Create shipping label slides for each order
        try:
            print(f"Creating shipping label slides for {len(order_details)} orders...")
            
            # Create a batch update request
            requests = []
            
            # If template slide exists, clean it up
            # (This code assumes we're not using template content)
            if template_id is None:
                # Delete any existing slides
                print("Deleting any default slides...")
                existing_slides = slides_service.presentations().get(
                    presentationId=presentation_id
                ).execute().get('slides', [])
                
                for slide in existing_slides:
                    slide_id = slide.get('objectId')
                    requests.append({
                        'deleteObject': {
                            'objectId': slide_id
                        }
                    })
            
            # Process each order and create a slide
            for i, order in enumerate(order_details):
                print(f"Processing order {i+1}: {order.get('order_number')}")
                
                # Create a new slide
                requests.append({
                    'createSlide': {
                        'insertionIndex': i,
                        'slideLayoutReference': {
                            'predefinedLayout': 'BLANK'
                        }
                    }
                })
                
                # Since we need the slide ID for adding content, we'll submit the batch
                # request to create the slide first, then add content in separate requests
                if len(requests) > 0:
                    print(f"Submitting batch request to create slide {i+1}...")
                    response = slides_service.presentations().batchUpdate(
                        presentationId=presentation_id,
                        body={'requests': requests}
                    ).execute()
                    print(f"Batch request completed: {len(response.get('replies', []))} replies")
                    requests = []  # Clear requests
                    
                    # Get the newly created slide ID
                    slide_id = response.get('replies', [{}])[0].get('createSlide', {}).get('objectId')
                    if not slide_id:
                        print(f"WARNING: Could not get slide ID for order {order.get('order_number')}")
                        continue
                    
                    # Now add content to the slide
                    content_requests = []
                    
                    # Add a title box with order number
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
                    
                    # Style the title text
                    content_requests.append({
                        'updateTextStyle': {
                            'objectId': f'title_{i}',
                            'textRange': {
                                'type': 'ALL'
                            },
                            'style': {
                                'bold': True,
                                'fontSize': {
                                    'magnitude': 20,
                                    'unit': 'PT'
                                }
                            },
                            'fields': 'bold,fontSize'
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
                        address_parts.append(f"Postal: {order.get('postal')}")
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
                    
                    # Style address text
                    content_requests.append({
                        'updateTextStyle': {
                            'objectId': f'customer_{i}',
                            'textRange': {
                                'type': 'ALL'
                            },
                            'style': {
                                'fontSize': {
                                    'magnitude': 12,
                                    'unit': 'PT'
                                }
                            },
                            'fields': 'fontSize'
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
                    
                    # Style product text
                    content_requests.append({
                        'updateTextStyle': {
                            'objectId': f'product_{i}',
                            'textRange': {
                                'type': 'ALL'
                            },
                            'style': {
                                'fontSize': {
                                    'magnitude': 12,
                                    'unit': 'PT'
                                }
                            },
                            'fields': 'fontSize'
                        }
                    })
                    
                    # Submit the content requests
                    if len(content_requests) > 0:
                        print(f"Submitting content requests for slide {i+1}...")
                        content_response = slides_service.presentations().batchUpdate(
                            presentationId=presentation_id,
                            body={'requests': content_requests}
                        ).execute()
                        print(f"Content batch completed: {len(content_response.get('replies', []))} replies")
            
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