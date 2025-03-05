import pandas as pd
import os
from collections import Counter, defaultdict
from google_slides import create_shipping_slides, get_template_id_from_url

def safe_str_slice(value, length):
    """Safely convert value to string and slice it, handling NaN values"""
    if pd.isna(value):
        return ''
    return str(value)[:length]

def clean_shopify_data(df):
    """Clean Shopify export data by removing incomplete/duplicate orders and handling updated orders"""
    # Remove rows where Financial Status is empty (unpaid/incomplete orders)
    df = df[df['Financial Status'].notna()].copy()
    
    # Extract order number from the Name column (removing the # symbol)
    df.loc[:, 'order_number'] = df['Name'].str.replace('#', '').astype(str)
    
    # For each order number with multiple rows (modified orders)
    duplicate_orders = df[df.duplicated(subset=['order_number'], keep=False)]
    unique_order_numbers = duplicate_orders['order_number'].unique()
    
    print(f"Found {len(unique_order_numbers)} orders with modifications")
    
    for order_num in unique_order_numbers:
        order_rows = df[df['order_number'] == order_num].copy()
        if len(order_rows) >= 2:
            # Always keep the last row (most recent modification)
            latest_row_index = order_rows.index[-1]
            # Keep only the latest row for this order
            df = df.drop(order_rows.index[:-1])
            print(f"Order #{order_num}: Keeping the most recent modification")
    
    # Drop the temporary order_number column
    df = df.drop('order_number', axis=1)
    
    return df

def filter_international_orders(df):
    """Filter out orders from Singapore, US, and Canada"""
    # Create copies to track different types of orders
    sg_orders = df[df['Shipping Country'] == 'SG'].copy()
    us_ca_orders = df[(df['Shipping Country'] == 'US') | (df['Shipping Country'] == 'CA')].copy()
    intl_orders = df[(df['Shipping Country'] != 'SG') & 
                     (df['Shipping Country'] != 'US') & 
                     (df['Shipping Country'] != 'CA')].copy()
    
    print(f"Total orders: {len(df)}")
    print(f"Singapore orders (excluded from SingPost): {len(sg_orders)}")
    print(f"US/Canada orders (excluded from SingPost): {len(us_ca_orders)}")
    print(f"International orders for SingPost: {len(intl_orders)}")
    
    return intl_orders, sg_orders, us_ca_orders

def parse_product_details(lineitem_name):
    """Parse product details from lineitem name"""
    lineitem_name = lineitem_name.upper()
    is_bundle = 'BUNDLE' in lineitem_name or '2 PAIRS' in lineitem_name or 'BUNDLE OF 2' in lineitem_name
    
    # Determine material
    if 'COTTON' in lineitem_name:
        material = 'Cotton'
    elif 'TENCEL' in lineitem_name:
        material = 'Tencel'
    else:
        material = 'Unknown'
        print(f"WARNING: Unknown material in product: {lineitem_name}")
            
    # Determine size
    if 'KID (100-110CM)' in lineitem_name or '(100-110CM)' in lineitem_name:
        size = '(100-110cm)'
    elif 'KID (110-120CM)' in lineitem_name or '(110-120CM)' in lineitem_name:
        size = '(110-120cm)'
    elif 'KID (120-130CM)' in lineitem_name or '(120-130CM)' in lineitem_name:
        size = '(120-130cm)'
    elif 'KID (130-140CM)' in lineitem_name or '(130-140CM)' in lineitem_name:
        size = '(130-140cm)'
    elif 'XS (140-150CM)' in lineitem_name or '(140-150CM)' in lineitem_name:
        size = 'XS (140-150cm)'
    elif 'S (150-160CM)' in lineitem_name or '(150-160CM)' in lineitem_name:
        size = 'S (150-160cm)'
    elif 'M (160-170CM)' in lineitem_name or '(160-170CM)' in lineitem_name:
        size = 'M (160-170cm)'
    elif 'L (170-180CM)' in lineitem_name or '(170-180CM)' in lineitem_name:
        size = 'L (170-180cm)'
    elif 'XL (180-190CM)' in lineitem_name or '(180-190CM)' in lineitem_name:
        size = 'XL (180-190cm)'
    else:
        size = 'Unknown'
        print(f"WARNING: Unknown size in product: {lineitem_name}")

    return is_bundle, material, size

def convert_shopify_to_singpost(shopify_file, output_file):
    # Check if input file exists
    if not os.path.exists(shopify_file):
        return f"Error: Input file '{shopify_file}' not found. Please check the file path.", None, None
    
    # Read and clean Shopify orders
    df = pd.read_csv(shopify_file)
    df = clean_shopify_data(df)
    
    # Filter international orders (exclude SG, US, CA)
    intl_orders, sg_orders, us_ca_orders = filter_international_orders(df)
    
    # Create output dataframe with SingPost required columns
    singpost_data = []
    
    # Counters for product breakdown
    sg_product_counter = defaultdict(int)
    us_ca_product_counter = defaultdict(int)
    intl_product_counter = defaultdict(int)
    
    sg_order_details = []
    us_ca_order_details = []
    intl_order_details = []
    
    # Process all orders for product breakdown, but only create SingPost entries for international
    for region_name, region_df, counter, details in [
        ("Singapore", sg_orders, sg_product_counter, sg_order_details),
        ("US/Canada", us_ca_orders, us_ca_product_counter, us_ca_order_details),
        ("International", intl_orders, intl_product_counter, intl_order_details)
    ]:
        for _, row in region_df.iterrows():
            # Extract product details
            is_bundle, material, size = parse_product_details(row['Lineitem name'])
            
            # Create combined key for counter and store order details
            product_key = f"{material} - {size}"
            counter[product_key] += 1
            
            # Store order details with more fields for Singapore orders
            detail = {
                'name': safe_str_slice(row['Shipping Name'], 35),
                'country': safe_str_slice(row['Shipping Country'], 2),
                'size': size.split(' ')[0] if ' ' in size else size,
                'material': material,
                'is_bundle': is_bundle,
                'product_key': product_key,
                'order_number': row['Name'],  # Add order number for reference
                'quantity': 2 if is_bundle else 1
            }
            
            # Add shipping address details for Singapore orders
            if region_name == "Singapore":
                detail.update({
                    'address1': safe_str_slice(row['Shipping Address1'], 50) if pd.notna(row['Shipping Address1']) else '',
                    'address2': safe_str_slice(row['Shipping Address2'], 50) if pd.notna(row['Shipping Address2']) else '',
                    'city': safe_str_slice(row['Shipping City'], 30) if pd.notna(row['Shipping City']) else '',
                    'postal': safe_str_slice(str(row['Shipping Zip']), 10) if pd.notna(row['Shipping Zip']) else '',
                    'phone': safe_str_slice(row['Shipping Phone'], 20) if pd.notna(row['Shipping Phone']) else ''
                })
            
            details.append(detail)
            
            # Only create SingPost entries for international non-US/CA orders
            if region_name == "International":
                # Determine weight and declared value based on bundle
                if is_bundle:
                    weight = 500
                    declared_value = 100
                    height = 4
                else:
                    weight = 250
                    declared_value = 50
                    height = 2
                    
                # Handle state/province field
                if pd.notna(row['Shipping Province Name']):
                    state = str(row['Shipping Province Name'])[:30]
                elif pd.notna(row['Shipping Province']):
                    state = str(row['Shipping Province'])[:30]
                else:
                    state = ''
                    
                # Handle Address Line 2
                address_line_2 = row['Shipping Address2'][:35] if pd.notna(row['Shipping Address2']) and str(row['Shipping Address2']).strip() != '' else 'NA'
                
                # Check for address truncation
                if (pd.notna(row['Shipping Address1']) and len(str(row['Shipping Address1'])) > 35 or
                    pd.notna(row['Shipping Address2']) and len(str(row['Shipping Address2'])) > 35):
                    print(f"WARNING: Address for {row['Name']} was truncated")
                
                # Create simplified product description for SingPost
                quantity_str = "2" if is_bundle else "1"
                
                # Format size string
                if "(100-110cm)" in size:
                    size_str = "(100cm)"
                elif "(110-120cm)" in size:
                    size_str = "(110cm)"
                elif "(120-130cm)" in size:
                    size_str = "(120cm)"
                elif "(130-140cm)" in size:
                    size_str = "(130cm)"
                else:
                    # Just use first part of size (XS, S, M, L, XL)
                    size_str = size.split(" ")[0] if " " in size else size
                
                # Create simplified item description
                simplified_description = f"Eczema mitten {quantity_str}{size_str} {material}"
                
                singpost_row = {
                    'Send to business name line 1 (Max 35 characters) - *': safe_str_slice(row['Shipping Name'], 35),
                    'Send to business name line 2 (Max 35 characters)': '',
                    'Send to address line 1 (Max 35 characters) - *': safe_str_slice(row['Shipping Address1'], 35),
                    'Send to address line 2  (Max 35 characters) - *': address_line_2,
                    'Send to address line 3 (Max 35 characters)': '',
                    'Send to town (Max 30 characters) (Please spell in full)': safe_str_slice(row['Shipping City'], 30),
                    'Send to state (Max 30 characters) (Please spell in full)': state,
                    'Send to country (Max 2 characters) - *': safe_str_slice(row['Shipping Country'], 2),
                    'Send to postcode (Max 10 characters)': safe_str_slice(str(row['Shipping Zip']), 10).replace("'", ""),
                    'Sender VAT/GST number (Max 50 characters)': '',
                    'Sender Reference (Max 20 characters)': safe_str_slice(str(row['Id']), 20),
                    'Type of article - Please type in either LL (for letter) or AS (for small packet) - (Max 2 characters) - *': 'AS',
                    'Size - Please type in either RG (for Regular), LG (for Large) or NS (for Non-standard) - (Max 2 characters) - *': 'NS',
                    'Category of Shipment- Please type in either D (for Document), G (for Gift), M (for Merchandise), S (for Sample) or O (for others) (Max 1 character) - *': 'M',
                    'If "Other", please describe (Max 50 characters)': '',
                    'Total Physical weight (min 1 gm) - *': weight,
                    'Item Length (cm)': 20,
                    'Item Width (cm)': 10,
                    'Item Height (cm)': height,
                    'Service code - Refer to Service List sheet (Max 20 characters)  - *': 'IRAIRA',
                    'Currency type - for all item values (3 characters) -*': 'SGD',
                    'Item content 1 description (Max 50 characters) - *': simplified_description,
                    'Item content 1 quantity': row['Lineitem quantity'],
                    'Total content 1 weight (min 1 gm)': weight,
                    'Item content 1 total value (in declared currency type)': declared_value,
                    'Item content 1 HS tariff number (Max 6 characters)': '392620',
                    'Item content 1 Country of origin (Max 2 characters) - *': 'SG'
                }
                
                # Add empty values for content 2 and 3
                for i in [2, 3]:
                    singpost_row.update({
                        f'Item content {i} description (Max 50 characters) - *': '',
                        f'Item content {i} quantity': '',
                        f'Total content {i} weight (min 1 gm)': '',
                        f'Item content {i} total value (in declared currency type)': '',
                        f'Item content {i} HS tariff number (Max 6 characters)': '',
                        f'Item content {i} Country of origin (Max 2 characters) - *': ''
                    })
                
                singpost_data.append(singpost_row)

    # Create summary message
    summary = "ORDER DETAILS BY REGION:\n"
    
    # Function to print order details for a region
    def print_region_orders(region_name, order_details):
        output = f"\n{region_name} ORDERS:\n"
        if not order_details:
            return output + "None\n"
        
        for detail in order_details:
            quantity = detail['quantity']
            output += f"{detail['order_number']} - {detail['name']} {detail['country']}: {quantity}{detail['size']} {detail['material']}\n"
        return output
    
    summary += print_region_orders("SINGAPORE", sg_order_details)
    summary += print_region_orders("US/CANADA", us_ca_order_details)
    summary += print_region_orders("INTERNATIONAL", intl_order_details)
    
    summary += "\nPRODUCT BREAKDOWN BY REGION:"
    
    # Function to print product breakdown for a region
    def print_region_breakdown(region_name, product_counter, order_details):
        output = f"\n\n{region_name}:"
        if not product_counter:
            return output + "\nNone"
            
        # Group products by material
        cotton_products = {k: v for k, v in product_counter.items() if k.startswith('Cotton')}
        tencel_products = {k: v for k, v in product_counter.items() if k.startswith('Tencel')}
        unknown_products = {k: v for k, v in product_counter.items() if not (k.startswith('Cotton') or k.startswith('Tencel'))}
        
        total_pieces = 0
        
        # Print Cotton products
        if cotton_products:
            cotton_pieces = 0
            output += "\n\nCotton Products:"
            for product, count in sorted(cotton_products.items()):
                pieces = sum(detail['quantity'] for detail in order_details if detail['product_key'] == product)
                cotton_pieces += pieces
                output += f"\n{product.split(' - ')[1]}: {count} order{'s' if count > 1 else ''} ({pieces} pieces)"
            output += f"\nTotal Cotton pieces: {cotton_pieces}"
            total_pieces += cotton_pieces
        
        # Print Tencel products
        if tencel_products:
            tencel_pieces = 0
            output += "\n\nTencel Products:"
            for product, count in sorted(tencel_products.items()):
                pieces = sum(detail['quantity'] for detail in order_details if detail['product_key'] == product)
                tencel_pieces += pieces
                output += f"\n{product.split(' - ')[1]}: {count} order{'s' if count > 1 else ''} ({pieces} pieces)"
            output += f"\nTotal Tencel pieces: {tencel_pieces}"
            total_pieces += tencel_pieces
        
        # Print Unknown products if any
        if unknown_products:
            unknown_pieces = 0
            output += "\n\nUnknown Products:"
            for product, count in sorted(unknown_products.items()):
                pieces = sum(detail['quantity'] for detail in order_details if detail['product_key'] == product)
                unknown_pieces += pieces
                output += f"\n{product.split(' - ')[1]}: {count} order{'s' if count > 1 else ''} ({pieces} pieces)"
            output += f"\nTotal Unknown pieces: {unknown_pieces}"
            total_pieces += unknown_pieces
        
        output += f"\n\nTotal {region_name} orders: {len(order_details)}"
        output += f"\nTotal {region_name} pieces: {total_pieces}"
        return output
    
    summary += print_region_breakdown("SINGAPORE", sg_product_counter, sg_order_details)
    summary += print_region_breakdown("US/CANADA", us_ca_product_counter, us_ca_order_details)
    summary += print_region_breakdown("INTERNATIONAL", intl_product_counter, intl_order_details)
    
    # Calculate grand totals
    total_orders = len(sg_order_details) + len(us_ca_order_details) + len(intl_order_details)
    total_pieces = (
        sum(detail['quantity'] for detail in sg_order_details) +
        sum(detail['quantity'] for detail in us_ca_order_details) +
        sum(detail['quantity'] for detail in intl_order_details)
    )
    
    summary += f"\n\nGRAND TOTAL:"
    summary += f"\nTotal orders: {total_orders}"
    summary += f"\nTotal pieces: {total_pieces}"
    
    # Only create the SingPost CSV if there are international orders
    if singpost_data:
        # Convert to DataFrame and save
        singpost_df = pd.DataFrame(singpost_data)
        singpost_df.to_csv(output_file, 
                        index=False,
                        sep=',',
                        quoting=1,
                        quotechar='"',
                        escapechar='\\',
                        encoding='utf-8'
        )
        
        # Validate required fields
        required_columns = [col for col in singpost_df.columns if col.endswith('- *')]
        for col in required_columns:
            missing = singpost_df[col].isna().sum()
            if missing > 0:
                print(f"WARNING: {missing} rows have missing values in required field: {col}")
        
        summary += f"\n\nCreated SingPost file with {len(singpost_data)} international orders (excluding SG, US, CA)"
    else:
        summary += "\n\nNo international orders (excluding SG, US, CA) to export to SingPost"
    
    # Generate shipping labels with Google Slides for Singapore orders
    slides_url = None
    
    if sg_order_details:
        # Try to create Google Slides if credentials are available
        credentials_path = os.getenv('GOOGLE_CREDENTIALS_PATH')
        template_url = os.getenv('SLIDES_TEMPLATE_URL')
        
        if credentials_path and os.path.exists(credentials_path):
            template_id = get_template_id_from_url(template_url) if template_url else None
            slides_url = create_shipping_slides(sg_order_details, credentials_path, template_id)
            if slides_url:
                summary += f"\n\nCreated Google Slides presentation: {slides_url}"
        
    return summary, None, slides_url  # Return summary, PDF path (None), and Slides URL

# Example usage
if __name__ == "__main__":
    result, pdf_path, slides_url = convert_shopify_to_singpost(
        'orders_export.csv',
        'singpost_orders.csv'
    )
    print(result)
    if slides_url:
        print(f"\nGoogle Slides URL: {slides_url}")