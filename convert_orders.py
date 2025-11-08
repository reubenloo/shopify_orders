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
    """Clean Shopify export data by merging amended orders and removing incomplete orders"""
    # Extract order number from the Name column (removing the # symbol)
    df['order_number'] = df['Name'].str.replace('#', '').astype(str)

    # Find orders with multiple rows (amended orders)
    duplicate_orders = df[df.duplicated(subset=['order_number'], keep=False)]
    unique_order_numbers = duplicate_orders['order_number'].unique()

    print(f"Found {len(unique_order_numbers)} orders with amendments")

    # Lineitem columns to copy from last row to first row
    lineitem_columns = ['Lineitem quantity', 'Lineitem name', 'Lineitem price', 'Lineitem discount']

    for order_num in unique_order_numbers:
        order_rows = df[df['order_number'] == order_num]
        if len(order_rows) >= 2:
            first_row_index = order_rows.index[0]
            last_row_index = order_rows.index[-1]

            # Copy Lineitem data from last row to first row
            for col in lineitem_columns:
                if col in df.columns:
                    df.at[first_row_index, col] = df.at[last_row_index, col]

            # Drop all rows except the first one
            df = df.drop(order_rows.index[1:])
            print(f"Order #{order_num}: Merged {len(order_rows)} rows, kept amended Lineitem data")

    # Now remove rows where Financial Status is empty (these are standalone unpaid orders, not amendments)
    df = df[df['Financial Status'].notna()].copy()

    # Drop the temporary order_number column
    df = df.drop('order_number', axis=1)

    return df

def filter_international_orders(df):
    """Filter orders by destination country"""
    # Create copies to track different types of orders
    sg_orders = df[df['Shipping Country'] == 'SG'].copy()
    us_orders = df[df['Shipping Country'] == 'US'].copy()
    ca_orders = df[df['Shipping Country'] == 'CA'].copy()
    intl_orders = df[(df['Shipping Country'] != 'SG') &
                     (df['Shipping Country'] != 'US') &
                     (df['Shipping Country'] != 'CA')].copy()

    print(f"Total orders: {len(df)}")
    print(f"Singapore orders (will generate Google Slides): {len(sg_orders)}")
    print(f"US orders (will generate US CSV): {len(us_orders)}")
    print(f"Canada orders (excluded): {len(ca_orders)}")
    print(f"International orders (will generate International CSV): {len(intl_orders)}")

    return intl_orders, sg_orders, us_orders, ca_orders

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
            
    # Determine size (handle both with and without CM suffix)
    if '100-110' in lineitem_name:
        size = '(100-110cm)'
    elif '110-120' in lineitem_name:
        size = '(110-120cm)'
    elif '120-130' in lineitem_name:
        size = '(120-130cm)'
    elif '130-140' in lineitem_name:
        size = '(130-140cm)'
    elif '140-150' in lineitem_name:
        size = 'XS (140-150cm)'
    elif '150-160' in lineitem_name:
        size = 'S (150-160cm)'
    elif '160-170' in lineitem_name:
        size = 'M (160-170cm)'
    elif '170-180' in lineitem_name:
        size = 'L (170-180cm)'
    elif '180-190' in lineitem_name:
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

    # Filter orders by destination country
    intl_orders, sg_orders, us_orders, ca_orders = filter_international_orders(df)

    # Create output dataframes for SingPost
    intl_singpost_data = []
    us_singpost_data = []

    # Counters for product breakdown
    sg_product_counter = defaultdict(int)
    us_product_counter = defaultdict(int)
    ca_product_counter = defaultdict(int)
    intl_product_counter = defaultdict(int)

    sg_order_details = []
    us_order_details = []
    ca_order_details = []
    intl_order_details = []

    # Process all orders for product breakdown
    for region_name, region_df, counter, details in [
        ("Singapore", sg_orders, sg_product_counter, sg_order_details),
        ("US", us_orders, us_product_counter, us_order_details),
        ("Canada", ca_orders, ca_product_counter, ca_order_details),
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

            # Create SingPost entries for international and US orders (separate CSVs)
            if region_name in ["International", "US"]:
                # Determine currency, pricing, and HS codes based on region
                if region_name == "US":
                    currency = 'USD'
                    # US-specific pricing in USD
                    if is_bundle:
                        weight = 500
                        height = 4
                        if material == 'Cotton':
                            declared_value = 120
                        elif material == 'Tencel':
                            declared_value = 240
                        else:
                            declared_value = 120  # Default
                    else:
                        weight = 250
                        height = 2
                        if material == 'Cotton':
                            declared_value = 75
                        elif material == 'Tencel':
                            declared_value = 150
                        else:
                            declared_value = 75  # Default

                    # US-specific HS codes
                    if material == 'Cotton':
                        hs_code = '6110202020'
                    elif material == 'Tencel':
                        hs_code = '6110303020'
                    else:
                        hs_code = '6110202020'  # Default to cotton

                else:  # International
                    currency = 'SGD'
                    # International pricing in SGD
                    if is_bundle:
                        weight = 500
                        declared_value = 40
                        height = 4
                    else:
                        weight = 250
                        declared_value = 20
                        height = 2

                    # International HS codes
                    if material == 'Cotton':
                        hs_code = '611020'
                    elif material == 'Tencel':
                        hs_code = '611030'
                    else:
                        hs_code = '611020'  # Default to cotton

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
                    'Currency type - for all item values (3 characters) -*': currency,
                    'Item content 1 description (Max 50 characters) - *': simplified_description,
                    'Item content 1 quantity': row['Lineitem quantity'],
                    'Total content 1 weight (min 1 gm)': weight,
                    'Item content 1 total value (in declared currency type)': declared_value,
                    'Item content 1 HS tariff number (Max 6 characters)': hs_code,
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

                # Append to the appropriate list
                if region_name == "US":
                    us_singpost_data.append(singpost_row)
                else:  # International
                    intl_singpost_data.append(singpost_row)

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
    summary += print_region_orders("US", us_order_details)
    summary += print_region_orders("CANADA", ca_order_details)
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
    summary += print_region_breakdown("US", us_product_counter, us_order_details)
    summary += print_region_breakdown("CANADA", ca_product_counter, ca_order_details)
    summary += print_region_breakdown("INTERNATIONAL", intl_product_counter, intl_order_details)

    # Calculate grand totals
    total_orders = len(sg_order_details) + len(us_order_details) + len(ca_order_details) + len(intl_order_details)
    total_pieces = (
        sum(detail['quantity'] for detail in sg_order_details) +
        sum(detail['quantity'] for detail in us_order_details) +
        sum(detail['quantity'] for detail in ca_order_details) +
        sum(detail['quantity'] for detail in intl_order_details)
    )
    
    summary += f"\n\nGRAND TOTAL:"
    summary += f"\nTotal orders: {total_orders}"
    summary += f"\nTotal pieces: {total_pieces}"

    # Create DataFrames for CSV output
    intl_df = None
    us_df = None

    # Create International CSV if there are international orders
    if intl_singpost_data:
        intl_df = pd.DataFrame(intl_singpost_data)
        intl_df.to_csv(output_file,
                       index=False,
                       sep=',',
                       quoting=1,
                       quotechar='"',
                       escapechar='\\',
                       encoding='utf-8')

        # Validate required fields
        required_columns = [col for col in intl_df.columns if col.endswith('- *')]
        for col in required_columns:
            missing = intl_df[col].isna().sum()
            if missing > 0:
                print(f"WARNING: {missing} rows have missing values in required field: {col}")

        summary += f"\n\nCreated International SingPost file with {len(intl_singpost_data)} orders (ex-SG, ex-US, ex-CA)"
    else:
        summary += "\n\nNo international orders (ex-SG, ex-US, ex-CA) to export"

    # Create US CSV if there are US orders
    if us_singpost_data:
        # Create US-specific output filename
        us_output_file = output_file.replace('.csv', '_us.csv')
        us_df = pd.DataFrame(us_singpost_data)
        us_df.to_csv(us_output_file,
                     index=False,
                     sep=',',
                     quoting=1,
                     quotechar='"',
                     escapechar='\\',
                     encoding='utf-8')

        # Validate required fields
        required_columns = [col for col in us_df.columns if col.endswith('- *')]
        for col in required_columns:
            missing = us_df[col].isna().sum()
            if missing > 0:
                print(f"WARNING: {missing} rows have missing values in required field: {col}")

        summary += f"\n\nCreated US SingPost file with {len(us_singpost_data)} orders"
    else:
        summary += "\n\nNo US orders to export"
    
    # Generate shipping labels with Google Slides for Singapore orders
    slides_url = None

    if sg_order_details:
        # Try to create Google Slides if credentials are available
        credentials_path = os.getenv('GOOGLE_CREDENTIALS_PATH')
        template_url = os.getenv('SLIDES_TEMPLATE_URL')

        if credentials_path and os.path.exists(credentials_path):
            try:
                template_id = get_template_id_from_url(template_url) if template_url else None
                if template_id:
                    slides_url = create_shipping_slides(sg_order_details, credentials_path, template_id)
                    if slides_url:
                        summary += f"\n\nCreated Google Slides presentation: {slides_url}"
                    else:
                        summary += "\n\nError: Could not create Google Slides. Check logs for details."
                else:
                    summary += "\n\nError: Could not parse template ID from URL."
            except Exception as e:
                import traceback
                traceback.print_exc()
                summary += f"\n\nError creating Google Slides: {str(e)}"
        else:
            summary += "\n\nNo Google credentials available for Slides integration."
        
    return summary, intl_df, us_df, slides_url  # Return summary, international DF, US DF, and Slides URL

# Example usage
if __name__ == "__main__":
    result, pdf_path, slides_url = convert_shopify_to_singpost(
        'orders_export.csv',
        'singpost_orders.csv'
    )
    print(result)
    if slides_url:
        print(f"\nGoogle Slides URL: {slides_url}")