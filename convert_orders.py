# Version 2.0.0 - retail-speedpost-worldwide-multiple template (42 columns)
import pandas as pd
import os
import time
import requests
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

def fetch_usd_prices_from_shopify(us_orders):
    """
    Fetch USD prices for US orders from Shopify API

    Args:
        us_orders: DataFrame with US orders

    Returns:
        dict: {order_number: usd_price}
        Example: {'2689': 156.25, '2690': 75.00}
    """

    # Get credentials from environment
    access_token = os.getenv('SHOPIFY_ACCESS_TOKEN')
    store_url = os.getenv('SHOPIFY_STORE_URL')

    if not access_token or not store_url:
        raise Exception("Shopify API credentials not configured. Please add Shopify credentials (access_token and store_url) to .streamlit/secrets.toml")

    usd_prices = {}

    print(f"\nFetching USD prices from Shopify API for {len(us_orders)} US orders...")

    for _, row in us_orders.iterrows():
        # Extract order number from 'Name' field (e.g., '#2689' -> '2689')
        order_number = row['Name'].replace('#', '').strip()

        # Build API URL
        url = f"https://{store_url}/admin/api/2024-01/orders.json"
        params = {
            'name': order_number,  # NOT '#2689', just '2689'
            'status': 'any',
            'fields': 'current_subtotal_price_set'
        }
        headers = {
            'X-Shopify-Access-Token': access_token
        }

        # Call API with rate limiting
        time.sleep(0.5)  # 2 requests/second = 0.5s between calls

        try:
            response = requests.get(url, params=params, headers=headers, timeout=10)
            response.raise_for_status()

            data = response.json()

            if data.get('orders') and len(data['orders']) > 0:
                order = data['orders'][0]
                usd_amount = order['current_subtotal_price_set']['presentment_money']['amount']
                usd_prices[order_number] = float(usd_amount)
                print(f"✓ Order #{order_number}: USD ${usd_amount}")
            else:
                error_msg = f"Order #{row['Name']} not found in Shopify API. Please verify the order exists and try again."
                raise Exception(error_msg)

        except requests.exceptions.RequestException as e:
            error_msg = f"Error connecting to Shopify API for order #{row['Name']}: {str(e)}"
            raise Exception(error_msg)
        except KeyError as e:
            error_msg = f"Unexpected API response format for order #{row['Name']}: {str(e)}"
            raise Exception(error_msg)

    print(f"✓ Successfully fetched USD prices for all {len(us_orders)} US orders\n")
    return usd_prices

def create_intl_singpost_row(row, is_bundle, material, size, hs_code, declared_value):
    """
    Create SingPost row for International orders with 42-column retail-speedpost-worldwide-multiple template

    Args:
        row: Shopify order row
        is_bundle: Boolean
        material: 'Cotton' or 'Tencel'
        size: Size string
        hs_code: 6-digit HS code
        declared_value: SGD value (20 for single, 40 for bundle)

    Returns:
        dict: 42-column row for International CSV
    """

    # Calculate weights (in kg) and dimensions
    weight_kg = 0.5 if is_bundle else 0.25  # 500g or 250g in kg
    height = 4 if is_bundle else 2

    # Handle state/province
    if pd.notna(row['Shipping Province Name']):
        state = str(row['Shipping Province Name'])[:30]
    elif pd.notna(row['Shipping Province']):
        state = str(row['Shipping Province'])[:30]
    else:
        state = ''

    # Handle Address Line 2
    address_line_2 = row['Shipping Address2'][:35] if pd.notna(row['Shipping Address2']) and str(row['Shipping Address2']).strip() != '' else 'NA'

    # Create simplified product description
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
        size_str = size.split(" ")[0] if " " in size else size

    simplified_description = f"Eczema Bolero Shrug {quantity_str}{size_str} {material}"

    # Extract order number for reference fields
    order_number = row['Name']  # e.g., '#2689'

    # Build the 42-column row in exact template order (retail-speedpost-worldwide-multiple)
    intl_row = {
        # 1-9: Recipient Address
        'Send to business name (Max 35 characters) - *': safe_str_slice(row['Shipping Name'], 35),
        'Send to contact person (Max 35 characters) ': safe_str_slice(row['Shipping Name'], 35),
        'Send to address line 1 eg. Block no. (Max 35 characters) - *': safe_str_slice(row['Shipping Address1'], 35),
        'Send to address line 2 eg. Street name and Unit no. (Max 35 characters)- *': address_line_2,
        'Send to address line 3 eg. Building name (Max 35 Characters)': '',
        'Send to town (Max 30 characters) (Please spell in full)': safe_str_slice(row['Shipping City'], 30),
        'Send to state (Max 30 characters) (Please spell in full)': state,
        'Send to country (Max 2 characters) - * (Refer to Country List sheet)': safe_str_slice(row['Shipping Country'], 2),
        'Send to postcode (Max 10 characters)': safe_str_slice(str(row['Shipping Zip']), 10).replace("'", ""),

        # 10-13: Contact and VAT Info
        'Send to phone no. (Max 20 characters)': safe_str_slice(row['Shipping Phone'], 20) if pd.notna(row['Shipping Phone']) else '',
        'Sender VAT/GST number (Max 50 characters) - Sender IOSS number for European Union (EU) destinations, or VAT/GST number for specific destination where applicable.': '',
        'Receiver VAT/GST number (Max 50 characters)': '',
        'Sender Reference (Max 20 characters)': safe_str_slice(order_number, 20),

        # 14-15: Shipment Category
        'Category of shipment - Please type in either D (for document) M (for merchandise) S (for sample) or O (for others) (Max 1 character) - *': 'M',
        'If "Others", please describe (Max 50 characters)': '',

        # 16-19: Package Dimensions
        'Total Item physical weight (min 0.001 kg) - *': weight_kg,
        'Item Length (cm)': 20,
        'Item Width (cm)': 10,
        'Item Height (cm)': height,

        # 20-25: Item Content No. 1
        ' Item Content No. 1 Description  (Max 50 characters) - *': simplified_description,
        'Item Content No. 1 Quantity - *': row['Lineitem quantity'],
        'Item Content No. 1 Weight (min 0.001 kg) - *': weight_kg,
        ' Item Content No. 1 Declared value (SGD)': declared_value,
        'Item Content No. 1 HS tariff number (Max 6 characters)': hs_code,
        'Item Content No. 1 Country of origin (Max 2 characters) - * (Refer to Country List sheet) ': 'SG',

        # 26-31: Item Content No. 2 (BLANK)
        'Item  Content No. 2 Description (Max 50 characters) ': '',
        'Item  Content No. 2 Quantity - *': '',
        'Item  Content No. 2 Weight (min 0.001 kg) - *': '',
        'Item  Content No. 2 Declared Value (SGD)': '',
        'Item Content No. 2 HS tariff number (Max 6 characters)': '',
        'Item Content No. 2 Country of origin(Max 2 characters) (Refer to Country List sheet) \n': '',

        # 32-37: Item Content No. 3 (BLANK)
        'Item  Content No. 3 Description (Max 50 characters) ': '',
        'Item  Content No. 3 Quantity - *': '',
        'Item  Content No. 3 Weight (min 0.001 kg) - *': '',
        'Item  Content No. 3 Declared Value (SGD)': '',
        'Item Content No. 3 HS tariff number (Max 6 characters)': '',
        'Item Content No. 3 Country of origin(Max 2 characters) (Refer to Country List sheet) \n': '',

        # 38-42: Additional Info and Service
        'Enhanced Liability Amount (must be equal or lower than Declared value)': '',
        'Invoice number': order_number,
        'Certificate number': '',
        'Export license number': '',
        'Service code - Refer to Service List sheet (Max 20 characters)  - *': 'IRREPK'
    }

    return intl_row


def create_us_singpost_row(row, is_bundle, material, size, hs_code, usd_price):
    """
    Create SingPost row for US orders with 55-column ezy2ship template

    Args:
        row: Shopify order row
        is_bundle: Boolean
        material: 'Cotton' or 'Tencel'
        size: Size string
        hs_code: 10-digit HS code
        usd_price: USD price from Shopify API (float)

    Returns:
        dict: 55-column row for US CSV
    """

    # Calculate weights and dimensions
    weight_grams = 500 if is_bundle else 250  # grams
    weight_kg = weight_grams / 1000  # Convert to kg for US CSV format
    height = 4 if is_bundle else 2

    # Handle state/province
    if pd.notna(row['Shipping Province Name']):
        state = str(row['Shipping Province Name'])[:30]
    elif pd.notna(row['Shipping Province']):
        state = str(row['Shipping Province'])[:30]
    else:
        state = ''

    # Handle Address Line 2
    address_line_2 = row['Shipping Address2'][:35] if pd.notna(row['Shipping Address2']) and str(row['Shipping Address2']).strip() != '' else 'NA'

    # Create simplified product description
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
        size_str = size.split(" ")[0] if " " in size else size

    simplified_description = f"Eczema Bolero Shrug {quantity_str}{size_str} {material}"

    # Extract order number for invoice field
    order_number = row['Name']  # e.g., '#2689'

    # Build the 55-column row in exact template order
    us_row = {
        # 1-18: Sender/Receiver Info
        'Cost centre code (Max 40 characters)': '',
        'Sender VAT/GST number (Max 50 characters)': '',
        'Send to business name (Max 35 characters) - *': safe_str_slice(row['Shipping Name'], 35),
        'Send to contact person (Max 35 characters)': safe_str_slice(row['Shipping Name'], 35),
        'Send to address line 1 eg. Block no. (Max 35 characters) - *': safe_str_slice(row['Shipping Address1'], 35),
        'Send to address line 2 eg. Street name and Unit no. (Max 35 characters)- *': address_line_2,
        'Send to address line 3 eg. Building name (Max 35 Characters)': '',
        'Send to town (Max 30 characters) (Please spell in full)': safe_str_slice(row['Shipping City'], 30),
        'Send to state (Max 30 characters) (Please spell in full)': state,
        'Send to country (Max 2 characters) - * (Refer to Country List sheet)': safe_str_slice(row['Shipping Country'], 2),
        'Send to postcode (Max 10 characters)': safe_str_slice(str(row['Shipping Zip']), 10).replace("'", ""),
        'Send to phone no. (Max 20 characters)': safe_str_slice(row['Shipping Phone'], 20) if pd.notna(row['Shipping Phone']) else '',
        'Send to email address': row['Email'] if pd.notna(row['Email']) else '',
        'Issuing Country of IOSS Number (Country code 2 characters) - only required if sending to EU and IOSS Number is provided': '',
        'Receiver VAT/GST number (Max 50 characters)': '',
        'Recipient EORI ID': '',
        'Issuing Country of Recipient EORI ID (Country code 2 characters) - required if the EORI Number is provided': '',
        'Sender Reference (Max 20 characters)': safe_str_slice(order_number, 20),

        # 19-25: Package Info
        'Item Type - Please type in either D (for document) or P (for package) - (Max 1 character) - *': 'P',
        'Category of shipment - Please type in either D (for document) M (for merchandise) S (for sample) or O (for others) (Max 1 character) - *': 'M',
        'If "Others", please describe (Max 50 characters)': '',
        'Total Item physical weight (min 0.001 kg) - *': weight_kg,
        'Item Length (cm)': 20,
        'Item Width (cm)': 10,
        'Item Height (cm)': height,

        # 26-32: Item Content No. 1
        'Item Content No. 1 Description  (Max 50 characters) - *': simplified_description,
        'Item Content No. 1 Quantity': row['Lineitem quantity'],
        'Item Content No. 1 Weight (min 0.001 kg)': weight_kg,
        'Item Content No. 1 Declared Currency (Only USD) *': 'USD',
        'Item Content No. 1 Declared value *': usd_price,
        'Item Content No. 1 HS tariff number (10 characters) *': hs_code,
        'Item Content No. 1 Country of origin (Max 2 characters) - * (Refer to Country List sheet)': 'SG',

        # 33-39: Item Content No. 2 (BLANK)
        'Item  Content No. 2 Description (Max 50 characters)': '',
        'Item  Content No. 2 Quantity': '',
        'Item Content No. 2 Weight (min 0.001 kg)': '',
        'Item Content No. 2 Declared Currency *': '',
        'Item content No. 2 Declared value': '',
        'Item Content No. 2 HS tariff number (10 characters) *': '',
        'Item Content No. 2 Country of origin(Max 2 characters) (Refer to Country List sheet)': '',

        # 40-46: Item Content No. 3 (BLANK)
        'Item  Content No. 3 Description (Max 50 characters)': '',
        'Item  Content No. 3 Quantity': '',
        'Item Content No. 3 Weight (min 0.001 kg)': '',
        'Item Content No. 3 Declared Currency *': '',
        'Item content No. 3 Declared value': '',
        'Item Content No. 3 HS tariff number (10 characters) *': '',
        'Item Content No. 3 Country of origin(Max 2 characters) (Refer to Country List sheet)': '',

        # 47-51: Additional Info
        'Enhanced Liability Amount (must be equal or lower than Declared value)': '',
        'Invoice number': order_number,
        'Certificate number': '',
        'Export license number': '',
        'Service code - Refer to Service List sheet (Max 20 characters)  - *': 'WWCCOM',

        # 52-55: Receiver Details
        'Receiver ID Type': '',
        'Receiver ID Number (Max 50 characters)': '',
        'Tax ID (Max 35 characters)': '',
        'Product URL (Max 100 characters)': ''
    }

    return us_row

def convert_shopify_to_singpost(shopify_file, output_file):
    # Check if input file exists
    if not os.path.exists(shopify_file):
        return f"Error: Input file '{shopify_file}' not found. Please check the file path.", None, None
    
    # Read and clean Shopify orders
    df = pd.read_csv(shopify_file)
    df = clean_shopify_data(df)

    # Filter orders by destination country
    intl_orders, sg_orders, us_orders, ca_orders = filter_international_orders(df)

    # Fetch USD prices for US orders from Shopify API
    us_usd_prices = {}
    if len(us_orders) > 0:
        try:
            us_usd_prices = fetch_usd_prices_from_shopify(us_orders)
        except Exception as e:
            # Stop entire process and show error
            error_msg = f"Error fetching USD prices from Shopify API:\n{str(e)}"
            return error_msg, None, None, None

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

            # Create SingPost entries for US orders (NEW 55-column format)
            if region_name == "US":
                # Extract order number
                order_num = row['Name'].replace('#', '').strip()
                usd_price = us_usd_prices.get(order_num)

                # US-specific HS codes (10 digits)
                if material == 'Cotton':
                    hs_code = '6114200060'
                elif material == 'Tencel':
                    hs_code = '6114303070'
                else:
                    hs_code = '6114200060'  # Default to cotton

                # Create US-specific row with 55 columns
                us_row = create_us_singpost_row(row, is_bundle, material, size, hs_code, usd_price)
                us_singpost_data.append(us_row)

            # Create SingPost entries for International orders (retail-speedpost-worldwide-multiple template)
            elif region_name == "International":
                # International pricing in SGD
                declared_value = 40 if is_bundle else 20

                # International HS codes (6 digits)
                if material == 'Cotton':
                    hs_code = '611420'
                elif material == 'Tencel':
                    hs_code = '611430'
                else:
                    hs_code = '611420'  # Default to cotton

                # Check for address truncation
                if (pd.notna(row['Shipping Address1']) and len(str(row['Shipping Address1'])) > 35 or
                    pd.notna(row['Shipping Address2']) and len(str(row['Shipping Address2'])) > 35):
                    print(f"WARNING: Address for {row['Name']} was truncated")

                # Create international row with 42-column speedpost template
                intl_row = create_intl_singpost_row(row, is_bundle, material, size, hs_code, declared_value)
                intl_singpost_data.append(intl_row)

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

        # Validate column count (should be exactly 42 columns for retail-speedpost-worldwide-multiple template)
        if len(intl_df.columns) != 42:
            error_msg = f"International CSV validation error: Expected 42 columns, got {len(intl_df.columns)}"
            print(f"ERROR: {error_msg}")
            raise ValueError(error_msg)

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

        # Validate column count (should be exactly 55 columns for US template)
        if len(us_df.columns) != 55:
            error_msg = f"US CSV validation error: Expected 55 columns, got {len(us_df.columns)}"
            print(f"ERROR: {error_msg}")
            raise ValueError(error_msg)

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