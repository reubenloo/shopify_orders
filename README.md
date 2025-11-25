# Shopify to SingPost Order Converter

A Streamlit web application that automates order fulfillment for an eczema mittens e-commerce business. Converts Shopify order exports into SingPost ezy2ship format and generates shipping labels via Google Slides.

## Overview

This application processes Shopify order exports and:
1. **International Orders (ex-SG, ex-US, ex-CA)**: Generates SingPost ezy2ship CSV with SGD pricing
2. **US Orders**: Generates separate SingPost ezy2ship CSV with USD pricing and US-specific HS codes
3. **Singapore Orders**: Creates Google Slides shipping labels with customer details

## File Structure

```
shopify_orders/
├── app.py                    # Streamlit web interface
├── convert_orders.py         # Core business logic for order processing
├── google_slides.py          # Google Slides API integration for SG labels
├── requirements.txt          # Python dependencies
├── .gitignore               # Excludes credentials, CSVs, etc.
└── README.md                # This file
```

## Business Logic Flow

### 1. Data Cleaning (`clean_shopify_data`)

**Problem**: When orders are amended in Shopify (e.g., size changes), the CSV export contains multiple rows per order:
- First row: Original order with complete customer/shipping data
- Subsequent rows: Amendments with only Lineitem fields populated

**Solution**:
```python
For each order with multiple rows:
  1. Keep FIRST row (has all customer/shipping/billing data)
  2. Copy from LAST row to FIRST row:
     - Lineitem quantity
     - Lineitem name (contains product details)
     - Lineitem price
     - Lineitem discount
  3. Delete all other rows
  4. Filter out unpaid orders (empty Financial Status)
```

**Example**:
```
Order #2692 (3 rows in CSV):
  Row 1: Cotton Single XS [FULL DATA]
  Row 2: Tencel Bundle L [LINEITEM ONLY]
  Row 3: Tencel Single M [LINEITEM ONLY] ← FINAL AMENDMENT

Result after merge:
  Row 1: Tencel Single M [FULL DATA + FINAL LINEITEM]
```

### 2. Order Filtering (`filter_international_orders`)

Orders are categorized by shipping destination:

- **Singapore (SG)**: Google Slides labels only, no CSV
- **United States (US)**: Separate CSV with USD pricing
- **Canada (CA)**: Excluded from processing (manual handling)
- **International (all others)**: Standard CSV with SGD pricing

### 3. Product Parsing (`parse_product_details`)

Extracts from `Lineitem name` field:

**Format**: `Eczema Bolero Shrug - {Material} / {Type} / {Size}`

**Extracted Data**:
- **Material**: Cotton or Tencel (Premium)
- **Bundle Detection**: "Bundle of 2" or "2 PAIRS" → 2 pieces
- **Size Parsing**: Matches patterns like `140-150`, `(160-170)`, `L (170-180)`, etc.

**Size Mapping**:
```
100-110 → (100-110cm) - Kid
110-120 → (110-120cm) - Kid
120-130 → (120-130cm) - Kid
130-140 → (130-140cm) - Kid
140-150 → XS (140-150cm)
150-160 → S (150-160cm)
160-170 → M (160-170cm)
170-180 → L (170-180cm)
180-190 → XL (180-190cm)
```

### 4. CSV Generation - International Orders

**Output File**: `singpost_orders.csv`

**Pricing (SGD)**:
- Single pair: $20
- Bundle (2 pairs): $40

**Weight**:
- Single: 250g, 2cm height
- Bundle: 500g, 4cm height

**HS Tariff Codes**:
- Cotton: `611420`
- Tencel: `611430`

**Service Details**:
- Service Code: `IRAIRA`
- Category: Merchandise (M)
- Type: Small packet (AS)
- Size: Non-standard (NS)
- Dimensions: 20cm × 10cm × (2cm or 4cm)
- Country of Origin: SG

### 5. CSV Generation - US Orders

**Output File**: `singpost_orders_us.csv`

**Pricing (USD)** - Higher declared values for US customs:
- Cotton Single: $75
- Tencel Single: $150
- Cotton Bundle: $120
- Tencel Bundle: $240

**HS Tariff Codes** (US-specific 10-digit):
- Cotton: `6114200060` (assumes women's garment)
- Tencel: `6114303070` (assumes women's garment)

**Other Details**: Same as international (weight, service code, etc.) but currency is USD

### 6. Google Slides Integration (Singapore Orders Only)

**Requirements**:
- Google service account credentials
- Google Slides template URL
- Template must be shared with service account email

**Process**:
1. Duplicates template slide for each order
2. Replaces placeholders with order data:
   - `#ORDERNUM#` → Order number
   - `#CUSTOMERNAME#` → Customer name
   - `#PHONE#` → Formatted phone (+65 XXXX XXXX)
   - `#ADDRESS#` → Address lines
   - `#POSTALCODE#` → Postal code
   - `#QUANTITY#` → 1 or 2 (bundle)
   - `#SIZE#` → Size letter (XS, S, M, L, XL)
   - `#MATERIAL#` → Cotton or Tencel

## Configuration

### Local Development

Create `.streamlit/secrets.toml`:

```toml
[google_credentials]
type = "service_account"
project_id = "your-project-id"
private_key_id = "your-key-id"
private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
client_email = "service-account@project.iam.gserviceaccount.com"
client_id = "client-id"
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "https://www.googleapis.com/robot/v1/metadata/x509/..."
universe_domain = "googleapis.com"

[google_slides_template]
url = "https://docs.google.com/presentation/d/YOUR_TEMPLATE_ID/edit"
```

### Streamlit Cloud Deployment

1. Add secrets in Streamlit Cloud app settings (same format as above)
2. Secrets are automatically excluded from git (never committed)
3. CSV files are automatically gitignored

## Installation

```bash
pip install -r requirements.txt
```

**Dependencies**:
- `streamlit` - Web interface
- `pandas` - Data processing
- `google-api-python-client` - Google Slides API
- `google-auth-httplib2` - Google authentication
- `google-auth-oauthlib` - OAuth support

## Usage

### Run Locally

```bash
streamlit run app.py
```

### Processing Orders

1. Upload Shopify orders export CSV
2. Click "Convert to SingPost Format"
3. Review summary and order breakdown
4. Download CSV files:
   - **International CSV**: For non-US/CA international orders
   - **US CSV**: For US orders only
5. Access Google Slides link (if Singapore orders exist)

## Expected Output

### Summary Report

```
ORDER DETAILS BY REGION:

SINGAPORE ORDERS:
#1234 - Customer Name SG: 1M Cotton

US ORDERS:
#5678 - Customer Name US: 2L Tencel

INTERNATIONAL ORDERS:
#9012 - Customer Name GB: 1S Cotton

PRODUCT BREAKDOWN BY REGION:
[Detailed breakdown by material and size]

GRAND TOTAL:
Total orders: 3
Total pieces: 4
```

### CSV Files

Both CSVs contain SingPost ezy2ship format with all required fields:
- Recipient information (name, address, country)
- Package details (weight, dimensions, service code)
- Item contents (description, quantity, value, HS code)

## Data Flow Diagram

```
Shopify CSV Export
        ↓
Clean & Merge Amended Orders
        ↓
    Parse Products
        ↓
   Filter by Country
        ↓
    ┌───┴───┬─────────┬──────────┐
    ↓       ↓         ↓          ↓
   SG      US      Canada    International
    ↓       ↓         ↓          ↓
Google   USD CSV   (Skip)    SGD CSV
Slides
```

## Security Considerations

### Safe to Make Public ✅

The codebase can be safely made public because:

1. **No Hardcoded Secrets**: All credentials are in Streamlit secrets (not in code)
2. **Gitignore Configured**:
   - `google_credentials.json` excluded
   - `*.csv` excluded (customer data)
   - `*.json` excluded
   - `.env` files excluded
3. **Business Logic Only**: Code contains workflow logic, not sensitive data

### What IS in the Code (Non-Sensitive):

- Pricing structure ($20/$40 SGD, $75/$150/$120/$240 USD)
- HS tariff codes (public information)
- Product types (eczema mittens, materials, sizes)
- Shipping service details (SingPost service codes)
- Template placeholder names

### What is NOT in the Code:

- ❌ Customer data (names, addresses, phone numbers)
- ❌ Google service account credentials
- ❌ API keys or secrets
- ❌ Actual order data

### Before Making Public:

1. ✅ Verify `.gitignore` includes all sensitive file types
2. ✅ Check git history doesn't contain committed secrets:
   ```bash
   git log --all --full-history -- google_credentials.json
   git log --all --full-history -- "*.csv"
   ```
3. ✅ Ensure no test CSVs with real customer data are committed
4. ✅ Review commit messages for any sensitive information

## Troubleshooting

### Size Shows as "Unknown"

- Check that Lineitem name contains size in format: `(140-150)` or `XS (140-150)`
- Verify size parsing patterns match your product naming

### Orders Not Merging Correctly

- Ensure orders have same `Name` field (e.g., `#2692`)
- Check that Financial Status exists in first row

### Google Slides Not Generating

- Verify credentials file path is set
- Ensure template is shared with service account email
- Check template URL is valid Google Slides presentation

### Wrong Currency in CSV

- US orders: Should see `USD` in currency field
- International: Should see `SGD` in currency field
- Verify `Shipping Country` column in Shopify export

## Maintenance Notes for Future AI

### When Pricing Changes:

Edit `convert_orders.py` lines:
- International: ~214-215 (SGD pricing)
- US: ~181-195 (USD pricing by material)

### When Adding New Countries to US CSV:

Edit `convert_orders.py` line ~116:
```python
us_orders = df[df['Shipping Country'].isin(['US', 'CA'])].copy()
```

### When HS Codes Change:

Edit `convert_orders.py` lines:
- International: ~218-223 (6-digit codes)
- US: ~198-203 (10-digit codes)

### When Adding New Sizes:

Edit `convert_orders.py` lines ~89-109 (size parsing logic)

## License

Private use only. Not for redistribution.

## Support

For issues or questions, contact the repository owner.
