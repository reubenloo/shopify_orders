import os
from fpdf import FPDF
from datetime import datetime

class LabelPDF(FPDF):
    def __init__(self):
        # Initialize PDF with dimensions: 150mm wide x 100mm high (horizontal)
        super().__init__(orientation='L', unit='mm', format=(100, 150))
        self.set_auto_page_break(auto=False)
        self.set_margins(5, 5, 5)
    
    def add_label(self, order):
        """Create a shipping label for a single order"""
        self.add_page()
        
        # Set font
        self.set_font('Helvetica', 'B', 16)
        
        # Header: To/From
        self.cell(70, 10, "To:", 1, 0, 'L')
        self.cell(70, 10, "From:", 1, 1, 'L')
        
        # Name & Company
        self.set_font('Helvetica', '', 12)
        self.cell(70, 15, f"Name: {order['order_number']} {order['name']}", 1, 0, 'L')
        self.cell(70, 15, "Company: Eczema Mitten\nPrivate Limited", 1, 1, 'L')
        
        # Contact
        # For demo we're using placeholder phone - you should add a real field in your data
        recipient_phone = order.get('phone', '+65 9XXX XXXX')
        self.cell(70, 10, f"Contact: {recipient_phone}", 1, 0, 'L')
        self.cell(70, 10, "Contact: +65 8889 5607", 1, 1, 'L')
        
        # Addresses
        address_line1 = order.get('address1', '')
        address_line2 = order.get('address2', '')
        address = f"{address_line1}\n{address_line2}" if address_line2 else address_line1
        
        self.cell(70, 15, f"Delivery Address:\n{address}", 1, 0, 'L')
        self.cell(70, 15, "Return Address: #04-23, Block\n235, Choa Chu Kang Central", 1, 1, 'L')
        
        # Postal
        postal_code = order.get('postal', '')
        self.cell(70, 10, f"Postal: {postal_code}", 1, 0, 'L')
        self.cell(70, 10, "Postal: 680235", 1, 1, 'L')
        
        # Item details - spans the whole width
        quantity = "2" if order['is_bundle'] else "1"
        size_display = order['size'].replace('(', '').replace(')', '').replace('-', '')
        if "cm" in size_display:
            size_display = size_display.split(' ')[0]
        
        item_text = f"Item: {quantity} {size_display} {order['material']} Eczema Mitten"
        self.cell(140, 10, item_text, 1, 1, 'L')


def generate_sg_shipping_labels(sg_order_details, output_folder="shipping_labels"):
    """Generate PDF shipping labels for Singapore orders"""
    try:
        # Create output folder if it doesn't exist
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
        
        # Create timestamp for filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_pdf = os.path.join(output_folder, f"sg_labels_{timestamp}.pdf")
        
        # Initialize PDF
        pdf = LabelPDF()
        
        # Process each Singapore order
        for order in sg_order_details:
            # Enhance order data with address info
            # You'll need to modify this to include the actual address fields
            # from your Shopify export data
            enhanced_order = order.copy()
            
            # Add label to PDF
            pdf.add_label(enhanced_order)
        
        # Save PDF
        pdf.output(output_pdf)
        print(f"Created shipping labels PDF at {output_pdf}")
        return output_pdf
    
    except Exception as e:
        print(f"Error creating PDF: {str(e)}")
        return None