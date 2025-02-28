import streamlit as st
import pandas as pd
import os
from convert_orders import convert_shopify_to_singpost

st.set_page_config(
    page_title="Shopify to SingPost Converter",
    page_icon="📦",
    layout="wide"
)

st.title("Shopify to SingPost Converter")
st.write("Upload your Shopify order export CSV to convert it to SingPost ezy2ship format")

# File uploader
uploaded_file = st.file_uploader("Upload Shopify orders_export.csv", type='csv')

if uploaded_file:
    # Save the uploaded file
    with open("orders_export.csv", "wb") as f:
        f.write(uploaded_file.getvalue())
    
    # Add a button to run the conversion
    if st.button("Convert to SingPost Format"):
        with st.spinner("Processing..."):
            # Run conversion
            result, sg_labels_pdf = convert_shopify_to_singpost('orders_export.csv', 'singpost_orders.csv')
            
            # Display results
            st.success("Conversion completed!")
            
            # Create tabs for summary, data preview, and shipping labels
            tab1, tab2 = st.tabs(["Summary", "Data Preview"])
            
            with tab1:
                st.text_area("Conversion Summary", result, height=400)
            
            with tab2:
                if os.path.exists("singpost_orders.csv"):
                    df = pd.read_csv("singpost_orders.csv")
                    st.dataframe(df)
            
            # Provide download links
            col1, col2 = st.columns(2)
            
            with col1:
                if os.path.exists("singpost_orders.csv"):
                    with open("singpost_orders.csv", "rb") as file:
                        st.download_button(
                            label="Download SingPost CSV",
                            data=file,
                            file_name="singpost_orders.csv",
                            mime="text/csv"
                        )
                else:
                    st.warning("No SingPost CSV was generated. There might be no international orders to process.")
            
            with col2:
                if sg_labels_pdf and os.path.exists(sg_labels_pdf):
                    with open(sg_labels_pdf, "rb") as file:
                        st.download_button(
                            label="Download Singapore Shipping Labels (PDF)",
                            data=file,
                            file_name=os.path.basename(sg_labels_pdf),
                            mime="application/pdf"
                        )
                else:
                    st.info("No Singapore orders to generate shipping labels.")