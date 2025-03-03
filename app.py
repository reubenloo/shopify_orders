# Update this part of app.py where the conversion is happening
# Replace this block in your app.py file

# Run conversion
try:
    result, pdf_path, slides_url, debug_log = convert_shopify_to_singpost('orders_export.csv', 'singpost_orders.csv')
    
    # Debug: Print results
    print(f"Conversion result - pdf_path: {pdf_path}, slides_url: {slides_url}")
    
    # Create an expander for post-conversion debugging
    with st.expander("Post-Conversion Debug Info"):
        st.code(f"PDF Path: {pdf_path}\nSlides URL: {slides_url}")
    
    # Display results
    st.success("Conversion completed!")
    
    # Create tabs for summary, data preview, and debug log
    tab1, tab2, tab3 = st.tabs(["Summary", "Data Preview", "Google Slides Debug Log"])
    
    with tab1:
        st.text_area("Conversion Summary", result, height=400)
    
    with tab2:
        if os.path.exists("singpost_orders.csv"):
            df = pd.read_csv("singpost_orders.csv")
            st.dataframe(df)
        else:
            st.info("No SingPost CSV was generated (no international orders)")
    
    with tab3:
        if debug_log:
            st.text_area("Debug Log", debug_log, height=600)
        else:
            st.info("No debug information available")
    
    # Download section
    st.divider()
    st.subheader("Download Results")
    
    col1, col2, col3 = st.columns(3)
    
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
            st.info("No SingPost CSV generated")
    
    with col2:
        if pdf_path and os.path.exists(pdf_path):
            with open(pdf_path, "rb") as file:
                st.download_button(
                    label="Download Shipping Labels PDF",
                    data=file,
                    file_name=os.path.basename(pdf_path),
                    mime="application/pdf"
                )
        else:
            st.info("No PDF shipping labels generated")
    
    with col3:
        if slides_url:
            st.markdown(f"[Open Google Slides Shipping Labels]({slides_url})")
        else:
            if ('credentials_path' in st.session_state and 
                st.session_state.credentials_path and 
                os.path.exists(st.session_state.credentials_path)):
                st.warning("No Google Slides generated. Check service account permissions.")
            else:
                st.warning("Google credentials not found. Upload Google credentials to use Slides integration.")
                        
except Exception as e:
    st.error(f"Error processing orders: {str(e)}")
    with st.expander("Detailed Error Information"):
        st.code(traceback.format_exc())
    print(f"Conversion error: {e}")
    print(traceback.format_exc())