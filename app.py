# app.py - Streamlit App for ICHRA Document Generation
import streamlit as st
import os
from io import BytesIO

# Import your existing functions
from allowance import create_ichra_document

# Page config
st.set_page_config(
    page_title="ICHRA Document Generator",
    page_icon="📄",
    layout="centered"
)

# Hide Streamlit branding
hide_st_style = """
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
</style>
"""
st.markdown(hide_st_style, unsafe_allow_html=True)

# Main app
st.title("📄 ICHRA Document Generator")
st.markdown("---")

# Simple instructions
st.markdown("### Upload your allowance CSV file to generate a formatted document")

# File uploader
uploaded_file = st.file_uploader(
    "Choose your CSV file",
    type=['csv'],
    help="Upload the CSV file containing your allowance data"
)

if uploaded_file is not None:
    # Show file details
    st.success(f"✅ File uploaded: {uploaded_file.name}")

    # Generate button
    if st.button("Generate Document", type="primary", use_container_width=True):
        with st.spinner("Creating your document..."):
            try:
                # Save uploaded file temporarily
                temp_csv = "temp_allowance.csv"
                with open(temp_csv, "wb") as f:
                    f.write(uploaded_file.getbuffer())

                # Generate document
                output_file = "ICHRA_Allowance_Model.docx"
                create_ichra_document(
                    output_filename=output_file,
                    header_image_path="zorro_header.png",  # This file is in your project
                    csv_path=temp_csv
                )

                # Read the generated file
                with open(output_file, "rb") as f:
                    doc_bytes = f.read()

                # Clean up temp files
                os.remove(temp_csv)
                os.remove(output_file)

                # Offer download
                st.success("✅ Document generated successfully!")
                st.download_button(
                    label="📥 Download Your Document",
                    data=doc_bytes,
                    file_name="ICHRA_Allowance_Model.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )

            except Exception as e:
                st.error(f"Error: {str(e)}")
                st.info(
                    "Please ensure your CSV has all required columns: ageFrom, ageTo, class, EE, ES, EC1, EC2, ECmax, FA1, FA2, FAmax")

# Footer
st.markdown("---")
st.markdown("💡 Need help? Make sure your CSV has all required columns for allowance data.")