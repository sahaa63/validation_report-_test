import streamlit as st
import pandas as pd
import os
from io import BytesIO
import base64 # For base64 image encoding
# openpyxl is needed for pd.ExcelWriter engine='openpyxl'
# Although not directly used in the logic shown, ensure it's installed
# from openpyxl.styles import PatternFill, Font # Not needed for this app's logic
# from openpyxl.utils import get_column_letter # Not needed for this app's logic
# import numpy as np # Not needed for this app's logic

# Page config
# Changed layout from "wide" to "centered"
st.set_page_config(page_title="Standardiser", layout="centered")

# Custom CSS for styling (from the first code)
st.markdown("""
    <style>
    .title {
        font-size: 36px;
        color: #FF4B4B; /* Using the color from the first code's title class */
        text-align: center;
        font-weight: bold;
        margin-bottom: 20px;
    }
    .instructions {
        background-color: #F0F8FF;
        color: #333333;
        padding: 15px;
        border-radius: 10px;
        border-left: 5px solid #4682B4;
        margin-bottom: 20px;
    }
    .file-list {
        background-color: #F5F5F5;
        color: #333333;
        padding: 10px;
        border-radius: 5px;
        margin-top: 10px;
        margin-bottom: 10px; /* Added margin bottom for spacing */
    }
    /* Button styling from the first code */
    .stButton>button {
        background-color: #4CAF50;
        color: white;
        border: none;
        padding: 10px 20px;
        border-radius: 5px;
        font-weight: bold;
    }
    .stButton>button:hover {
        background-color: #45A049;
    }
    .success-box {
        background-color: #E6FFE6; /* Lighter green */
        color: #333333;
        padding: 15px;
        border-radius: 10px;
        border-left: 5px solid #2ECC71; /* Darker green border */
        margin-top: 20px;
        margin-bottom: 20px; /* Added margin bottom */
    }
    .error-box {
        background-color: #FFE6E6; /* Lighter red */
        color: #333333;
        padding: 15px;
        border-radius: 10px;
        border-left: 5px solid #FF4B4B; /* Red border */
        margin-top: 20px;
        margin-bottom: 20px; /* Added margin bottom */
    }
    </style>
""", unsafe_allow_html=True)

# -------------------------------
# Header: Title
# Using the custom CSS class for the title
st.markdown('<div class="title">Standardiser</div>', unsafe_allow_html=True)

# -------------------------------
# Instructions (Styled like the first code)
# -------------------------------
st.markdown("""
    <div class="instructions">
    <h3 style="color: #4682B4;">How to Use:</h3>
    <ul>
        <li>Upload an Excel file.</li>
        <li>Ensure the file contains sheets named "excel" and "PBI".</li>
        <li>Columns common to both sheets will be standardized (numeric, date, or string).</li>
        <li>Download the new Excel file with standardized data.</li>
    </ul>
    </div>
    """, unsafe_allow_html=True)


# -------------------------------
# File Upload
# -------------------------------
st.markdown("### 📤 Upload Excel File") # Keep subheader for clarity
uploaded_file = st.file_uploader(
    "Upload an Excel file containing sheets named 'excel' and 'PBI'",
    type=["xlsx"]
)

# -------------------------------
# Main Processing
# -------------------------------
if uploaded_file:
    # Indicate uploaded file name using the file-list style
    st.markdown(f'<div class="file-list"><strong>Uploaded File:</strong> {uploaded_file.name}</div>', unsafe_allow_html=True)

    with st.spinner("Standardizing your data..."): # Added a spinner similar to the first code
        try:
            # Read sheets
            xl = pd.ExcelFile(uploaded_file)
            df_excel = xl.parse('excel')
            df_pbi = xl.parse('PBI')

            # Common columns
            common_columns = [col for col in df_excel.columns if col in df_pbi.columns]

            # Standardize function
            def standardize_column_data(df1, df2, common_columns):
                for col in common_columns:
                    # Numeric
                    if pd.api.types.is_numeric_dtype(df1[col]) and pd.api.types.is_numeric_dtype(df2[col]):
                        df1[col] = pd.to_numeric(df1[col], errors='coerce')
                        df2[col] = pd.to_numeric(df2[col], errors='coerce')

                    # Date
                    elif pd.api.types.is_datetime64_any_dtype(df1[col]) or pd.api.types.is_datetime64_any_dtype(df2[col]):
                        df1[col] = pd.to_datetime(df1[col], errors='coerce').dt.date
                        df2[col] = pd.to_datetime(df2[col], errors='coerce').dt.date

                    # String
                    else:
                        df1[col] = df1[col].astype(str).str.strip()
                        df2[col] = df2[col].astype(str).str.strip()

                return df1, df2

            # Apply standardization
            df_excel_std, df_pbi_std = standardize_column_data(df_excel.copy(), df_pbi.copy(), common_columns)

            # Output Excel in memory
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_excel_std.to_excel(writer, sheet_name='excel', index=False)
                df_pbi_std.to_excel(writer, sheet_name='PBI', index=False)
            output.seek(0)

            # Filename setup
            original_name = os.path.splitext(uploaded_file.name)[0]
            output_filename = f"{original_name}_std.xlsx"

            # Success message using custom styled div
            st.markdown(
                 f'<div class="success-box">✅ Standardization complete. Download the standardized file below:</div>',
                 unsafe_allow_html=True
            )

            # Download button
            st.download_button(
                label="📥 Download Standardized Excel", # Kept original label for clarity
                data=output,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except ValueError as e:
            # Error message using custom styled div
            st.markdown(
                 f'<div class="error-box">⚠️ Sheet error: {e}</div>',
                 unsafe_allow_html=True
            )
        except Exception as e:
            # Exception message using custom styled div
             st.markdown(
                 f'<div class="error-box">🚨 Unexpected error: {e}</div>',
                 unsafe_allow_html=True
             )

# -------------------------------
# Footer (Styled like the first code)
# Function to encode local image as base64 (from the first code)
def get_base64_image(image_path):
    try:
        with open(image_path, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode()
    except FileNotFoundError:
        return None # Return None if file not found

st.markdown("---") # Separator

# Image handling for footer
image_base64 = get_base64_image("Sigmoid_Logo.jpg")
except_message = "" # Initialize message here

if image_base64:
    image_src = f"data:image/jpeg;base64,{image_base64}"
    # except_message remains ""
else:
    # Use a placeholder or handle missing image differently
    # Using the URL from the original second code as a fallback if local fails
    image_src = "https://sigmoidanalytics.com/wp-content/uploads/2021/10/Sigmoid_Logo.png"
    except_message = "<p style='color: orange; font-size: 0.8em;'>Note: Local Sigmoid_Logo.jpg not found, using web image.</p>"


footer_html = f"""
    <div style='background-color: #FFFFFF; color: #000000; padding: 20px; border-radius: 10px; box-shadow: 0 4px 8px rgba(0,0,0,0.2); margin-top: 30px; position: relative;'>
        <img src="{image_src}" alt="Sigmoid Logo" style='position: absolute; top: 10px; left: 10px; width: 100px; height: auto; border-radius: 5px;'>
        <div style='margin-left: 120px;'>
            <p style='font-size: 16px; font-weight: bold; margin: 10px 0 5px 0;'>Contact Us</p>
            <p style='font-size: 14px; margin: 0;'>
                Email: <a href='mailto:arkaprova@sigmoidanalytics.com' style='color: #1E90FF;'>arkaprova@sigmoidanalytics.com</a><br>
                Phone: <span style='color: #FFD700;'>+91 9330492917</span><br>
                Website: <a href='https://github.com/sahaa63/validation_report-_test' style='color: #1E90FF;'>Github</a>
            </p>
             {except_message}
    </div>
    """
st.markdown(footer_html, unsafe_allow_html=True)