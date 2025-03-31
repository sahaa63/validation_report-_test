import streamlit as st
import pandas as pd
import io
import os
from openpyxl import Workbook  # Changed from load_workbook to Workbook for new file creation

# Custom CSS for styling with improved contrast
st.markdown("""
    <style>
    .title {
        font-size: 36px;
        color: #FF4B4B;
        text-align: center;
        font-weight: bold;
        margin-bottom: 20px;
    }
    .instructions {
        background-color: #F0F8FF;  /* Light blue background */
        color: #333333;  /* Dark gray text for contrast */
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
    }
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
        background-color: #E6FFE6;
        color: #333333;
        padding: 15px;
        border-radius: 10px;
        border-left: 5px solid #2ECC71;
        margin-top: 20px;
    }
    .error-box {
        background-color: #FFE6E6;
        color: #333333;
        padding: 15px;
        border-radius: 10px;
        border-left: 5px solid #FF4B4B;
        margin-top: 20px;
    }
    </style>
""", unsafe_allow_html=True)

def combine_excel_files(file_list):
    if not file_list or len(file_list) > 10:
        return None, None

    # Extract original filename from the first file
    first_filename = os.path.splitext(file_list[0].name)[0]
    original_filename = first_filename[:-1] if first_filename[-1].isdigit() else first_filename
    output_filename = f"{original_filename}_validation_report.xlsx"

    # Create a new workbook in memory
    output_buffer = io.BytesIO()
    output_wb = Workbook()  # Create a new blank workbook

    # Dictionary to track sheet names and avoid duplicates
    sheet_name_count = {}

    # Process each uploaded file
    for uploaded_file in file_list:
        file_bytes = uploaded_file.read()
        wb = load_workbook(filename=io.BytesIO(file_bytes))

        for sheet_name in wb.sheetnames:
            base_sheet_name = sheet_name
            if sheet_name in sheet_name_count:
                sheet_name_count[sheet_name] += 1
                new_sheet_name = f"{base_sheet_name}_{sheet_name_count[sheet_name]}"
            else:
                sheet_name_count[sheet_name] = 0
                new_sheet_name = sheet_name

            ws_source = wb[base_sheet_name]
            ws_target = output_wb.create_sheet(title=new_sheet_name)
            for row in ws_source.rows:
                for cell in row:
                    ws_target[cell.coordinate].value = cell.value

    # Remove default sheet if it exists (created by Workbook())
    if 'Sheet' in output_wb.sheetnames:
        output_wb.remove(output_wb['Sheet'])

    # Save to buffer
    output_wb.save(output_buffer)
    output_buffer.seek(0)

    return output_buffer, output_filename

def main():
    # Title with custom styling
    st.markdown('<div class="title">Excel File Merger</div>', unsafe_allow_html=True)

    # Instructions box with improved contrast
    st.markdown("""
    <div class="instructions">
    <h3 style="color: #4682B4;">How to Use:</h3>
    <ul>
        <li>Upload up to 10 Excel files using the button below.</li>
        <li>All sheets from each file will be merged into one awesome output file.</li>
        <li>Duplicate sheet names will get a cool numeric suffix (e.g., 'Sheet_1').</li>
        <li>The output file will be named based on your first file (e.g., 'Report_validation_report.xlsx').</li>
    </ul>
    </div>
    """, unsafe_allow_html=True)

    # File uploader with colorful styling
    uploaded_files = st.file_uploader(
        "Drop Your Excel Files Here!",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        help="Upload up to 10 Excel files to merge into one.",
        key="file_uploader"
    )

    if uploaded_files:
        if len(uploaded_files) > 10:
            st.markdown(
                '<div class="error-box">Whoops! Maximum 10 files allowed. Please upload fewer files.</div>',
                unsafe_allow_html=True
            )
        else:
            # Display uploaded files in a styled box
            st.markdown(f'<div class="file-list"><strong>Uploaded {len(uploaded_files)} File(s):</strong>', unsafe_allow_html=True)
            for file in uploaded_files:
                st.markdown(f"- {file.name}", unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

            # Combine files and provide download
            with st.spinner("Merging your files... Hang tight!"):
                result = combine_excel_files(uploaded_files)
                if result:
                    output_buffer, output_filename = result
                    st.markdown(
                        f'<div class="success-box">Success! Your merged file is ready: <strong>{output_filename}</strong></div>',
                        unsafe_allow_html=True
                    )
                    st.download_button(
                        label="Download Your Merged Excel!",
                        data=output_buffer,
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="download_button"
                    )

if __name__ == "__main__":
    main()
