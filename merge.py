import streamlit as st
import pandas as pd
import io
import os
from openpyxl import load_workbook

def combine_excel_files(file_list):
    if not file_list or len(file_list) > 10:
        st.error("Please upload between 1 and 10 Excel files.")
        return None

    # Extract original filename from the first file (assuming common prefix)
    first_filename = os.path.splitext(file_list[0].name)[0]
    original_filename = first_filename[:-1] if first_filename[-1].isdigit() else first_filename
    output_filename = f"{original_filename}_validation_report.xlsx"

    # Create a new workbook in memory
    output_buffer = io.BytesIO()
    output_wb = load_workbook(filename=output_buffer)

    # Dictionary to track sheet names and avoid duplicates
    sheet_name_count = {}

    # Process each uploaded file
    for uploaded_file in file_list:
        # Load the workbook from the uploaded file
        file_bytes = uploaded_file.read()
        wb = load_workbook(filename=io.BytesIO(file_bytes))

        # Copy all sheets from the current file
        for sheet_name in wb.sheetnames:
            # Handle duplicate sheet names by appending a number
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

    # Remove the default sheet created by openpyxl if it exists
    if 'Sheet' in output_wb.sheetnames:
        output_wb.remove(output_wb['Sheet'])

    # Save the combined workbook to the buffer
    output_wb.save(output_buffer)
    output_buffer.seek(0)

    return output_buffer, output_filename

def main():
    st.title("Excel File Merger")

    st.markdown("""
    **Instructions:**
    - Upload up to 10 Excel files to merge.
    - All sheets from each file will be combined into a single output file.
    - If sheet names conflict, they will be renamed with a numeric suffix (e.g., 'Sheet_1', 'Sheet_2').
    - The output file will be named based on the first uploaded file's name (e.g., 'Report_validation_report.xlsx').
    """)

    # File uploader for multiple files (max 10)
    uploaded_files = st.file_uploader(
        "Upload Excel Files",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        help="Select up to 10 Excel files to merge."
    )

    if uploaded_files:
        if len(uploaded_files) > 10:
            st.error("Maximum 10 files allowed. Please upload fewer files.")
        else:
            st.write(f"Uploaded {len(uploaded_files)} file(s):")
            for file in uploaded_files:
                st.write(f"- {file.name}")

            # Combine the files
            result = combine_excel_files(uploaded_files)
            if result:
                output_buffer, output_filename = result

                # Provide download button
                st.download_button(
                    label="Download Merged Excel File",
                    data=output_buffer,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

if __name__ == "__main__":
    main()
