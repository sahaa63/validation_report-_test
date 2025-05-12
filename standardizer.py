import streamlit as st
import pandas as pd
import os
from io import BytesIO

st.title("Standardiser App")

# Upload Excel file
uploaded_file = st.file_uploader("Upload an Excel file with sheets 'excel' and 'PBI'", type=["xlsx"])

if uploaded_file:
    try:
        # Read both sheets
        xl = pd.ExcelFile(uploaded_file)
        df_excel = xl.parse('excel')
        df_pbi = xl.parse('PBI')

        # Get common columns
        common_columns = [col for col in df_excel.columns if col in df_pbi.columns]

        # Function to standardize column formats
        def standardize_column_data(df1, df2, common_columns):
            for col in common_columns:
                # Handle numeric
                if pd.api.types.is_numeric_dtype(df1[col]) and pd.api.types.is_numeric_dtype(df2[col]):
                    df1[col] = pd.to_numeric(df1[col], errors='coerce')
                    df2[col] = pd.to_numeric(df2[col], errors='coerce')
                
                # Handle dates
                elif pd.api.types.is_datetime64_any_dtype(df1[col]) or pd.api.types.is_datetime64_any_dtype(df2[col]):
                    df1[col] = pd.to_datetime(df1[col], errors='coerce').dt.date
                    df2[col] = pd.to_datetime(df2[col], errors='coerce').dt.date

                # Handle everything else as string
                else:
                    df1[col] = df1[col].astype(str).str.strip().str.lower()
                    df2[col] = df2[col].astype(str).str.strip().str.lower()
            return df1, df2

        # Apply standardization
        df_excel_std, df_pbi_std = standardize_column_data(df_excel.copy(), df_pbi.copy(), common_columns)

        # Create in-memory Excel file
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_excel_std.to_excel(writer, sheet_name='excel', index=False)
            df_pbi_std.to_excel(writer, sheet_name='PBI', index=False)

        output.seek(0)

        # Construct output filename
        original_name = os.path.splitext(uploaded_file.name)[0]
        output_filename = f"{original_name}_std.xlsx"

        # Download link
        st.success("Standardization complete. Download the standardized file below:")
        st.download_button(label="📥 Download Standardized Excel",
                           data=output,
                           file_name=output_filename,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except ValueError as e:
        st.error(f"Error: {e}")
    except Exception as e:
        st.exception(f"Unexpected error: {e}")
