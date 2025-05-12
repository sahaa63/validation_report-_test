import streamlit as st
import pandas as pd
import os
import io

def standardize_column_data(df1, df2, common_columns):
    for col in common_columns:
        # Handle numeric columns
        if pd.api.types.is_numeric_dtype(df1[col]) and pd.api.types.is_numeric_dtype(df2[col]):
            df1[col] = pd.to_numeric(df1[col], errors='coerce')
            df2[col] = pd.to_numeric(df2[col], errors='coerce')

        # Handle datetime columns
        elif pd.api.types.is_datetime64_any_dtype(df1[col]) or pd.api.types.is_datetime64_any_dtype(df2[col]):
            df1[col] = pd.to_datetime(df1[col], errors='coerce')
            df2[col] = pd.to_datetime(df2[col], errors='coerce')

        # Handle string/object columns
        else:
            df1[col] = df1[col].astype(str).str.strip().str.lower()
            df2[col] = df2[col].astype(str).str.strip().str.lower()
    return df1, df2

def main():
    st.set_page_config(page_title="Standardiser", layout="centered")
    st.title("🧮 Standardiser App")

    uploaded_file = st.file_uploader("Upload Excel file with 'Excel' and 'PBI' sheets", type=["xlsx"])

    if uploaded_file:
        try:
            file_name = os.path.splitext(uploaded_file.name)[0]
            output_filename = f"{file_name}_std.xlsx"

            # Load the Excel file
            xls = pd.ExcelFile(uploaded_file)
            sheet_names = [sheet.lower() for sheet in xls.sheet_names]

            if 'excel' in sheet_names and 'pbi' in sheet_names:
                df_excel = pd.read_excel(xls, sheet_name=xls.sheet_names[sheet_names.index('excel')])
                df_pbi = pd.read_excel(xls, sheet_name=xls.sheet_names[sheet_names.index('pbi')])

                # Find common columns
                common_columns = list(set(df_excel.columns) & set(df_pbi.columns))

                # Standardize data for common columns
                df_excel_std, df_pbi_std = standardize_column_data(df_excel.copy(), df_pbi.copy(), common_columns)

                # Write standardized data back to Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_excel_std.to_excel(writer, sheet_name='Excel', index=False)
                    df_pbi_std.to_excel(writer, sheet_name='PBI', index=False)

                st.success("✅ File standardized successfully.")
                st.download_button(
                    label=f"📥 Download Standardized File: {output_filename}",
                    data=output.getvalue(),
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("❌ Both 'Excel' and 'PBI' sheets must exist (case-insensitive).")
        except Exception as e:
            st.error(f"❌ Error processing file: {e}")

if __name__ == "__main__":
    main()
