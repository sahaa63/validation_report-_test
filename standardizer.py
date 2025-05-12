import streamlit as st
import pandas as pd
import os
import io

def standardize_column_data(df1, df2, common_columns):
    for col in common_columns:
        # Standardize numeric columns
        if pd.api.types.is_numeric_dtype(df1[col]) and pd.api.types.is_numeric_dtype(df2[col]):
            df1[col] = pd.to_numeric(df1[col], errors='coerce')
            df2[col] = pd.to_numeric(df2[col], errors='coerce')

        # Standardize datetime columns
        elif pd.api.types.is_datetime64_any_dtype(df1[col]) or pd.api.types.is_datetime64_any_dtype(df2[col]):
            df1[col] = pd.to_datetime(df1[col], errors='coerce')
            df2[col] = pd.to_datetime(df2[col], errors='coerce')

        # Standardize text columns
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

            # Read all sheets
            xls = pd.ExcelFile(uploaded_file)
            sheet_names = xls.sheet_names

            # Validate presence of 'Excel' and 'PBI' sheets (case-sensitive)
            if 'Excel' in sheet_names and 'PBI' in sheet_names:
                df_excel = pd.read_excel(xls, sheet_name='Excel')
                df_pbi = pd.read_excel(xls, sheet_name='PBI')

                # Find common columns
                common_columns = list(set(df_excel.columns) & set(df_pbi.columns))

                # Standardize common columns
                df_excel_std, df_pbi_std = standardize_column_data(df_excel.copy(), df_pbi.copy(), common_columns)

                # Save to output Excel
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
                st.error("❌ The uploaded file must contain two sheets named 'Excel' and 'PBI' (case-sensitive).")
        except Exception as e:
            st.error(f"❌ Error processing file: {e}")

if __name__ == "__main__":
    main()
