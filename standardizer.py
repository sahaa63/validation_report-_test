import streamlit as st
import pandas as pd
import os
import io

# 🧠 Function to generate the validation report
def generate_validation_report(excel_df, pbi_df):
    try:
        excel_grouped = excel_df.groupby(['Date', 'DC', 'Material'], dropna=False).agg('sum').reset_index()
        pbi_grouped = pbi_df.groupby(['Date', 'DC', 'Material'], dropna=False).agg('sum').reset_index()

        merged_df = pd.merge(excel_grouped, pbi_grouped, on=['Date', 'DC', 'Material'], suffixes=('_excel', '_pbi'), how='outer')

        validation_results = merged_df.copy()
        for col in excel_grouped.columns:
            if col not in ['Date', 'DC', 'Material']:
                excel_col = f"{col}_excel"
                pbi_col = f"{col}_pbi"
                if excel_col in merged_df.columns and pbi_col in merged_df.columns:
                    validation_results[f"{col}_Validation"] = merged_df.apply(
                        lambda row: "Matching" if pd.isna(row[excel_col]) and pd.isna(row[pbi_col]) 
                        else "Matching" if row[excel_col] == row[pbi_col] 
                        else "Not Matching", axis=1)

        return validation_results, excel_grouped, pbi_grouped
    except Exception as e:
        st.error(f"Error during validation report generation: {e}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

# 📊 Function to generate diff checker
def generate_diff_checker(validation_df):
    diff_df = pd.DataFrame(columns=['Column', 'Matching', 'Not Matching'])
    for col in validation_df.columns:
        if col.endswith('_Validation'):
            matching = (validation_df[col] == 'Matching').sum()
            not_matching = (validation_df[col] == 'Not Matching').sum()
            diff_df = pd.concat([diff_df, pd.DataFrame([{
                'Column': col.replace('_Validation', ''),
                'Matching': matching,
                'Not Matching': not_matching
            }])], ignore_index=True)
    return diff_df

# 🔍 Function to check for missing columns
def column_checklist(excel_df, pbi_df):
    excel_columns = set(excel_df.columns)
    pbi_columns = set(pbi_df.columns)
    missing_in_excel = pbi_columns - excel_columns
    missing_in_pbi = excel_columns - pbi_columns

    checklist = []
    for col in missing_in_excel:
        checklist.append({'Column': col, 'Missing In': 'Excel'})
    for col in missing_in_pbi:
        checklist.append({'Column': col, 'Missing In': 'PBI'})

    return pd.DataFrame(checklist)

# 🚀 Main Streamlit App
def main():
    st.set_page_config(page_title="Standardiser", layout="wide")
    st.title("📊 Standardiser: Excel vs Power BI Validation App")

    uploaded_file = st.file_uploader("Upload your Excel file with 'Excel' and 'PBI' sheets", type=["xlsx"])

    if uploaded_file:
        file_name = os.path.splitext(uploaded_file.name)[0]
        std_file_name = f"{file_name}_std.xlsx"

        try:
            xls = pd.ExcelFile(uploaded_file)
            sheet_names = [sheet.lower() for sheet in xls.sheet_names]

            if 'excel' in sheet_names and 'pbi' in sheet_names:
                excel_df = pd.read_excel(xls, sheet_name=xls.sheet_names[sheet_names.index('excel')])
                pbi_df = pd.read_excel(xls, sheet_name=xls.sheet_names[sheet_names.index('pbi')])

                validation_report, excel_agg, pbi_agg = generate_validation_report(excel_df, pbi_df)
                diff_checker = generate_diff_checker(validation_report)
                column_check_df = column_checklist(excel_df, pbi_df)

                st.success("✅ Validation Report Generated Successfully!")

                with st.expander("📄 Preview: Validation Report"):
                    st.dataframe(validation_report.head(100))

                with st.expander("📌 Summary: Matching vs Not Matching"):
                    st.dataframe(diff_checker)

                with st.expander("🧾 Column Checklist"):
                    st.dataframe(column_check_df)

                # Save output Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    validation_report.to_excel(writer, sheet_name='Validation Report', index=False)
                    diff_checker.to_excel(writer, sheet_name='Summary', index=False)
                    column_check_df.to_excel(writer, sheet_name='Column Check', index=False)

                st.download_button(
                    label=f"📥 Download {std_file_name}",
                    data=output.getvalue(),
                    file_name=std_file_name,
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
            else:
                st.error("❌ Please ensure the Excel file has both 'Excel' and 'PBI' sheets (case-insensitive).")

        except Exception as e:
            st.error(f"Something went wrong: {e}")

if __name__ == "__main__":
    main()
