import streamlit as st
import pandas as pd
import io
import numpy as np

def generate_validation_report(cognos_df, pbi_df):
    # Identify dimensions and measures
    dims = [col for col in cognos_df.columns if col in pbi_df.columns and 
            (cognos_df[col].dtype == 'object' or '_id' in col.lower() or '_key' in col.lower() or
             '_ID' in col or '_KEY' in col)]
    cognos_measures = [col for col in cognos_df.columns if col not in dims]
    pbi_measures = [col for col in pbi_df.columns if col not in dims]
    all_measures = list(set(cognos_measures) & set(pbi_measures))  # Only measures present in both

    # Create a unique key by concatenating all dimensions
    cognos_df['unique_key'] = cognos_df[dims].astype(str).agg('-'.join, axis=1).str.upper()  # Capitalize for case-insensitive comparison
    pbi_df['unique_key'] = pbi_df[dims].astype(str).agg('-'.join, axis=1).str.upper()  # Capitalize for case-insensitive comparison

    # Move 'unique_key' to the first column
    cognos_df = cognos_df[['unique_key'] + [col for col in cognos_df.columns if col != 'unique_key']]
    pbi_df = pbi_df[['unique_key'] + [col for col in pbi_df.columns if col != 'unique_key']]

    # Create the validation report dataframe
    validation_report = pd.DataFrame({'unique_key': list(set(cognos_df['unique_key']) | set(pbi_df['unique_key']))})

    # Add dimensions
    for dim in dims:
        validation_report[dim] = validation_report['unique_key'].map(dict(zip(cognos_df['unique_key'], cognos_df[dim])))
        validation_report[dim].fillna(validation_report['unique_key'].map(dict(zip(pbi_df['unique_key'], pbi_df[dim]))), inplace=True)

    # Determine presence in sheets
    validation_report['presence'] = validation_report['unique_key'].apply(
        lambda key: 'Present in Both' if key in cognos_df['unique_key'].values and key in pbi_df['unique_key'].values
        else ('Present in Cognos' if key in cognos_df['unique_key'].values
              else 'Present in PBI')
    )

    # Add measures and calculate differences
    for measure in all_measures:
        validation_report[f'{measure}_Cognos'] = validation_report['unique_key'].map(dict(zip(cognos_df['unique_key'], cognos_df[measure])))
        validation_report[f'{measure}_PBI'] = validation_report['unique_key'].map(dict(zip(pbi_df['unique_key'], pbi_df[measure])))
        
        # Calculate difference (PBI - Cognos)
        #validation_report[f'{measure}_Diff'] = validation_report[f'{measure}_PBI'] - validation_report[f'{measure}_Cognos']
        validation_report[f'{measure}_Diff'] = validation_report[f'{measure}_PBI'].fillna(0) - validation_report[f'{measure}_Cognos'].fillna(0)

    # Reorder columns
    column_order = ['unique_key'] + dims + ['presence'] + \
                   [col for measure in all_measures for col in 
                    [f'{measure}_Cognos', f'{measure}_PBI', f'{measure}_Diff']]
    validation_report = validation_report[column_order]

    return validation_report, cognos_df, pbi_df


def main():
    st.title("Validation Report Generator")

    # Add helper text
    st.markdown("""
    **Important Assumptions:**
    1. Upload the Excel file with two sheets: "Cognos" and "PBI".
    2. Make sure the column names are similar in both sheets.
    3. If there are ID/Key/Code columns, make sure the ID or Key columns contains "_ID" or "_KEY" (case insensitive)
    4. Working with merged reports? unmerge them like this [link](https://www.loom.com/share/c876bb4cf67e45e7b01cd64facb6f7d8?sid=fdd1bb3e-96cf-4eaa-af3e-2a951861a8cc)
   """)

    #st.markdown("Working with merged reports? unmerge them like this [link](https://www.loom.com/share/c876bb4cf67e45e7b01cd64facb6f7d8?sid=fdd1bb3e-96cf-4eaa-af3e-2a951861a8cc)")


    st.markdown("---")  # Add a horizontal line for visual separation

    uploaded_file = st.file_uploader("Upload Excel file", type="xlsx")

    if uploaded_file is not None:
        try:
            xls = pd.ExcelFile(uploaded_file)
            cognos_df = pd.read_excel(xls, 'Cognos')
            pbi_df = pd.read_excel(xls, 'PBI')

            validation_report, cognos_df, pbi_df = generate_validation_report(cognos_df, pbi_df)

            st.subheader("Validation Report Preview")
            st.dataframe(validation_report)

            # Generate Excel file for download
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                cognos_df.to_excel(writer, sheet_name='Cognos', index=False)
                pbi_df.to_excel(writer, sheet_name='PBI', index=False)
                validation_report.to_excel(writer, sheet_name='Validation_Report', index=False)

            output.seek(0)
            
            st.download_button(
                label="Download Excel Report",
                data=output,
                file_name="validation_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    main()
