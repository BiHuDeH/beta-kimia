import streamlit as st
import pandas as pd

st.title("Material Calculation App")
st.write("Upload an Excel file to calculate 'گرماژ' based on 'فروش' and 'تعداد'.")

# File upload
uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

if uploaded_file:
    # Print all sheet names to verify the correct name
    excel_file = pd.ExcelFile(uploaded_file)
    st.write("Available sheet names:", excel_file.sheet_names)

    # Specify the exact sheet name here once confirmed
    sheet_name = 'بهای تمام شده کالای ساخته شده'  # Adjust if necessary after checking the names
    df = pd.read_excel(uploaded_file, sheet_name=sheet_name)

    # Define columns for "فروش" and "تعداد" based on the exact column names
    sales_column = 'فروش'    # Replace with exact name
    quantity_column = 'تعداد'  # Replace with exact name

    # Check if columns exist
    if sales_column in df.columns and quantity_column in df.columns:
        df['گرماژ'] = df[sales_column] * df[quantity_column]
        summary_row = pd.DataFrame({sales_column: ['Total'], 'گرماژ': [df['گرماژ'].sum()]})
        df = pd.concat([df, summary_row], ignore_index=True)
        
        st.write("Modified Data:")
        st.dataframe(df)

        output_file = 'Modified_Material_File.xlsx'
        df.to_excel(output_file, index=False, sheet_name=sheet_name)
        with open(output_file, "rb") as file:
            btn = st.download_button(
                label="Download Modified Excel File",
                data=file,
                file_name=output_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.error(f"Columns '{sales_column}' or '{quantity_column}' not found in the sheet.")