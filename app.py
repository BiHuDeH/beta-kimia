import streamlit as st
import pandas as pd

# Set up Streamlit app title and instructions
st.title("Material Calculation App")
st.write("Upload an Excel file to calculate 'گرماژ' based on 'فروش' and 'تعداد'.")

# File upload
uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

if uploaded_file:
    # Load the uploaded Excel file and select the specific sheet
    sheet_name = 'بهای تمام شده کالای ساخته شده'
    df = pd.read_excel(uploaded_file, sheet_name=sheet_name)

    # Define column names for 'فروش' and 'تعداد' based on the sheet's structure
    sales_column = 'فروش'    # Adjust to the exact name in your file
    quantity_column = 'تعداد'  # Adjust to the exact name in your file

    # Check if both columns exist in the DataFrame
    if sales_column in df.columns and quantity_column in df.columns:
        # Calculate "گرماژ" by multiplying 'فروش' and 'تعداد' columns
        df['گرماژ'] = df[sales_column] * df[quantity_column]

        # Add a summary row with the sum of the "گرماژ" column
        summary_row = pd.DataFrame({sales_column: ['Total'], 'گرماژ': [df['گرماژ'].sum()]})
        df = pd.concat([df, summary_row], ignore_index=True)

        # Display the updated DataFrame in the Streamlit app
        st.write("Modified Data:")
        st.dataframe(df)

        # Option to download the modified file
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