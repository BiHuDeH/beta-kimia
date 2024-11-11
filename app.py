import streamlit as st
import pandas as pd

st.title("Material Calculation App")
st.write("Upload an Excel file to calculate 'گرماژ' based on 'فروش' and 'تعداد'.")

# File upload
uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

if uploaded_file:
    # Load the uploaded Excel file
    excel_file = pd.ExcelFile(uploaded_file)
    
    # Display all sheet names to confirm the correct one
    st.write("Available sheet names in the uploaded file:", excel_file.sheet_names)
    
    # Try to specify the exact sheet name after confirming
    sheet_name = 'بهای تمام شده کالای ساخته شده'  # Modify this once you verify the exact name
    
    # Check if the desired sheet name is actually in the file
    if sheet_name in excel_file.sheet_names:
        # Load the specified sheet
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
        
        # Define columns for "فروش" and "تعداد" based on the exact column names in the file
        sales_column = 'فروش'    # Adjust if needed based on column names
        quantity_column = 'تعداد'  # Adjust if needed based on column names

        # Verify that both columns exist in the sheet
        if sales_column in df.columns and quantity_column in df.columns:
            # Calculate 'گرماژ'
            df['گرماژ'] = df[sales_column] * df[quantity_column]
            
            # Add summary row
            summary_row = pd.DataFrame({sales_column: ['Total'], 'گرماژ': [df['گرماژ'].sum()]})
            df = pd.concat([df, summary_row], ignore_index=True)
            
            # Display the modified data
            st.write("Modified Data:")
            st.dataframe(df)

            # Allow file download
            output_file = 'Modified_Material_File.xlsx'
            df.to_excel(output_file, index=False, sheet_name=sheet_name)
            with open(output_file, "rb") as file:
                st.download_button(
                    label="Download Modified Excel File",
                    data=file,
                    file_name=output_file,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.error(f"Columns '{sales_column}' or '{quantity_column}' not found in the sheet.")
    else:
        st.error(f"Sheet '{sheet_name}' not found in the uploaded file.")