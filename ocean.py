import streamlit as st
import pandas as pd
from io import BytesIO

# Function to process the Excel files and map columns
def process_files(client_file, template_file):
    df = pd.read_excel(client_file)
    template_df = pd.read_excel(template_file)

    # Column mapping from df to the template_df
    column_mapping = {
        'Res_Date': 'Job Close Date',
        'Facility': 'Consolidation Type',
        'Departure': 'POL',
        'Start Date': 'ATD',
        "End Date": "ATA",
        "Arrival": "POD",
        "Weight Ton": "Weight(Tons)"
    }

    # Preserve the first row (header) of the template
    preserved_header = template_df.iloc[:0, :]

    # Create a DataFrame with the same columns as the template
    matched_data = pd.DataFrame(columns=template_df.columns)

    # Map and populate data from df to matched_data based on column_mapping
    for template_col, client_col in column_mapping.items():
        if client_col in df.columns:
            matched_data[template_col] = df[client_col]
        else:
            st.write(f"Column '{client_col}' not found in client file")

    # Combine the preserved header with the matched data
    matched_data = pd.concat([preserved_header, matched_data], ignore_index=True)

    # Add extra columns to matched_data
    matched_data['CF Standard'] = "IATA"
    matched_data['Gas'] = "CO2"
    matched_data['Activity Unit'] = "Kg"

    # Drop any rows with missing values
    matched_data.dropna(inplace=True)

    return matched_data

# Function to convert DataFrame to Excel file in-memory
def to_excel_bytes(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    processed_data = output.getvalue()
    return processed_data

# Streamlit app layout
st.title('Excel Processing App')

# File uploader for the client file and template file
client_file = st.file_uploader("Upload Client Workbook", type="xls")
template_file = 'Freight-Sample_scope3.xlsx'  # Directly loading template from file system

if client_file is not None:
    # Process files and map columns
    processed_data = process_files(client_file, template_file)
    
    # Show a sample of the processed data
    st.write("Processed Data Preview:")
    st.dataframe(processed_data.head())

    # Convert DataFrame to Excel and make it downloadable
    excel_data = to_excel_bytes(processed_data)

    # Button to download the processed data
    st.download_button(
        label="Download Processed Data",
        data=excel_data,
        file_name="processed_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
