import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

# Function to transform the "VPO No" column
def transform_vpo_no(vpo_no):
    if isinstance(vpo_no, str):
        if vpo_no.startswith('8'):
            return vpo_no[:8]
        elif vpo_no.startswith('D'):
            return 'P' + vpo_no[1:-3]
    return vpo_no

# Function to convert 'PCD' column from float to datetime
def convert_to_date(x):
    try:
        if x and x != '':
            x = str(int(float(x)))
            return pd.to_datetime(x, format='%Y%m%d', errors='coerce')
        return pd.NaT
    except:
        return pd.NaT

# Function to process the uploaded Excel file
def process_excel(file):
    data = pd.read_excel(file, sheet_name='Sheet1')

    # Drop unnecessary columns
    columns_to_drop = ['CBU', 'Buyer','Buyer Division Code', 'Cust Style No', 'Product Group',
                       'Style Category', 'Garment Fabrication', 'Lead Factory', 'Prod Warehouse',
                       'Max CO Sts', 'Delivery Term', 'Color Code', 'FOB Date', 'Max Departure Date - CO',
                       'Cum Wash Rev Qty', 'Cum Wash Rev Rej Qty', 'Remaining Qty', 'Allocated Qty',
                       'Invoiced Qty', 'FOB Price', 'FOB after discount', 'SMV', 'Planned SAH',
                       'Costing Efficiency', 'CO Responsible', 'CO Create Min Date', 'CO Create Max Date',
                       'Drop Dead Date', 'AOQ', 'Type', 'Projection Ref']
    data_cleaned = data.drop(columns=columns_to_drop)

    # Filter data
    data_cleaned = data_cleaned[data_cleaned['Group Tech Class'] == "BELUNIQLO"]

    # Apply the transformation function to the "VPO No" column
    data_cleaned['PO'] = data_cleaned['VPO No'].apply(transform_vpo_no)

    # Update 'Production Plan ID' column
    data_cleaned['Production Plan ID'] = np.where(
        data_cleaned['Production Plan ID'].isna() & data_cleaned['PO'].str.startswith('8'),
        data_cleaned['PO'],
        np.where(
            data_cleaned['Production Plan ID'].isna() & data_cleaned['Season'].str[-2:] == '23',
            'Season-23',
            data_cleaned['Production Plan ID']
        )
    )

    # Convert 'PCD' column to datetime
    data_cleaned['PCD'] = data_cleaned['PCD'].apply(convert_to_date)

    return data_cleaned

# Function to process the uploaded CSV file
def process_csv(file):
    new_csv_data = pd.read_csv(file)

    # List of columns to keep
    columns_to_keep = ['Production Plan ID', 'Main Sample Code', 'Season', 'Year', 
                       'Production Plan Type Name', 'EXF', 'Contracted Date', 
                       'Requested Wh Date', 'Business Unit', 'PO Order NO']
    new_csv_data_cleaned = new_csv_data[columns_to_keep].copy()

    # Rename the column 'PO Order NO' to 'PO'
    new_csv_data_cleaned.rename(columns={'PO Order NO': 'PO'}, inplace=True)

    # Convert the date columns
    date_columns = ['EXF', 'Contracted Date', 'Requested Wh Date']
    new_csv_data_cleaned[date_columns] = new_csv_data_cleaned[date_columns].apply(pd.to_datetime, format='%m/%d/%Y', errors='coerce')

    return new_csv_data_cleaned

# Function to process the third Excel file (RFID Gihan)
def process_rfid_excel(file):
    data = pd.read_excel(file, sheet_name='sheet1')

    # Forward fill 'DO No./Product No.' column
    data['DO No./Product No.'] = data['DO No./Product No.'].ffill()

    # Apply the condition to 'Set Detail'
    data['Set Detail'] = np.where(
        data['Set Detail'].isna() | (data['Set Detail'] == '-'),
        data['Set Code'],
        data['Set Detail']
    )

    # Extract relevant columns
    columns_to_keep = ['DO No./Product No.', 'Set Code', 'Set Detail',
                       data.columns[4], data.columns[6], data.columns[8], data.columns[10]]
    data_cleaned = data.iloc[1:, data.columns.get_indexer(columns_to_keep)]

    data_cleaned.columns = ['DO No./Product No.', 'Set Code', 'Set Detail', 
                            'Order Quantity', 'Packing Quantity', 
                            'Loading Quantity', 'Inventory Quantity']

    # Convert specified fields to integers
    data_cleaned[['Order Quantity', 'Packing Quantity', 'Loading Quantity', 'Inventory Quantity']] = data_cleaned[
        ['Order Quantity', 'Packing Quantity', 'Loading Quantity', 'Inventory Quantity']].astype(int)

    # Create a new field 'Color Code' and 'Pack Method'
    data_cleaned['Color Code'] = data_cleaned['Set Code'].str[:2].astype(str)
    data_cleaned['Pack Method'] = np.where(
        data_cleaned['Set Code'].str[:2].str.isalpha(), 'AST', 
        np.where(data_cleaned['Set Code'].str[:2].str.isdigit(), '1SL', '')
    )

    return data_cleaned

# Function to merge and update dataframes
def merge_and_update_data(ob_clean_df, spl_clean_df, rfid_clean_df):
    ob_clean_df['Production Plan ID'] = ob_clean_df['Production Plan ID'].fillna(
        ob_clean_df['PO'].map(spl_clean_df.set_index('PO')['Production Plan ID'])
    ).fillna('N/A')

    # Merge updated OB_clean DataFrame with SPL_clean DataFrame
    merged_df = pd.merge(ob_clean_df, spl_clean_df, on='Production Plan ID', how='left')

    # Perform final calculations and add columns
    merged_df['Cut%'] = ((merged_df['Cum Cut Qty'] / merged_df['CO Qty']) * 100).round(2)
    merged_df['Sewin%'] = ((merged_df['Cum Sew In Qty'] / merged_df['CO Qty']) * 100).round(2)
    merged_df['Sewin Rej%'] = ((merged_df['Cum Sew In Rej Qty'] / merged_df['Cum Sew In Qty']) * 100).round(2)
    merged_df['Sewout%'] = ((merged_df['Cum SewOut Qty'] / merged_df['CO Qty']) * 100).round(2)
    merged_df['Sewout Rej%'] = ((merged_df['Cum Sew Out Rej Qty'] / merged_df['Cum SewOut Qty']) * 100).round(2)
    merged_df['CTN%'] = ((merged_df['Cum CTN Qty'] / merged_df['CO Qty']) * 100).round(2)
    merged_df['Del%'] = ((merged_df['Delivered Qty'] / merged_df['CO Qty']) * 100).round(2)

    merged_df['Delays'] = (merged_df['Requested Wh Date'] - merged_df['POWH-PLN']).dt.days
    merged_df['Delay/Early'] = np.where(merged_df['Delays'] > 0, "Delay", "No Delay")

    # Add 'Color Code' and merge with RFID data
    merged_df['Color Code'] = merged_df['Color Name'].str[:2]
    rfid_grouped = rfid_clean_df.groupby(['DO No./Product No.', 'Color Code', 'Pack Method'])['Packing Quantity'].sum().reset_index()
    rfid_grouped.rename(columns={'Packing Quantity': 'RFID'}, inplace=True)

    final_merged_data = pd.merge(merged_df, rfid_grouped, how='left', left_on=['VPO No', 'Color Code', 'Pack Method'], right_on=['DO No./Product No.', 'Color Code', 'Pack Method'])
    final_merged_data.drop(columns=['DO No./Product No.'], inplace=True)

    final_merged_data['RFID'] = pd.to_numeric(final_merged_data['RFID'], errors='coerce').fillna(0)
    final_merged_data['RFID%'] = (final_merged_data['RFID'] / final_merged_data['CO Qty']).fillna(0) * 100
    final_merged_data['RFID%'] = final_merged_data['RFID%'].round(2)

    final_merged_data['Del_Dummy%'] = final_merged_data['Del%'].str.rstrip('%').astype(float).fillna(0)
    final_merged_data['Min CO Sts'] = pd.to_numeric(final_merged_data['Min CO Sts'], errors='coerce').fillna(0)

    final_merged_data['Status'] = np.select(
        [
            final_merged_data['Del_Dummy%'] >= 100.0,
            final_merged_data['Del_Dummy%'] <= 0.0,
            (final_merged_data['Del_Dummy%'] > 0.0) & (final_merged_data['Del_Dummy%'] < 100.0) & (final_merged_data['Min CO Sts'] < 66),
            (final_merged_data['Del_Dummy%'] > 0.0) & (final_merged_data['Del_Dummy%'] < 100.0) & (final_merged_data['Min CO Sts'] >= 66)
        ],
        ['Shipped', 'Pending', 'Short Ship', 'Short Close'],
        default=''
    )

    final_merged_data = final_merged_data.fillna('')
    return final_merged_data

# Streamlit app
def main():
    st.set_page_config(page_title="OB Macro", layout="wide")
    st.sidebar.title("OB Macro")
    st.sidebar.write("Upload the required files for processing.")
    
    # Upload files
    uploaded_excel_1 = st.sidebar.file_uploader("Choose the first Excel file", type="xlsx")
    uploaded_csv = st.sidebar.file_uploader("Choose a CSV file", type="csv")
    uploaded_excel_2 = st.sidebar.file_uploader("Choose the second Excel file (RFID Gihan)", type="xlsx")

    st.markdown("<h2 style='text-align: center; color: #4CAF50;'>OB Macro Processing Tool</h2>", unsafe_allow_html=True)
    st.write("This tool processes multiple files, merges them, and applies updates.")

    if uploaded_excel_1 and uploaded_csv and uploaded_excel_2:
        ob_clean_df = process_excel(uploaded_excel_1)
        spl_clean_df = process_csv(uploaded_csv)
        rfid_clean_df = process_rfid_excel(uploaded_excel_2)

        # Merge and update data
        final_merged_data = merge_and_update_data(ob_clean_df, spl_clean_df, rfid_clean_df)

        # Reorder columns (example)
        final_order = ['Main Sample Code', 'Style No', 'Season_y', 'Year', 'Production Plan Type Name', 
                      'Production Plan ID', 'VPO No', 'Item Description', 'Destination', 
                      'Business Unit', 'EXF', 'Contracted Date', 'Requested Wh Date', 
                      'Color Name', 'PED', 'Shipment Mode', 'MODE', 'EXF-PLN', 'ETD-PLN', 
                      'POWH-PLN', 'Delays', 'Delay/Early', 'Factory – Remarks', 'CO Qty', 
                      'Cum Cut Qty', 'Cut%', 'Cum Sew In Qty', 'Sewin%', 'Cum Sew In Rej Qty', 
                      'Sewin Rej%', 'Cum SewOut Qty', 'Sewout%', 'Cum Sew Out Rej Qty', 
                      'Sewout Rej%', 'Cum CTN Qty', 'CTN%', 'Cum CTN Rej Qty', 'RFID', 
                      'RFID%', 'Delivered Qty', 'Del%', 'Excess/Short Shipped Qty', 'PCD', 
                      'Status']

        final_merged_data = final_merged_data[final_order]

        # Save to an in-memory Excel file
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            final_merged_data.to_excel(writer, index=False, sheet_name='Final Report')
            writer.save()

        st.download_button(
            label="Download Finalized Report",
            data=output.getvalue(),
            file_name="finalized_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()
