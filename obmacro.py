import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# Function to transform the "VPO No" column
def transform_vpo_no(vpo_no):
    if isinstance(vpo_no, str):
        if vpo_no.startswith('8'):
            return vpo_no[:8]
        elif vpo_no.startswith('D'):
            return 'P' + vpo_no[1:-3]
    return vpo_no

# Function to convert 'PCD' column to datetime
def convert_to_date(x):
    if pd.notnull(x):
        x = str(int(float(x)))
        return pd.to_datetime(x, format='%Y%m%d', errors='coerce')
    return pd.NaT

# Process the uploaded Excel file
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

    # Filter data and transform columns
    data_cleaned = data_cleaned[data_cleaned['Group Tech Class']=="BELUNIQLO"]
    data_cleaned['PO'] = data_cleaned['VPO No'].apply(transform_vpo_no)
    data_cleaned['Production Plan ID'] = np.where(
        data_cleaned['Production Plan ID'].isna(),
        np.where(data_cleaned['PO'].str.startswith('8'), data_cleaned['PO'], "Season-23"),
        data_cleaned['Production Plan ID']
    )
    data_cleaned['PCD'] = data_cleaned['PCD'].apply(convert_to_date)

    return data_cleaned

# Function to process the uploaded CSV file
def process_csv(file):
    new_csv_data = pd.read_csv(file)

    columns_to_keep = ['Production Plan ID', 'Main Sample Code', 'Season', 'Year', 
                       'Production Plan Type Name', 'EXF', 'Contracted Date', 
                       'Requested Wh Date', 'Business Unit', 'PO Order NO']
    new_csv_data_cleaned = new_csv_data.loc[:, columns_to_keep].copy()
    new_csv_data_cleaned.rename(columns={'PO Order NO': 'PO'}, inplace=True)

    date_columns = ['EXF', 'Contracted Date', 'Requested Wh Date']
    new_csv_data_cleaned[date_columns] = new_csv_data_cleaned[date_columns].apply(pd.to_datetime, format='%m/%d/%Y', errors='coerce')

    return new_csv_data_cleaned

# Function to process the third Excel file (RFID Gihan)
def process_rfid_excel(file):
    data = pd.read_excel(file, sheet_name='sheet1')
    data['DO No./Product No.'] = data['DO No./Product No.'].ffill()

    data['Set Detail'] = np.where(
        (data['Set Detail'].isna()) | (data['Set Detail'] == '-'),
        data['Set Code'],
        data['Set Detail']
    )
    data_cleaned = data.iloc[1:].copy()  # Skip the first row

    columns_to_keep = ['DO No./Product No.', 'Set Code', 'Set Detail',
                       data.columns[4], data.columns[6], data.columns[8], data.columns[10]]
    data_cleaned = data_cleaned[columns_to_keep]

    data_cleaned.columns = ['DO No./Product No.', 'Set Code', 'Set Detail', 
                            'Order Quantity', 'Packing Quantity', 
                            'Loading Quantity', 'Inventory Quantity']
    data_cleaned[['Order Quantity', 'Packing Quantity', 'Loading Quantity', 'Inventory Quantity']] = data_cleaned[
        ['Order Quantity', 'Packing Quantity', 'Loading Quantity', 'Inventory Quantity']].astype(int)

    data_cleaned['Color Code'] = data_cleaned['Set Code'].str[:2].astype(str)
    data_cleaned['Pack Method'] = data_cleaned['Set Code'].fillna('').apply(
        lambda x: 'AST' if x[:2].isalpha() else ('1SL' if x[:2].isdigit() else ' ')
    )

    return data_cleaned

# Update 'Production Plan ID'
def update_production_plan_id(ob_clean_df, spl_clean_df):
    merged = ob_clean_df.merge(spl_clean_df[['PO', 'Production Plan ID']], on='PO', how='left', suffixes=('', '_spl'))
    merged['Production Plan ID'] = np.where(
        merged['Production Plan ID'].isna(), merged['Production Plan ID_spl'], merged['Production Plan ID']
    )
    return merged.drop(columns=['Production Plan ID_spl'])

# Merge dataframes
def merge_dataframes(ob_clean_final_df, spl_clean_df):
    return ob_clean_final_df.merge(spl_clean_df, on='Production Plan ID', how='left')

# Final calculations and add columns
def perform_final_calculations(merged_df_corrected):
    merged_df_corrected['Cut%'] = ((merged_df_corrected['Cum Cut Qty'] / merged_df_corrected['CO Qty']) * 100).round(2)
    merged_df_corrected['Sewin%'] = ((merged_df_corrected['Cum Sew In Qty'] / merged_df_corrected['CO Qty']) * 100).round(2)
    merged_df_corrected['Sewin Rej%'] = ((merged_df_corrected['Cum Sew In Rej Qty'] / merged_df_corrected['Cum Sew In Qty']) * 100).round(2)
    merged_df_corrected['Sewout%'] = ((merged_df_corrected['Cum SewOut Qty'] / merged_df_corrected['CO Qty']) * 100).round(2)
    merged_df_corrected['Sewout Rej%'] = ((merged_df_corrected['Cum Sew Out Rej Qty'] / merged_df_corrected['Cum SewOut Qty']) * 100).round(2)
    merged_df_corrected['CTN%'] = ((merged_df_corrected['Cum CTN Qty'] / merged_df_corrected['CO Qty']) * 100).round(2)
    merged_df_corrected['Del%'] = ((merged_df_corrected['Delivered Qty'] / merged_df_corrected['CO Qty']) * 100).round(2)

    if 'Requested Wh Date' in merged_df_corrected.columns and 'POWH-PLN' in merged_df_corrected.columns:
        merged_df_corrected['Delays'] = (merged_df_corrected['Requested Wh Date'] - merged_df_corrected['POWH-PLN']).dt.days
    merged_df_corrected['Delay/Early'] = np.where(merged_df_corrected['Delays'] > 0, "Delay", "No Delay")

    return merged_df_corrected[(merged_df_corrected['Production Plan ID'] != '0') & (merged_df_corrected['Production Plan ID'] != 'Season-23')]

# Final merge with RFID and add Status column
def final_merge_and_status(merged_data, rfid_data):
    rfid_grouped = rfid_data.groupby(['DO No./Product No.', 'Color Code', 'Pack Method'])['Packing Quantity'].sum().reset_index()
    rfid_grouped.rename(columns={'Packing Quantity': 'RFID'}, inplace=True)

    merged_final_data = merged_data.merge(rfid_grouped, how='left', left_on=['VPO No', 'Color Code', 'Pack Method'], right_on=['DO No./Product No.', 'Color Code', 'Pack Method'])
    merged_final_data.drop(columns=['DO No./Product No.'], inplace=True)

    merged_final_data['RFID'] = pd.to_numeric(merged_final_data['RFID'], errors='coerce').fillna(0)
    merged_final_data['CO Qty'] = pd.to_numeric(merged_final_data['CO Qty'], errors='coerce').fillna(0)

    merged_final_data['RFID%'] = (merged_final_data['RFID'] / merged_final_data['CO Qty']).fillna(0) * 100
    merged_final_data['RFID%'] = merged_final_data['RFID%'].round(2)

    merged_final_data['Del_Dummy%'] = merged_final_data['Del%'].str.rstrip('%').astype(float).fillna(0)
    merged_final_data['Min CO Sts'] = pd.to_numeric(merged_final_data['Min CO Sts'], errors='coerce').fillna(0)

    merged_final_data['Status'] = np.select(
        [
            merged_final_data['Del_Dummy%'] >= 100.0,
            merged_final_data['Del_Dummy%'] <= 0.0,
            (merged_final_data['Del_Dummy%'] > 0.0) & (merged_final_data['Del_Dummy%'] < 100.0) & (merged_final_data['Min CO Sts'] < 66),
            (merged_final_data['Del_Dummy%'] > 0.0) & (merged_final_data['Del_Dummy%'] < 100.0) & (merged_final_data['Min CO Sts'] >= 66)
        ],
        ['Shipped', 'Pending', 'Short Ship', 'Short Close'],
        default=''
    )

    return merged_final_data.fillna('')

# Streamlit app
def main():
    st.set_page_config(page_title="OB Macro", layout="wide")
    st.sidebar.title("OB Macro")
    st.sidebar.write("Upload the required files for processing.")
    
    uploaded_excel_1 = st.sidebar.file_uploader("Choose the first Excel file", type="xlsx")
    uploaded_csv = st.sidebar.file_uploader("Choose a CSV file", type="csv")
    uploaded_excel_2 = st.sidebar.file_uploader("Choose the second Excel file (RFID Gihan)", type="xlsx")
    uploaded_delivery_status = st.sidebar.file_uploader("Choose the Delivery Status Excel file", type="xlsx")

    st.markdown("<h2 style='text-align: center; color: #4CAF50;'>OB Macro Processing Tool</h2>", unsafe_allow_html=True)
    st.write("This tool processes multiple files, merges them, and applies updates and conditional formatting.")
    
    if uploaded_excel_1 and uploaded_csv and uploaded_excel_2 and uploaded_delivery_status:
        ob_clean_df = process_excel(uploaded_excel_1)
        spl_clean_df = process_csv(uploaded_csv)
        rfid_clean_df = process_rfid_excel(uploaded_excel_2)

        # Update and merge dataframes
        updated_df = update_production_plan_id(ob_clean_df, spl_clean_df)
        merged_df = merge_dataframes(updated_df, spl_clean_df)

        # Final calculations and merges
        final_df = perform_final_calculations(merged_df)
        final_merged_data_with_status = final_merge_and_status(final_df, rfid_clean_df)

        # Reorder columns and save the final report
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            final_merged_data_with_status.to_excel(writer, index=False, sheet_name='Final Report')
            writer.save()
            processed_data = output.getvalue()

        st.download_button(
            label="Download Finalized Report with Updates",
            data=processed_data,
            file_name="finalizedreport_updated.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()
