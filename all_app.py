import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime, date
import re

def process_attendance_files(file, start_date, end_date,holidays_file):
    df = pd.read_excel(file)

    # Check column names

    datetime_col = 'Date/Time'  # Replace with actual column name
    if datetime_col not in df.columns:
        st.warning(f"'{datetime_col}' column not found in {file.name}")
        return pd.DataFrame()

    # Convert to datetime and drop rows with invalid dates
    df[datetime_col] = pd.to_datetime(df[datetime_col], errors='coerce')
    df = df.dropna(subset=[datetime_col])
    df['Date'] = df[datetime_col].dt.date

    # Filter data within selected date range
    df_filtered = df[(df['Date'] >= start_date) & (df['Date'] <= end_date)]

    if df_filtered.empty:
        st.warning(f"No valid data between {start_date} and {end_date} in {file.name}")
        return pd.DataFrame()

    # Extract Check-In and Check-Out times
    df_filtered['Check_In_Time'] = df_filtered[datetime_col].dt.time
    df_filtered['Check_Out_Time'] = df_filtered[datetime_col].dt.time

    # Group by date and calculate earliest/latest times
    result = df_filtered.groupby('Date').agg(
        Check_In_Time=('Check_In_Time', 'min'),
        Check_Out_Time=('Check_Out_Time', 'max')
    ).reset_index()

    # Preserve employee metadata
    meta_cols = [col for col in ['Name', 'Department', 'No.', 'Date'] if col in df_filtered.columns]
    df_filtered_grouped = df_filtered[meta_cols].drop_duplicates()
    full_result = pd.merge(result, df_filtered_grouped, on='Date', how='left')

    # --- Fill missing dates from 25th last month to 26th current month ---
    full_range = pd.date_range(start=start_date, end=end_date).date
    full_result = pd.DataFrame({'Date': full_range}).merge(full_result, on='Date', how='left')

    # Convert Check_In/Out to datetime.time or placeholders
    full_result['Check_In_Time'] = full_result['Check_In_Time'].where(full_result['Check_In_Time'].notna(), pd.NaT)
    full_result['Check_Out_Time'] = full_result['Check_Out_Time'].where(full_result['Check_Out_Time'].notna(), pd.NaT)

    # Fill same time for in/out if both exist but are the same
    full_result.loc[
        full_result['Check_In_Time'] == full_result['Check_Out_Time'],
        'Check_Out_Time'
    ] = pd.to_datetime("17:00").time()

    # Set weekend names for missing values
    full_result['Date'] = pd.to_datetime(full_result['Date'])
    full_result['Weekday'] = full_result['Date'].dt.weekday

    friday = full_result['Weekday'] == 4
    saturday = full_result['Weekday'] == 5

    full_result.loc[friday, ['Check_In_Time', 'Check_Out_Time']] = 'Friday'
    full_result.loc[saturday, ['Check_In_Time', 'Check_Out_Time']] = 'Saturday'

    # Fill missing with "Missing" or "Holiday"
    full_result['Check_In_Time'].fillna("Missing", inplace=True)
    full_result['Check_Out_Time'].fillna("Missing", inplace=True)

    # Add Employee Name from file name
    full_result['Employee Name'] = file.name.split('.')[0]
    full_result['Date'] = full_result['Date'].dt.date

    # Drop helper column
    full_result.drop(columns=['Weekday'], inplace=True)
    
    # Process holidays
    holidays_df = pd.read_excel(holidays_file)
    holidays_df['Date'] = pd.to_datetime(holidays_df['Date']).dt.date

    # Merge the DataFrames on 'Date' to add the Holiday_Name
    full_result = pd.merge(full_result, holidays_df, on='Date', how='left')

    
    # Replace Check_In_Time and Check_Out_Time with the actual Holiday_Name
    holiday_mask = full_result['Holiday_Name'].notna() & (full_result['Holiday_Name'] != '')
    full_result.loc[holiday_mask, 'Check_In_Time'] = full_result.loc[holiday_mask, 'Holiday_Name']
    full_result.loc[holiday_mask, 'Check_Out_Time'] = full_result.loc[holiday_mask, 'Holiday_Name']

    # Remove the Holiday_Name column
    full_result.drop(columns=['Holiday_Name','Day','Name','Department','No.','Employee Name'], inplace=True)

    return full_result

# Function to process Excel file and calculate worked hours (from second app)
def process_excel(file_path, holidays_file):
    # Read the Excel file with multiple sheets
    df = pd.read_excel(file_path, sheet_name=None)
    # Clean the sheet name to only have letters
    
    summary = []

    # Function to clean sheet names (only letters)
    def clean_sheet_name(sheet_name):
        return re.sub(r'[^a-zA-Z]', '', sheet_name)

    # Create an Excel writer to save the processed data
    with pd.ExcelWriter("processed_full_attendans.xlsx", engine='xlsxwriter') as output:
        
        for sheet_name, data in df.items():
            cleaned_sheet_name = clean_sheet_name(sheet_name)
            # Convert Check_In_Time and Check_Out_Time to datetime, errors='coerce' to handle invalid values
            data['Check_In_Time'] = pd.to_datetime(data['Check_In_Time'], errors='coerce')
            data['Check_Out_Time'] = pd.to_datetime(data['Check_Out_Time'], errors='coerce')

            # Extract only the time part from datetime
            data['Check_In_Time'] = data['Check_In_Time'].dt.time
            data['Check_Out_Time'] = data['Check_Out_Time'].dt.time

            # Create an 'Invalid_Row' column to flag rows where Check_In or Check_Out is missing
            data['Invalid_Row'] = data['Check_In_Time'].isna() | data['Check_Out_Time'].isna()

            # Calculate Worked_Hours, only for valid rows
            data['Worked_Hours'] = None
            data.loc[~data['Invalid_Row'], 'Worked_Hours'] = (
                pd.to_datetime(data.loc[~data['Invalid_Row'], 'Check_Out_Time'].astype(str), errors='coerce') - 
                pd.to_datetime(data.loc[~data['Invalid_Row'], 'Check_In_Time'].astype(str), errors='coerce')
            ).dt.total_seconds() / 3600  # Convert time difference to hours

            # Group by Date and sum the Worked_Hours
            daily_hours = data.groupby('Date')['Worked_Hours'].sum().reset_index()
            daily_hours['Sheet_Name'] = sheet_name

            # Merge daily hours back into the original data
            data = data.merge(daily_hours[['Date', 'Worked_Hours']], on='Date', how='left', suffixes=('', '_y'))

            # Drop extra column if it exists
            if 'Worked_Hours_y' in data.columns:
                data.drop(columns="Worked_Hours_y", inplace=True)

            # Remove Invalid_Row column
            data.drop(columns=['Invalid_Row'], inplace=True)

            # Handle weekends (Friday and Saturday)
            data.loc[data['Date'].dt.weekday == 4, 'Check_In_Time'] = 'Friday'
            data.loc[data['Date'].dt.weekday == 5, 'Check_In_Time'] = 'Saturday'
            data.loc[data['Date'].dt.weekday == 4, 'Check_Out_Time'] = 'Friday'
            data.loc[data['Date'].dt.weekday == 5, 'Check_Out_Time'] = 'Saturday'

            # Fill missing values in Check_In_Time and Check_Out_Time
            data['Check_In_Time'].fillna("Holiday", inplace=True)
            data['Check_Out_Time'].fillna("Holiday", inplace=True)
            data = pd.DataFrame({
                'Date': data['Date'].dt.date,
                'Check_In_Time': data['Check_In_Time'],
                'Check_Out_Time': data['Check_Out_Time'],
                'Worked_Hours': data['Worked_Hours'].round(2)
            })
            # Calculate total worked hours for all data
            total_worked_hours = data['Worked_Hours'].sum()
            total_worked_hours_rounded = round(total_worked_hours, 0)

            # Add a total row
            total_row = pd.DataFrame({'Date': ['Total'], 'Worked_Hours': [total_worked_hours_rounded]})
            data = pd.concat([data, total_row], ignore_index=True)

            # Process holidays
            holidays_df = pd.read_excel(holidays_file)
            holidays_df['Date'] = pd.to_datetime(holidays_df['Date']).dt.date

            # Merge the DataFrames on 'Date' to add the Holiday_Name
            data = pd.merge(data, holidays_df, on='Date', how='left')

            # Replace Check_In_Time and Check_Out_Time with the actual Holiday_Name
            holiday_mask = data['Holiday_Name'].notna() & (data['Holiday_Name'] != '')
            data.loc[holiday_mask, 'Check_In_Time'] = data.loc[holiday_mask, 'Holiday_Name']
            data.loc[holiday_mask, 'Check_Out_Time'] = data.loc[holiday_mask, 'Holiday_Name']

            # Remove the Holiday_Name column
            data.drop(columns=['Holiday_Name','Day'], inplace=True)
            
            
            summary.append({'Employee': cleaned_sheet_name, 'Total Worked Hours': total_worked_hours_rounded})

            total_row = pd.DataFrame({'Date': ['Total'], 'Worked_Hours': [total_worked_hours_rounded]})
            data = pd.concat([data, total_row], ignore_index=True)

            # Clean the sheet name to only have letters

            # Write each sheet's data to the output file
            data.to_excel(output, sheet_name=cleaned_sheet_name[:31], index=False)

        summary_df = pd.DataFrame(summary)
        summary_df.to_excel(output, sheet_name='Summary', index=False)

    # Return the path to the processed file
    return "processed_full_attendans.xlsx"

# Streamlit app UI
def main():
    st.set_page_config(page_title="Attendance Data Processor", page_icon="📊", layout="wide")
    
    # Adding a title and description with custom styling
    st.markdown("""
        <h1 style='text-align:center;'>Attendance Data Processor <span style="font-size: 50px;">📊</span></h1>
        <p style='text-align:center;'>Process and calculate attendance data effortlessly</p>
    """, unsafe_allow_html=True)
    
    # Tabs using streamlit
    tab1, tab2 = st.tabs(["📅 Process Attendance Data", "🕒 Calculate Worked Hours"])

    with tab1:
        st.header("Upload Attendance Data")
        
        # Default date range: from 25th of last month to 26th of this month
        today = date.today()
        default_start = (today.replace(day=1) - pd.Timedelta(days=1)).replace(day=25)
        default_end = today.replace(day=26)

        # Create two columns for date input side by side
        col1, col2 = st.columns(2)
        
        with col1:
            start_date = st.date_input("Start Date", value=default_start)
        
        with col2:
            end_date = st.date_input("End Date", value=default_end)

        # File uploaders with better descriptions
        holidays_file = st.file_uploader("Holiday File (Excel)", type=["xls", "xlsx"], accept_multiple_files=False)
        uploaded_files = st.file_uploader("Upload Attendance Excel Files", type=["xls", "xlsx"], accept_multiple_files=True)

        if uploaded_files:
            all_data = []

            for uploaded_file in uploaded_files:
                try:
                    file_data = process_attendance_files(uploaded_file, start_date, end_date, holidays_file)
                    all_data.append(file_data)
                except Exception as e:
                    st.error(f"Error processing file {uploaded_file.name}: {e}")

            if all_data:
                combined_df = pd.concat(all_data, ignore_index=True)
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    for i, file_data in enumerate(all_data):
                        file_name = uploaded_files[i].name.split('.')[0]
                        file_data.to_excel(writer, index=False, sheet_name=file_name)

                output.seek(0)  # Rewind the buffer

                st.success("Data processed successfully! 🎉")

                st.download_button(
                    label="📥 Download Processed Excel File",
                    data=output,
                    file_name=f"Half Attendance - {datetime.today().strftime('%Y-%m-%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    with tab2:
        st.header("Upload Excel File to Calculate Worked Hours")

        holidays_file = st.file_uploader("Holidays File (Excel)", type=["xls", "xlsx"], accept_multiple_files=False)
        uploaded_file = st.file_uploader("Upload Attendance Excel (Half)", type=["xlsx"])

        if uploaded_file is not None:
            st.write("Processing your file... 🔄")
            with open("uploaded_file.xlsx", "wb") as f:
                f.write(uploaded_file.getbuffer())

            processed_file = process_excel("uploaded_file.xlsx", holidays_file)

            st.write("File processed successfully! 🎉 You can download the updated file below:")

            with open(processed_file, "rb") as f:
                st.download_button(
                    label="📥 Download Processed Worked Hours File",
                    data=f,
                    file_name=f"Full Attendance - {datetime.today().strftime('%Y-%m-%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

if __name__ == "__main__":
    main()
