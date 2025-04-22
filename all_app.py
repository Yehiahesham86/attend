import streamlit as st
import pandas as pd
from io import BytesIO

# Function to process attendance data (from the first app)
def process_attendance_files(file):
    df = pd.read_excel(file)

    # Check column names
    st.write(f"Columns in {file.name}: {df.columns}")

    # Ensure that 'Date/Time' column exists
    datetime_col = 'Date/Time'  # Replace with the actual column name
    if datetime_col in df.columns:
        df[datetime_col] = pd.to_datetime(df[datetime_col], errors='coerce')
    else:
        st.warning(f"'{datetime_col}' column not found in {file.name}")
        return pd.DataFrame()  # Return empty DataFrame if column is not found

    # Extract the date
    df['Date'] = df[datetime_col].dt.date

    # Filter the data to include only the 1st to the 26th of the month

    # Filter the DataFrame
     df_filtered = df[df['Date'].apply(lambda x: 1 <= x.day <= 26)]

    # Extract Check-In and Check-Out times
    df_filtered['Check_In_Time'] = df_filtered[datetime_col].dt.time  # Extract time for Check-In
    df_filtered['Check_Out_Time'] = df_filtered[datetime_col].dt.time  # Extract time for Check-Out

    # Group by date and calculate Check-In and Check-Out times
    result = df_filtered.groupby('Date').agg(
        Check_In_Time=('Check_In_Time', 'min'),  # Get the earliest time for Check-In
        Check_Out_Time=('Check_Out_Time', 'max')  # Get the latest time for Check-Out
    ).reset_index()

    # Add 'Name', 'Department', and 'No.' columns from the filtered DataFrame
    df_filtered_grouped = df_filtered[['Name', 'Department', 'No.', 'Date']].drop_duplicates()

    # Merge the grouped data with the result to keep the columns
    full_result = pd.merge(result, df_filtered_grouped, on='Date', how='left')

    # Generate a date range from the 1st to the 26th of the month
    month_start = result['Date'].min().replace(day=1)  # Get the first day of the month
    date_range = pd.date_range(start=month_start, end=month_start.replace(day=26)).date

    # Ensure all dates from the range are present
    full_result = pd.DataFrame({'Date': date_range})  # Create a DataFrame with all dates
    full_result = full_result.merge(result, on='Date', how='left')  # Merge with existing data
    full_result.loc[full_result['Check_In_Time'] == full_result['Check_Out_Time'], 'Check_Out_Time'] = pd.to_datetime("17:00").time()

    # Fill missing Check-In/Check-Out with NaN or placeholders
    full_result['Check_In_Time'].fillna("Missing", inplace=True)
    full_result['Check_Out_Time'].fillna("Missing", inplace=True)

    # Convert 'Date' to datetime if it's not already in datetime format
    full_result['Date'] = pd.to_datetime(full_result['Date'], errors='coerce')

    # Set 'Friday' and 'Saturday' for missing data based on the weekday
    full_result.loc[full_result['Date'].dt.weekday == 4, 'Check_In_Time'] = 'Friday'
    full_result.loc[full_result['Date'].dt.weekday == 5, 'Check_In_Time'] = 'Saturday'
    full_result.loc[full_result['Date'].dt.weekday == 4, 'Check_Out_Time'] = 'Friday'
    full_result.loc[full_result['Date'].dt.weekday == 5, 'Check_Out_Time'] = 'Saturday'

    # Add the employee name from the file name
    full_result['Employee Name'] = file.name.split('.')[0]  # Get the file name (without extension)
    full_result['Date'] = pd.to_datetime(full_result['Date']).dt.date

    return full_result

# Function to process Excel file and calculate worked hours (from second app)
def process_excel(file_path):
    df = pd.read_excel(file_path, sheet_name=None)

    with pd.ExcelWriter("processed_full_attendans.xlsx", engine='xlsxwriter') as output:
        for sheet_name, data in df.items():
            # Convert to datetime, coercing errors (invalid values will be turned to NaT)
            data['Check_In_Time'] = pd.to_datetime(data['Check_In_Time'], errors='coerce')
            data['Check_Out_Time'] = pd.to_datetime(data['Check_Out_Time'], errors='coerce')

            data['Invalid_Row'] = data['Check_In_Time'].isna() | data['Check_Out_Time'].isna()

            data['Worked_Hours'] = None
            data.loc[~data['Invalid_Row'], 'Worked_Hours'] = (data.loc[~data['Invalid_Row'], 'Check_Out_Time'] - 
                                                               data.loc[~data['Invalid_Row'], 'Check_In_Time']).dt.total_seconds() / 3600

            daily_hours = data.groupby('Date')['Worked_Hours'].sum().reset_index()

            daily_hours['Sheet_Name'] = sheet_name

            data = data.merge(daily_hours[['Date', 'Worked_Hours']], on='Date', how='left', suffixes=('', '_y'))

            if 'Worked_Hours_y' in data.columns:
                data.drop(columns="Worked_Hours_y", inplace=True)

            data.drop(columns=['Invalid_Row'], inplace=True)

            data.loc[data['Date'].dt.weekday == 4, 'Check_In_Time'] = 'Friday'
            data.loc[data['Date'].dt.weekday == 5, 'Check_In_Time'] = 'Saturday'
            data.loc[data['Date'].dt.weekday == 4, 'Check_Out_Time'] = 'Friday'
            data.loc[data['Date'].dt.weekday == 5, 'Check_Out_Time'] = 'Saturday'

            data['Check_In_Time'].fillna("Holiday", inplace=True)
            data['Check_Out_Time'].fillna("Holiday", inplace=True)

            total_worked_hours = data['Worked_Hours'].sum()
            total_worked_hours_rounded = round(total_worked_hours, 0)

            total_row = pd.DataFrame({'Date': ['Total'], 'Worked_Hours': [total_worked_hours_rounded]})

            data = pd.concat([data, total_row], ignore_index=True)

            data.to_excel(output, sheet_name=sheet_name, index=False)

    return "processed_full_attendans.xlsx"


# Streamlit app UI
def main():
    st.title("Attendance Data Processor")

    # Tabs using streamlit
    tab1, tab2 = st.tabs(["Process Attendance Data", "Calculate Worked Hours"])

    with tab1:
        st.header("Upload Attendance Data")
        uploaded_files = st.file_uploader("Upload Excel Files", type=["xls", "xlsx"], accept_multiple_files=True)

        if uploaded_files:
            with BytesIO() as output:
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    for uploaded_file in uploaded_files:
                        st.write(f"Processing file: {uploaded_file.name}")
                        try:
                            file_data = process_attendance_files(uploaded_file)
                            if isinstance(file_data, pd.DataFrame) and not file_data.empty:
                                file_data.to_excel(writer, sheet_name=uploaded_file.name.split('.')[0], index=False)
                            else:
                                st.warning(f"Skipping file {uploaded_file.name} due to missing data or invalid format.")
                        except Exception as e:
                            st.error(f"Error processing {uploaded_file.name}: {e}")

                output.seek(0)
                st.download_button(
                    label="Download Combined Data",
                    data=output,
                    file_name="half_attendance_data.xlsx",
                    mime="application/vnd.ms-excel"
                )

    with tab2:
        st.header("Upload Excel File to Calculate Worked Hours")
        uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx"])

        if uploaded_file is not None:
            st.write("Processing your file...")
            with open("uploaded_file.xlsx", "wb") as f:
                f.write(uploaded_file.getbuffer())

            processed_file = process_excel("uploaded_file.xlsx")

            st.write("File processed successfully! You can download the updated file below:")

            with open(processed_file, "rb") as f:
                st.download_button(
                    label="Download Processed Excel",
                    data=f,
                    file_name="full_attendans.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )


if __name__ == "__main__":
    main()
