import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime, date
import re


def process_attendance_files(file, start_date, end_date, holidays_file):
    df = pd.read_excel(file)
    datetime_col = 'Date/Time'
    if datetime_col not in df.columns:
        st.warning(f"'{datetime_col}' column not found in {file.name}")
        return pd.DataFrame()

    df[datetime_col] = pd.to_datetime(df[datetime_col], errors='coerce')
    df = df.dropna(subset=[datetime_col])
    df['Date'] = df[datetime_col].dt.date
    df_filtered = df[(df['Date'] >= start_date) & (df['Date'] <= end_date)]

    if df_filtered.empty:
        st.warning(f"No valid data between {start_date} and {end_date} in {file.name}")
        return pd.DataFrame()

    df_filtered['Check_In_Time'] = df_filtered[datetime_col].dt.time
    df_filtered['Check_Out_Time'] = df_filtered[datetime_col].dt.time

    result = df_filtered.groupby('Date').agg(
        Check_In_Time=('Check_In_Time', 'min'),
        Check_Out_Time=('Check_Out_Time', 'max')
    ).reset_index()

    meta_cols = [col for col in ['Name', 'Department', 'No.', 'Date'] if col in df_filtered.columns]
    df_filtered_grouped = df_filtered[meta_cols].drop_duplicates()
    full_result = pd.merge(result, df_filtered_grouped, on='Date', how='left')

    full_range = pd.date_range(start=start_date, end=end_date).date
    full_result = pd.DataFrame({'Date': full_range}).merge(full_result, on='Date', how='left')

    full_result['Check_In_Time'] = full_result['Check_In_Time'].where(full_result['Check_In_Time'].notna(), pd.NaT)
    full_result['Check_Out_Time'] = full_result['Check_Out_Time'].where(full_result['Check_Out_Time'].notna(), pd.NaT)
    full_result.loc[full_result['Check_In_Time'] == full_result['Check_Out_Time'], 'Check_Out_Time'] = "No Check Out"

    full_result['Date'] = pd.to_datetime(full_result['Date'])
    full_result['Weekday'] = full_result['Date'].dt.weekday
    full_result.loc[full_result['Weekday'] == 4, ['Check_In_Time', 'Check_Out_Time']] = 'Friday'
    full_result.loc[full_result['Weekday'] == 5, ['Check_In_Time', 'Check_Out_Time']] = 'Saturday'

    full_result['Check_In_Time'].fillna("Missing", inplace=True)
    full_result['Check_Out_Time'].fillna("Missing", inplace=True)
    full_result['Employee Name'] = file.name.split('.')[0]
    full_result['Date'] = full_result['Date'].dt.date
    full_result.drop(columns=['Weekday'], inplace=True)

    holidays_df = pd.read_excel(holidays_file)
    holidays_df['Date'] = pd.to_datetime(holidays_df['Date']).dt.date
    full_result = pd.merge(full_result, holidays_df, on='Date', how='left')

    holiday_mask = full_result['Holiday_Name'].notna() & (full_result['Holiday_Name'] != '')
    full_result.loc[holiday_mask, 'Check_In_Time'] = full_result.loc[holiday_mask, 'Holiday_Name']
    full_result.loc[holiday_mask, 'Check_Out_Time'] = full_result.loc[holiday_mask, 'Holiday_Name']

    full_result.drop(columns=['Holiday_Name','Day','Name','Department','No.','Employee Name'], inplace=True)

    return full_result


def process_excel(file_path, holidays_file):
    df = pd.read_excel(file_path, sheet_name=None)
    summary = []

    def clean_sheet_name(sheet_name):
        return re.sub(r'[^a-zA-Z]', '', sheet_name)

    with pd.ExcelWriter("processed_full_attendans.xlsx", engine='xlsxwriter') as output:
        for sheet_name, data in df.items():
            data['Check_In_Time'] = pd.to_datetime(data['Check_In_Time'], errors='coerce')
            data['Check_Out_Time'] = pd.to_datetime(data['Check_Out_Time'], errors='coerce')
            data['Check_In_Time'] = data['Check_In_Time'].dt.time
            data['Check_Out_Time'] = data['Check_Out_Time'].dt.time

            data['Invalid_Row'] = data['Check_In_Time'].isna() | data['Check_Out_Time'].isna()

            data['Worked_Hours'] = None
            data.loc[~data['Invalid_Row'], 'Worked_Hours'] = (
                pd.to_datetime(data.loc[~data['Invalid_Row'], 'Check_Out_Time'].astype(str), errors='coerce') - 
                pd.to_datetime(data.loc[~data['Invalid_Row'], 'Check_In_Time'].astype(str), errors='coerce')
            ).dt.total_seconds() / 3600

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

            data['Check_In_Time'].fillna("No Check In", inplace=True)
            data['Check_Out_Time'].fillna("No Check Out", inplace=True)

            data = pd.DataFrame({
                'Date': data['Date'].dt.date,
                'Check_In_Time': data['Check_In_Time'],
                'Check_Out_Time': data['Check_Out_Time'],
                'Worked_Hours': data['Worked_Hours'].round(2)
            })

            total_worked_hours = data['Worked_Hours'].sum()
            total_worked_hours_rounded = round(total_worked_hours, 0)
            summary.append({'Employee': clean_sheet_name(sheet_name), 'Total Worked Hours': total_worked_hours_rounded})

            total_row = pd.DataFrame({'Date': ['Total'], 'Worked_Hours': [total_worked_hours_rounded]})
            data = pd.concat([data, total_row], ignore_index=True)

            holidays_df = pd.read_excel(holidays_file)
            holidays_df['Date'] = pd.to_datetime(holidays_df['Date']).dt.date
            data = pd.merge(data, holidays_df, on='Date', how='left')

            holiday_mask = data['Holiday_Name'].notna() & (data['Holiday_Name'] != '')
            data.loc[holiday_mask, 'Check_In_Time'] = data.loc[holiday_mask, 'Holiday_Name']
            data.loc[holiday_mask, 'Check_Out_Time'] = data.loc[holiday_mask, 'Holiday_Name']
            data.drop(columns=['Holiday_Name','Day'], inplace=True)

            cleaned_sheet_name = clean_sheet_name(sheet_name)[:31]
            data.to_excel(output, sheet_name=cleaned_sheet_name, index=False)

        summary_df = pd.DataFrame(summary)
        summary_df.to_excel(output, sheet_name='Summary', index=False)

    return "processed_full_attendans.xlsx"


# Streamlit UI
def main():
    st.set_page_config(page_title="Attendance Processor", layout="wide")
    st.title("🕒 Attendance Data Processor")

    tab1, tab2 = st.tabs(["📁 Process Attendance", "📊 Calculate Worked Hours"])

    with tab1:
        st.header("Step 1: Upload Attendance Files")
        today = date.today()
        default_start = (today.replace(day=1) - pd.Timedelta(days=1)).replace(day=25)
        default_end = today.replace(day=26)

        col1, col2 = st.columns(2)
        with col1:
            start_date = st.date_input("Start Date", default_start)
        with col2:
            end_date = st.date_input("End Date", default_end)

        holidays_file = st.file_uploader("Upload Holidays File", type=["xls", "xlsx"])
        uploaded_files = st.file_uploader("Upload Attendance Files", type=["xls", "xlsx"], accept_multiple_files=True)

        if uploaded_files and holidays_file:
            all_data = []
            for file in uploaded_files:
                try:
                    result = process_attendance_files(file, start_date, end_date, holidays_file)
                    all_data.append((file.name.split('.')[0][:31], result))
                except Exception as e:
                    st.error(f"Error processing file {file.name}: {e}")

            if all_data:
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    for name, df in all_data:
                        df.to_excel(writer, index=False, sheet_name=name)

                output.seek(0)
                st.success("✅ Attendance processed!")
                st.download_button(
                    "📥 Download Combined File",
                    data=output,
                    file_name=f"Half Attendance - {datetime.today().strftime('%Y-%m-%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    with tab2:
        st.header("Step 2: Upload File to Calculate Worked Hours")
        holidays_file = st.file_uploader("Upload Holidays File", type=["xls", "xlsx"], key="holiday_hours")
        file_to_process = st.file_uploader("Upload Combined Attendance Excel File", type=["xlsx"])

        if file_to_process and holidays_file:
            with open("uploaded_file.xlsx", "wb") as f:
                f.write(file_to_process.read())

            processed_file = process_excel("uploaded_file.xlsx", holidays_file)
            with open(processed_file, "rb") as f:
                st.download_button(
                    "📥 Download Worked Hours Report",
                    data=f,
                    file_name=f"Full Attendance - {datetime.today().strftime('%Y-%m-%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )


if __name__ == "__main__":
    main()