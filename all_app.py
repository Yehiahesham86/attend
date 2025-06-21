import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime, date

# Process each attendance file between start_date and end_date
def process_attendance_files(file, start_date, end_date):
    df = pd.read_excel(file)
    datetime_col = 'Date/Time'
    
    if datetime_col not in df.columns:
        st.warning(f"'{datetime_col}' column not found in {file.name}")
        return pd.DataFrame()

    df[datetime_col] = pd.to_datetime(df[datetime_col], errors='coerce')
    df.dropna(subset=[datetime_col], inplace=True)
    df['Date'] = df[datetime_col].dt.date

    df_filtered = df[(df['Date'] >= start_date) & (df['Date'] <= end_date)]
    if df_filtered.empty:
        st.warning(f"No data in {file.name} between {start_date} and {end_date}")
        return pd.DataFrame()

    df_filtered['Check_In_Time'] = df_filtered[datetime_col].dt.time
    df_filtered['Check_Out_Time'] = df_filtered[datetime_col].dt.time

    result = df_filtered.groupby('Date').agg(
        Check_In_Time=('Check_In_Time', 'min'),
        Check_Out_Time=('Check_Out_Time', 'max')
    ).reset_index()

    meta_cols = [col for col in ['Name', 'Department', 'No.', 'Date'] if col in df_filtered.columns]
    meta_data = df_filtered[meta_cols].drop_duplicates(subset='Date') if meta_cols else pd.DataFrame({'Date': result['Date']})

    full_result = pd.merge(result, meta_data, on='Date', how='left')
    full_range = pd.date_range(start=start_date, end=end_date).date
    full_result = pd.DataFrame({'Date': full_range}).merge(full_result, on='Date', how='left')

    # Handle missing values and Friday/Saturday logic
    full_result['Check_In_Time'] = full_result['Check_In_Time'].where(full_result['Check_In_Time'].notna(), pd.NaT)
    full_result['Check_Out_Time'] = full_result['Check_Out_Time'].where(full_result['Check_Out_Time'].notna(), pd.NaT)

    full_result.loc[
        full_result['Check_In_Time'] == full_result['Check_Out_Time'],
        'Check_Out_Time'
    ] = pd.to_datetime("17:00").time()

    full_result['Weekday'] = pd.to_datetime(full_result['Date']).dt.weekday
    full_result.loc[full_result['Weekday'] == 4, ['Check_In_Time', 'Check_Out_Time']] = 'Friday'
    full_result.loc[full_result['Weekday'] == 5, ['Check_In_Time', 'Check_Out_Time']] = 'Saturday'

    full_result['Check_In_Time'].fillna("Missing", inplace=True)
    full_result['Check_Out_Time'].fillna("Missing", inplace=True)

    full_result['Employee Name'] = file.name.split('.')[0]
    full_result.drop(columns=['Weekday'], inplace=True)

    return full_result

# Mark holidays
def apply_holiday_names(df, holiday_file):
    if holiday_file is None:
        return df

    try:
        holidays = pd.read_excel(holiday_file)
        holidays['Date'] = pd.to_datetime(holidays['Date']).dt.date
    except Exception as e:
        st.warning(f"Could not read holiday file: {e}")
        return df

    if 'Date' not in holidays or 'Holiday_Name' not in holidays:
        st.warning("Holiday file must have 'Date' and 'Holiday_Name' columns.")
        return df

    holiday_dict = dict(zip(holidays['Date'], holidays['Holiday_Name']))

    def replace(row):
        holiday = holiday_dict.get(row['Date'])
        if holiday:
            if row['Check_In_Time'] == "Missing":
                row['Check_In_Time'] = holiday
            if row['Check_Out_Time'] == "Missing":
                row['Check_Out_Time'] = holiday
        return row

    return df.apply(replace, axis=1)

# Post-process attendance file to calculate hours
def process_excel(file_path):
    df = pd.read_excel(file_path, sheet_name=None)

    with pd.ExcelWriter("processed_full_attendans.xlsx", engine='xlsxwriter') as writer:
        for sheet_name, data in df.items():
            try:
                data['Date'] = pd.to_datetime(data['Date'])
                data['Check_In_Time'] = pd.to_datetime(data['Check_In_Time'], errors='coerce').dt.time
                data['Check_Out_Time'] = pd.to_datetime(data['Check_Out_Time'], errors='coerce').dt.time
                data['Worked_Hours'] = None

                valid_rows = data['Check_In_Time'].notna() & data['Check_Out_Time'].notna()
                check_in = pd.to_datetime(data.loc[valid_rows, 'Check_In_Time'].astype(str), errors='coerce')
                check_out = pd.to_datetime(data.loc[valid_rows, 'Check_Out_Time'].astype(str), errors='coerce')
                data.loc[valid_rows, 'Worked_Hours'] = (check_out - check_in).dt.total_seconds() / 3600

                total_hours = round(data['Worked_Hours'].sum(skipna=True), 2)
                total_row = pd.DataFrame({'Date': ['Total'], 'Worked_Hours': [total_hours]})
                data['Date'] = data['Date'].dt.date

                result = data[['Date', 'Check_In_Time', 'Check_Out_Time', 'Worked_Hours']]
                result = pd.concat([result, total_row], ignore_index=True)
                result.to_excel(writer, sheet_name=sheet_name[:31], index=False)

            except Exception as e:
                st.error(f"Error processing sheet {sheet_name}: {e}")

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

        start_date = st.date_input("Start Date", default_start)
        end_date = st.date_input("End Date", default_end)
        holiday_file = st.file_uploader("Optional: Upload Holidays File (Date, Holiday_Name)", type=["xls", "xlsx"])
        uploaded_files = st.file_uploader("Upload Attendance Excel Files", type=["xls", "xlsx"], accept_multiple_files=True)

        if uploaded_files:
            all_data = []
            for file in uploaded_files:
                try:
                    processed = process_attendance_files(file, start_date, end_date)
                    processed = apply_holiday_names(processed, holiday_file)
                    all_data.append(processed)
                except Exception as e:
                    st.error(f"Error processing {file.name}: {e}")

            if all_data:
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    for i, data in enumerate(all_data):
                        sheet_name = uploaded_files[i].name.split('.')[0][:31]
                        data.to_excel(writer, index=False, sheet_name=sheet_name)

                output.seek(0)
                st.success("✅ Attendance processed!")
                st.download_button(
                    "📥 Download Combined File",
                    data=output,
                    file_name="combined_attendance.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    with tab2:
        st.header("Step 2: Upload File to Calculate Worked Hours")
        file_to_process = st.file_uploader("Upload Combined Attendance Excel File", type=["xlsx"])

        if file_to_process:
            with open("uploaded_file.xlsx", "wb") as f:
                f.write(file_to_process.read())
            result_path = process_excel("uploaded_file.xlsx")
            with open(result_path, "rb") as f:
                st.download_button(
                    "📥 Download Worked Hours Report",
                    data=f,
                    file_name=result_path,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

if __name__ == "__main__":
    main()
