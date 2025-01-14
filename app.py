import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from io import BytesIO
import matplotlib.pyplot as plt

# Function to process each area based on the date column and starting row
def process_area(sheet, date_column, date_row, start_row, location):
    date = sheet[f'{date_column}{date_row}'].value
    am_names = [sheet[f'{date_column}{i}'].value for i in range(start_row, start_row + 10)]
    pm_names = [sheet[f'{date_column}{i}'].value for i in range(start_row + 10, start_row + 20)]
    names = [(name, 'AM') for name in am_names] + [(name, 'PM') for name in pm_names]

    processed_data = []
    for name, type_ in names:
        if name:
            preceptor, student = (name.split(' ~ ') if ' ~ ' in name else (name, None))
            student_placed = 'Yes' if student else 'No'
            student_type = None
            if student:
                if '(MD)' in student:
                    student_type = 'MD'
                elif '(PA)' in student:
                    student_type = 'PA'

            processed_data.append({
                'Date': date,
                'Type': type_,
                'Description': name,
                'Preceptor': preceptor.strip(),
                'Student': student.strip() if student else None,
                'Student Placed': student_placed,
                'Student Type': student_type,
                'Location': location
            })
    return processed_data

# Streamlit app
st.title('OPD Data Processor')

uploaded_files = st.file_uploader("Choose Excel files", type="xlsx", accept_multiple_files=True)

if uploaded_files:
    all_data = []
    for uploaded_file in uploaded_files:
        wb = load_workbook(uploaded_file)
        for sheet_name in wb.sheetnames:
            sheet = wb[sheet_name]
            for date_column in ['B', 'C', 'D', 'E', 'F', 'G', 'H']:
                for date_row in [4, 28, 52, 76]:
                    start_row = date_row + 2
                    area_data = process_area(sheet, date_column, date_row, start_row, sheet_name)
                    all_data.extend(area_data)

    df = pd.DataFrame(all_data)

    # Exclude rows with "COM CLOSED" or "Closed" in the Description column
    df = df[~df['Description'].str.contains('COM CLOSED|Closed', case=False, na=False)]

    # Add weekday column
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
    df['Weekday'] = df['Date'].dt.day_name()

    # Filter rows where 'Student Placed' is 'Yes'
    filtered_df = df[df['Student Placed'] == 'Yes']

    # Clean Preceptor names
    df['Preceptor'] = df['Preceptor'].str.strip()
    df['Preceptor'] = df['Preceptor'].str.replace(r' ~$', '', regex=True)

    filtered_df['Preceptor'] = filtered_df['Preceptor'].str.strip()
    filtered_df['Preceptor'] = filtered_df['Preceptor'].str.replace(r' ~$', '', regex=True)

    # Calculate available and used shifts
    total_shifts = df.groupby(['Location', 'Weekday', 'Type']).size().reset_index(name='Total Shifts')
    open_shifts = df[df['Student Placed'] == 'No'].groupby(['Location', 'Weekday', 'Type']).size().reset_index(name='Open Shifts')

    # Merge total shifts and open shifts
    location_shifts = pd.merge(total_shifts, open_shifts, on=['Location', 'Weekday', 'Type'], how='left')
    location_shifts['Open Shifts'] = location_shifts['Open Shifts'].fillna(0)

    # Calculate percentage open shifts by location
    location_shifts['Percentage Open'] = (location_shifts['Open Shifts'] / location_shifts['Total Shifts']) * 100

    # Total percentage (all locations combined)
    total_shifts_summary = location_shifts.groupby(['Weekday', 'Type'])[['Total Shifts', 'Open Shifts']].sum().reset_index()
    total_shifts_summary['Percentage Open'] = (total_shifts_summary['Open Shifts'] / total_shifts_summary['Total Shifts']) * 100

    # Reshape data for display
    location_shifts_summary = location_shifts.pivot_table(
        index=['Weekday', 'Type'], columns='Location', values='Percentage Open', aggfunc='mean'
    ).fillna(0).reset_index()

    total_shifts_summary_pivot = total_shifts_summary.pivot(index='Weekday', columns='Type', values='Percentage Open').fillna(0).reset_index()

    # Display percentage open shifts by location
    st.write("Percentage of Open Shifts by Weekday, Type (AM/PM), and Location:")
    st.write(location_shifts_summary)

    # Display total percentage open shifts across all locations
    st.write("Total Percentage of Open Shifts by Weekday and Type (AM/PM):")
    st.write(total_shifts_summary_pivot)

    # Include all data in the downloadable Excel file
    output_file = BytesIO()
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Combined Dataset')  # Full dataset
        location_shifts_summary.to_excel(writer, index=False, sheet_name='Location Open Shifts')  # Open shifts by location
        total_shifts_summary_pivot.to_excel(writer, index=False, sheet_name='Total Open Shifts')  # Total open shifts
    output_file.seek(0)

    st.download_button(
        label="Download Combined and Summary Data",
        data=output_file,
        file_name="combined_and_open_shifts_by_location.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

