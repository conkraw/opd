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

    # Calculate open shifts (unassigned shifts)
    open_shifts = df[df['Student Placed'] == 'No']

    # Calculate average open shifts by weekday and type (AM/PM)
    avg_open_shifts = (
        open_shifts.groupby(['Weekday', 'Type'])
        .size()
        .reset_index(name='Open Shifts')
        .groupby(['Weekday', 'Type'])['Open Shifts']
        .mean()
        .unstack()
        .fillna(0)
        .reset_index()
    )

    # Display the average open shifts table
    st.write("Average Open Shifts by Weekday and Type:")
    st.write(avg_open_shifts)

    # Include results in the downloadable Excel file
    output_file = BytesIO()
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Combined Dataset')  # Full dataset
        avg_open_shifts.to_excel(writer, index=False, sheet_name='Avg Open Shifts')  # Avg open shifts
    output_file.seek(0)

    st.download_button(
        label="Download Combined Dataset and Open Shifts Data",
        data=output_file,
        file_name="combined_and_open_shifts_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


