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

    # Calculate Days Worked by Preceptor
    filtered_df['Half Day'] = filtered_df['Type'].apply(lambda x: 0.5 if x in ['AM', 'PM'] else 0)
    days_worked = (
        filtered_df.groupby(['Preceptor', 'Date'])['Half Day']
        .sum()
        .reset_index()
        .rename(columns={'Half Day': 'Total Day Fraction'})
    )

    preceptor_days_summary = (
        days_worked.groupby('Preceptor')['Total Day Fraction']
        .sum()
        .reset_index()
        .rename(columns={'Total Day Fraction': 'Total Days'})
    )

    # Calculate available and used shifts
    available_shifts = (
        df.groupby(['Preceptor', 'Date', 'Type'])
        .size()
        .reset_index(name='Available Shifts')
    )
    available_shifts = (
        available_shifts.groupby('Preceptor')['Available Shifts']
        .sum()
        .reset_index()
    )

    used_shifts = (
        filtered_df.groupby(['Preceptor', 'Date', 'Type'])
        .size()
        .reset_index(name='Used Shifts')
    )
    used_shifts = (
        used_shifts.groupby('Preceptor')['Used Shifts']
        .sum()
        .reset_index()
    )

    # Count MD and PA shifts
    md_shifts = filtered_df[filtered_df['Student Type'] == 'MD'].groupby('Preceptor').size().reset_index(name='MD Shifts')
    pa_shifts = filtered_df[filtered_df['Student Type'] == 'PA'].groupby('Preceptor').size().reset_index(name='PA Shifts')

    # Merge all summary data
    shifts_summary = pd.merge(available_shifts, used_shifts, on='Preceptor', how='left')
    shifts_summary = pd.merge(shifts_summary, md_shifts, on='Preceptor', how='left')
    shifts_summary = pd.merge(shifts_summary, pa_shifts, on='Preceptor', how='left')
    shifts_summary['Used Shifts'] = shifts_summary['Used Shifts'].fillna(0)
    shifts_summary['MD Shifts'] = shifts_summary['MD Shifts'].fillna(0)
    shifts_summary['PA Shifts'] = shifts_summary['PA Shifts'].fillna(0)
    shifts_summary['Unused Shifts'] = shifts_summary['Available Shifts'] - shifts_summary['Used Shifts']
    shifts_summary['Percentage of Shifts Filled'] = (
        (shifts_summary['Used Shifts'] / shifts_summary['Available Shifts']) * 100
    )
    shifts_summary['Percentage MD'] = (
        (shifts_summary['MD Shifts'] / shifts_summary['Used Shifts']) * 100
    ).fillna(0)
    shifts_summary['Percentage PA'] = (
        (shifts_summary['PA Shifts'] / shifts_summary['Used Shifts']) * 100
    ).fillna(0)

    # Calculate total and percentage open shifts by weekday and type (AM/PM)
    total_shifts = (
        df.groupby(['Weekday', 'Type'])
        .size()
        .reset_index(name='Total Shifts')
    )

    open_shifts = (
        df[df['Student Placed'] == 'No']
        .groupby(['Weekday', 'Type'])
        .size()
        .reset_index(name='Open Shifts')
    )

    # Merge total shifts and open shifts
    weekday_shifts = pd.merge(total_shifts, open_shifts, on=['Weekday', 'Type'], how='left')
    weekday_shifts['Open Shifts'] = weekday_shifts['Open Shifts'].fillna(0)
    weekday_shifts['Percentage Open'] = (weekday_shifts['Open Shifts'] / weekday_shifts['Total Shifts']) * 100

    # Reshape data for better display
    weekday_shifts_summary = weekday_shifts.pivot(index='Weekday', columns='Type', values='Percentage Open').fillna(0).reset_index()

    # Display the percentage open shifts table
    st.write("Percentage of Open Shifts by Weekday and Type (AM/PM):")
    st.write(weekday_shifts_summary)

    # Display the shifts summary table
    st.write("Summary Table (Available vs. Used Shifts by Preceptor):")
    st.write(shifts_summary)

    # Include all data in the downloadable Excel file
    output_file = BytesIO()
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Combined Dataset')  # Full dataset
        weekday_shifts_summary.to_excel(writer, index=False, sheet_name='Percentage Open Shifts')  # Percentage open shifts
        shifts_summary.to_excel(writer, index=False, sheet_name='Shifts Summary')  # Shifts summary
    output_file.seek(0)

    st.download_button(
        label="Download Combined and Summary Data",
        data=output_file,
        file_name="combined_and_summary_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

