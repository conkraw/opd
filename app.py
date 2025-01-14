import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from io import BytesIO
import matplotlib.pyplot as plt
import numpy as np

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

    # Sort weekdays (Monday to Sunday)
    weekday_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
    df['Weekday'] = pd.Categorical(df['Weekday'], categories=weekday_order, ordered=True)

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

    # Individual Preceptor Percentage
    preceptor_summary = df[df['Student Placed'] == 'Yes'].groupby(['Preceptor', 'Type'])['Student Placed'].count()
    preceptor_summary = preceptor_summary.div(
        df.groupby(['Preceptor', 'Type'])['Student Placed'].size()
    ).reset_index(name='Percentage Filled') * 100

    # Reshape data for display
    total_shifts_summary_pivot = total_shifts_summary.pivot(index='Weekday', columns='Type', values='Percentage Open').fillna(0).reset_index()

    # Plot total percentage open shifts by AM/PM
    fig, ax = plt.subplots(figsize=(10, 6))
    for shift_type in ['AM', 'PM']:
        ax.plot(total_shifts_summary_pivot['Weekday'], total_shifts_summary_pivot[shift_type], marker='o', label=shift_type)
    ax.set_title('Total Percentage of Open Shifts by Weekday (AM/PM)')
    ax.set_ylabel('Percentage Open Shifts')
    ax.set_xlabel('Weekday')
    plt.xticks(rotation=45)
    ax.legend()
    st.pyplot(fig)

    # Plot percentage open shifts by location and type
    locations = location_shifts['Location'].unique()
    fig, ax = plt.subplots(figsize=(12, 8))
    x = np.arange(len(weekday_order))
    bar_width = 0.2
    for i, location in enumerate(locations):
        data = location_shifts[location_shifts['Location'] == location]
        data = data.groupby(['Weekday', 'Type'])['Percentage Open'].mean().unstack()
        ax.bar(x + (i * bar_width), data['AM'], bar_width, label=f'{location} - AM')
        ax.bar(x + (i * bar_width) + (bar_width / 2), data['PM'], bar_width, label=f'{location} - PM')
    ax.set_title('Percentage of Open Shifts by Location and Type')
    ax.set_ylabel('Percentage Open Shifts')
    ax.set_xlabel('Weekday')
    ax.set_xticks(x + (len(locations) - 1) * bar_width / 2)
    ax.set_xticklabels(weekday_order)
    plt.xticks(rotation=45)
    ax.legend(loc='upper left', bbox_to_anchor=(1, 1))
    st.pyplot(fig)

    # Individual Preceptor Graphs
    for preceptor in preceptor_summary['Preceptor'].unique():
        preceptor_data = preceptor_summary[preceptor_summary['Preceptor'] == preceptor]
        fig, ax = plt.subplots(figsize=(8, 5))
        ax.bar(preceptor_data['Type'], preceptor_data['Percentage Filled'], color=['skyblue', 'orange'])
        ax.set_title(f'{preceptor} - Percentage of Shifts with Students')
        ax.set_ylabel('Percentage of Shifts')
        ax.set_xlabel('Shift Type')
        st.pyplot(fig)

    # Include all data in the downloadable Excel file
    output_file = BytesIO()
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Combined Dataset')  # Full dataset
        location_shifts.to_excel(writer, index=False, sheet_name='Location Open Shifts')  # Open shifts by location
        total_shifts_summary.to_excel(writer, index=False, sheet_name='Total Open Shifts')  # Total open shifts
        preceptor_summary.to_excel(writer, index=False, sheet_name='Preceptor Filled Shifts')  # Preceptor shifts
    output_file.seek(0)

    st.download_button(
        label="Download Combined and Summary Data",
        data=output_file,
        file_name="combined_and_open_shifts_with_preceptors.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

