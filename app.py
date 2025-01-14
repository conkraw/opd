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

    # Filter rows where 'Student Placed' is 'Yes'
    filtered_df = df[df['Student Placed'] == 'Yes']

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

    # Merge the available and used shifts into one summary table
    shifts_summary = pd.merge(available_shifts, used_shifts, on='Preceptor', how='left')
    shifts_summary['Used Shifts'] = shifts_summary['Used Shifts'].fillna(0)
    shifts_summary['Unused Shifts'] = shifts_summary['Available Shifts'] - shifts_summary['Used Shifts']

    # Display the shifts summary table
    st.write("Summary Table (Available vs. Used Shifts by Preceptor):")
    st.write(shifts_summary)

    # Plot the graph for available and used shifts
    fig, ax = plt.subplots()
    ax.bar(shifts_summary['Preceptor'], shifts_summary['Available Shifts'], label='Available Shifts', alpha=0.7)
    ax.bar(shifts_summary['Preceptor'], shifts_summary['Used Shifts'], label='Used Shifts', alpha=0.7)
    ax.set_xlabel('Preceptor')
    ax.set_ylabel('Shifts')
    ax.set_title('Available vs. Used Shifts by Preceptor')
    ax.legend()
    plt.xticks(rotation=45)
    st.pyplot(fig)

    # Include shifts summary in the download file
    output_file = BytesIO()
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        filtered_df.to_excel(writer, index=False, sheet_name='Combined Dataset')
        days_worked.to_excel(writer, index=False, sheet_name='Days Worked Detail')
        preceptor_days_summary.to_excel(writer, index=False, sheet_name='Total Days Summary')
        shifts_summary.to_excel(writer, index=False, sheet_name='Shifts Summary')
    output_file.seek(0)

    st.download_button(
        label="Download Combined and Summary Data",
        data=output_file,
        file_name="combined_and_summary_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


