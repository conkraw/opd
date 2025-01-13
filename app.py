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

    # Display the table of total days worked
    st.write("Summary Table (Total Days Worked by Preceptor):")
    st.write(preceptor_days_summary)

    # Plot the graph for total days worked
    fig, ax = plt.subplots()
    ax.bar(preceptor_days_summary['Preceptor'], preceptor_days_summary['Total Days'])
    ax.set_xlabel('Preceptor')
    ax.set_ylabel('Total Days Worked')
    ax.set_title('Total Days Worked by Preceptor')
    plt.xticks(rotation=45)
    st.pyplot(fig)

    # Allow download of the detailed days worked data
    output_file = BytesIO()
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        days_worked.to_excel(writer, index=False, sheet_name='Days Worked Detail')
        preceptor_days_summary.to_excel(writer, index=False, sheet_name='Total Days Summary')
    output_file.seek(0)

    st.download_button(
        label="Download Days Worked Data",
        data=output_file,
        file_name="days_worked_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
