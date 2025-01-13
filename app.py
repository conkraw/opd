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

    # Display only rows where 'Student Placed' is 'Yes'
    filtered_df = df[df['Student Placed'] == 'Yes']
    st.write("Filtered Data (Only 'Student Placed' = Yes):")
    st.write(filtered_df)

    # Count how many times each preceptor worked with a student
    preceptor_counts = filtered_df['Preceptor'].value_counts().reset_index()
    preceptor_counts.columns = ['Preceptor', 'Count']

    st.write("Summary Table (Preceptor Student Interactions):")
    st.write(preceptor_counts)

    # Plot the graph
    fig, ax = plt.subplots()
    ax.bar(preceptor_counts['Preceptor'], preceptor_counts['Count'])
    ax.set_xlabel('Preceptor')
    ax.set_ylabel('Count')
    ax.set_title('Preceptor - Student Interaction Count')
    plt.xticks(rotation=45)
    st.pyplot(fig)

    # Allow downloading of the filtered data
    output_file = BytesIO()
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        filtered_df.to_excel(writer, index=False, sheet_name='Filtered Data')
    output_file.seek(0)

    st.download_button(
        label="Download Filtered Data",
        data=output_file,
        file_name="filtered_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

