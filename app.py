import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from io import BytesIO

# Function to process each area based on the date column and starting row
def process_area(sheet, date_column, date_row, start_row, location):
    # Extract the date from the date cell (e.g., B4, C4, D4)
    date = sheet[f'{date_column}{date_row}'].value
    
    # Extract the AM and PM names from the corresponding rows (e.g., B6:B15 for AM, B16:B25 for PM)
    am_names = [sheet[f'{date_column}{i}'].value for i in range(start_row, start_row + 10)]  # B6 to B15 for AM
    pm_names = [sheet[f'{date_column}{i}'].value for i in range(start_row + 10, start_row + 20)]  # B16 to B25 for PM
    
    # Combine AM and PM names with their respective types (AM/PM)
    names = [(name, 'AM') for name in am_names] + [(name, 'PM') for name in pm_names]

    # Create a list to store the processed data in the desired format
    processed_data = []
    for name, type_ in names:
        if name:
            # Split the name based on the pattern 'Preceptor ~ Student'
            preceptor, student = (name.split(' ~ ') if ' ~ ' in name else (name, None))
            
            # Check if the student exists
            student_placed = 'Yes' if student else 'No'
            
            # Determine the Student Type (MD or PA)
            student_type = None
            if student:
                if '(MD)' in student:
                    student_type = 'MD'
                elif '(PA)' in student:
                    student_type = 'PA'

            # Append the data to the list
            processed_data.append({
                'Date': date,
                'Type': type_,
                'Description': name,
                'Preceptor': preceptor.strip(),
                'Student': student.strip() if student else None,
                'Student Placed': student_placed,
                'Student Type': student_type,  # New column for MD or PA
                'Location': location  # Add the Location column (sheet name)
            })
    
    return processed_data

# Streamlit app
st.title('OPD Data Processor')

# File uploader allows for multiple files to be uploaded
uploaded_files = st.file_uploader("Choose Excel files", type="xlsx", accept_multiple_files=True)

# If files are uploaded, process them
if uploaded_files:
    all_data = []
    
    # Process each uploaded file
    for uploaded_file in uploaded_files:
        # Read the Excel file
        wb = load_workbook(uploaded_file)
        
        # Process all sheets in the current file
        for sheet_name in wb.sheetnames:
            sheet = wb[sheet_name]
            
            # Process all the date columns (B, C, D, E, F, G, H) with their respective rows for each sheet
            for date_column in ['B', 'C', 'D', 'E', 'F', 'G', 'H']:  # Adjust this list if there are more date columns
                for date_row in [4, 28, 52, 76]:  # Process dates in rows 4, 28, 52, 76, etc.
                    start_row = date_row + 2  # AM starts at B6, C6, etc., and PM starts at B16, C16, etc.
                    
                    # Process the area and append the data to the list
                    area_data = process_area(sheet, date_column, date_row, start_row, sheet_name)
                    all_data.extend(area_data)

    # Convert the collected data into a DataFrame
    df = pd.DataFrame(all_data)

    # Display the DataFrame in the Streamlit app
    st.write(df)

    # Allow the user to download the processed data as an Excel file
    output_file = BytesIO()
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Processed Data')
    output_file.seek(0)
    
    # Provide a download link for the processed file
    st.download_button(
        label="Download Processed Data",
        data=output_file,
        file_name="processed_hope_drive_all_opd_files_with_location.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
