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

    # Display the combined dataset
    st.write("Filtered Combined Dataset (Only 'Student Placed' = Yes):")
    st.write(filtered_df)

    # Display the summary table of total days worked
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

    # Allow download of the combined dataset, daily breakdown, and summary
    output_file = BytesIO()
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        filtered_df.to_excel(writer, index=False, sheet_name='Combined Dataset')
        days_worked.to_excel(writer, index=False, sheet_name='Days Worked Detail')
        preceptor_days_summary.to_excel(writer, index=False, sheet_name='Total Days Summary')
    output_file.seek(0)

    st.download_button(
        label="Download Combined and Summary Data",
        data=output_file,
        file_name="combined_and_summary_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

