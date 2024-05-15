import pandas as pd
import streamlit as st
import time
from io import BytesIO
import xlwings as xw
from pathlib import Path

def compare_excel_files(df1, df2):
    key_column = 'key_column'  # Change this to the actual key column name if different

    # Merge DataFrames on the key column to compare rows with the same value in key_column
    merged_df = pd.merge(df1, df2, on=key_column, how='outer', suffixes=('_file1', '_file2'), indicator=True)
    
    # Discrepancies in existing rows
    discrepancies = merged_df[merged_df['_merge'] == 'both'].copy()
    discrepancies = discrepancies.loc[:, discrepancies.columns != '_merge']
    discrepancies = discrepancies.loc[
        discrepancies.filter(like='_file1').ne(discrepancies.filter(like='_file2')).any(axis=1)
    ]
    
    # Rows only in file1
    only_in_file1 = merged_df[merged_df['_merge'] == 'left_only'].copy()
    only_in_file1 = only_in_file1.loc[:, only_in_file1.columns != '_merge']
    
    # Rows only in file2
    only_in_file2 = merged_df[merged_df['_merge'] == 'right_only'].copy()
    only_in_file2 = only_in_file2.loc[:, only_in_file2.columns != '_merge']

    # Create a new Excel file with three sheets in memory
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        discrepancies.to_excel(writer, sheet_name='Discrepancies', index=False)
        only_in_file1.to_excel(writer, sheet_name='Only in file1', index=False)
        only_in_file2.to_excel(writer, sheet_name='Only in file2', index=False)

    return output.getvalue()

def highlight_discrepancies(file1_path, file2_path):
    try:
        # Open workbooks using xlwings
        with xw.App(visible=False) as app:
            initial_wb = app.books.open(file1_path)
            initial_ws = initial_wb.sheets[1]

            updated_wb = app.books.open(file2_path)
            updated_ws = updated_wb.sheets[1]

            for cell in updated_ws.used_range:
                old_value = initial_ws.range((cell.row, cell.column)).value
                if cell.value != old_value:
                    cell.api.AddComment(f"Value from {initial_wb.name}: {old_value}")  # WARNING: Platform specific (!)
                    cell.color = (255, 71, 76)  # light red

            highlighted_path = Path.cwd() / "Difference_Highlighted.xlsx"
            updated_wb.save(highlighted_path)
        return highlighted_path
    except Exception as e:
        st.error(f"Error occurred while highlighting discrepancies: {e}")
        return None

st.title("Excel File Comparison")
st.write("Instructions: 1. Select files to be compared (must be named 'file1.xlsx' and 'file2.xlsx'). 2. Once files are selected and loaded, engage the 'Analyze Data' button. 3. The app will begin to run. Once complete, select the 'Comparison Report' to download the report.")

# Add a placeholder for the loading bar
latest_iteration = st.empty()
bar = st.progress(0)

# File upload widgets
uploaded_file1 = st.file_uploader("Upload the first Excel file", type=["xlsx"], key="file1")
uploaded_file2 = st.file_uploader("Upload the second Excel file", type=["xlsx"], key="file2")

# Analyze Data button
if uploaded_file1 and uploaded_file2:
    if st.button("Analyze Data"):
        for i in range(100):
            # Update the progress bar with each iteration.
            latest_iteration.text(f'Iteration {i+1}')
            bar.progress(i + 1)
            time.sleep(0.01)

        # Save the uploaded files
        file1_path = Path("uploaded_file1.xlsx")
        file2_path = Path("uploaded_file2.xlsx")
        with open(file1_path, "wb") as f:
            f.write(uploaded_file1.getbuffer())
        with open(file2_path, "wb") as f:
            f.write(uploaded_file2.getbuffer())

        # Read the uploaded files into DataFrames
        df1 = pd.read_excel(file1_path)
        df2 = pd.read_excel(file2_path)
        
        # Generate the comparison report
        report = compare_excel_files(df1, df2)
        
        # Provide download links for the report and highlighted file
        st.download_button(
            label="Download Comparison Report",
            data=report,
            file_name="comparison_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        # Highlight discrepancies in the original files (if possible)
        highlighted_path = highlight_discrepancies(file1_path, file2_path)
        
        if highlighted_path:
            with open(highlighted_path, "rb") as f:
                st.download_button(
                    label="Download Highlighted Differences",
                    data=f,
                    file_name="Difference_Highlighted.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.write("Could not generate highlighted differences.")


