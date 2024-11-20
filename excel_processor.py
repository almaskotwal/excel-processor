import streamlit as st
import openpyxl
import io
from zipfile import ZipFile
from io import BytesIO

def process_excel(uploaded_input_file, uploaded_template_file):
    try:
        # Load input and template workbooks
        input_workbook = openpyxl.load_workbook(uploaded_input_file)
        input_worksheet = input_workbook.active
        template_workbook = openpyxl.load_workbook(uploaded_template_file)
        template_worksheet = template_workbook.active

        # Define target columns
        target_columns = {"Trip ID": None, "Driver Name": None, "Facility Sequence": None, "Estimated Cost": None}

        # Find column indices
        for header_cell in input_worksheet.iter_rows(min_row=2, values_only=True):
            if header_cell:
                for col_idx, value in enumerate(header_cell):
                    if value is not None:
                        value_str = str(value).strip()
                        if value_str in target_columns:
                            target_columns[value_str] = col_idx + 1

        # Check if all target columns were found
        if not all(value for value in target_columns.values()):
            raise ValueError(f"Missing target columns in input file: {', '.join(missing for missing in target_columns if not target_columns[missing])}")

        # Process data and create output files
        driver_to_workbook = {}

        for row_num, row in enumerate(input_worksheet.iter_rows(min_row=3, values_only=True), 3):
            if row:
                # Extract data from input row
                trip_id = row[target_columns["Trip ID"] - 1]
                driver_name = row[target_columns["Driver Name"] - 1]
                facility_sequence = row[target_columns["Facility Sequence"] - 1]
                estimated_cost = row[target_columns["Estimated Cost"] - 1]

                # Get the workbook for the current driver
                workbook = driver_to_workbook.setdefault(driver_name, openpyxl.Workbook())
                worksheet = workbook.active

                # Copy formatting from template to output worksheet
                for row in template_worksheet.iter_rows():
                    for cell in row:
                        output_worksheet.cell(row=cell.row, column=cell.column).style = cell.style

                # Populate output worksheet with data
                worksheet['B11'] = trip_id
                worksheet['D4'] = driver_name
                worksheet['B13'] = facility_sequence
                worksheet['B14'] = estimated_cost

        # Create a zip file to download all output files
        zip_buffer = BytesIO()
        with ZipFile(zip_buffer, 'w') as zipf:
            for driver_name, workbook in driver_to_workbook.items():
                filename = f"{driver_name}.xlsx"
                with io.BytesIO() as file_buffer:
                    workbook.save(file_buffer)
                    file_buffer.seek(0)
                    zipf.writestr(filename, file_buffer.read())

        zip_buffer.seek(0)
        st.download_button(label="Download All Files", data=zip_buffer, file_name="output_files.zip")

    except FileNotFoundError:
        st.error(f"Failed to open input file.")
    except Exception as e:
        st.error(f"An error occurred: {str(e)}")

def main():
    st.title("Excel Processor")

    uploaded_input_file = st.file_uploader("Upload Input File", type=["xlsx"])
    uploaded_template_file = st.file_uploader("Upload Template File", type=["xlsx"])

    if st.button("Process"):
        if uploaded_input_file is not None and uploaded_template_file is not None:
            process_excel(uploaded_input_file, uploaded_template_file)
        else:
            st.warning("Please upload both input and template files.")

if __name__ == "__main__":
    main()