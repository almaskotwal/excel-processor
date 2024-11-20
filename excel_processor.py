import streamlit as st
import openpyxl
import io

def process_excel(uploaded_input_file, uploaded_template_file):
    """Processes an Excel file, populates templates, and saves new files.

    Args:
        uploaded_input_file: Uploaded input file object (Streamlit.UploadedFile).
        uploaded_template_file: Path to the template Excel file.
    """

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

        # Create output workbook in memory
        output_workbook = openpyxl.Workbook()
        output_worksheet = output_workbook.active

        # Process data and populate output worksheet
        for row_num, row in enumerate(input_worksheet.iter_rows(min_row=3, values_only=True), 3):
            if row:
                # Extract data from input row
                trip_id = row[target_columns["Trip ID"] - 1]
                driver_name = row[target_columns["Driver Name"] - 1]
                facility_sequence = row[target_columns["Facility Sequence"] - 1]
                estimated_cost = row[target_columns["Estimated Cost"] - 1]

                # Populate output worksheet
                output_worksheet['B11'] = trip_id
                output_worksheet['D4'] = driver_name
                output_worksheet['B13'] = facility_sequence
                output_worksheet['B14'] = estimated_cost

        # Create a temporary file object in memory for download
        output_data = io.BytesIO()
        output_workbook.save(output_data)
        output_data.seek(0)

        # Download the processed file
        st.download_button(label="Download Processed File", data=output_data, file_name="processed_output.xlsx")

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