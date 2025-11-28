import streamlit as st
import pandas as pd
from io import BytesIO

# --- Configuration for Streamlit Page ---
st.set_page_config(
    page_title="EnergyAnalyser",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- Helper Function for Excel Column Conversion ---
def excel_col_to_index(col_str):
    """
    Converts an Excel column string (e.g., 'A', 'AA', 'BI') to a 0-based column index.
    Raises a ValueError if the string is invalid.
    """
    col_str = col_str.upper().strip()
    index = 0
    # A=1, B=2, ..., Z=26
    for char in col_str:
        if 'A' <= char <= 'Z':
            # Calculate the 1-based index (e.g., 'B' is 2, 'AA' is 27)
            index = index * 26 + (ord(char) - ord('A') + 1)
        else:
            raise ValueError(f"Invalid character in column string: {col_str}")
    
    # Convert 1-based index to 0-based index for Pandas (A=0, B=1)
    return index - 1

# --- App Title and Description ---
st.title("⚡ EnergyAnalyser: Data Consolidation")
st.markdown("""
    Upload your raw energy data CSV files (up to 10) to extract **Date**, **Time**, and **PSum** and consolidate them into a single Excel file.
    
    The application uses the third row (index 2) of the CSV file as the main column header.
""")

# --- Constants for Data Processing ---
PSUM_OUTPUT_NAME = 'PSum (W)' 

# Mapping user-friendly format to Python's datetime format strings
DATE_FORMAT_MAP = {
    "DD/MM/YYYY": "%d/%m/%Y %H:%M:%S",
    "YYYY-MM-DD": "%Y-%m-%d %H:%M:%S"
}

# --- User Configuration Section (Sidebar) ---
st.sidebar.header("⚙️ Column Configuration")
st.sidebar.markdown("Define the column letter for each data field.")

# Get user-defined column letters (A, B, C...)
date_col_str = st.sidebar.text_input(
    "Date Column Letter (Default: A)", 
    value='A', 
    key='date_col_str'
)

time_col_str = st.sidebar.text_input(
    "Time Column Letter (Default: B)", 
    value='B', 
    key='time_col_str'
)

ps_um_col_str = st.sidebar.text_input(
    "PSum Column Letter (Default: BI)", 
    value='BI', 
    key='psum_col_str',
    help="PSum (Total Active Power) is expected in Excel column BI. Adjust if needed."
)

# --- New Configuration for Data Reading ---
st.sidebar.header("📄 CSV File Settings")

# New input for CSV Delimiter
delimiter_input = st.sidebar.text_input(
    "CSV Delimiter (Separator)",
    value=',',
    key='delimiter_input',
    help="The character used to separate values in your CSV file (e.g., ',', ';', or '\\t' for tab)."
)

start_row_num = st.sidebar.number_input(
    "Data reading starts from (Row Number)",
    min_value=1,
    value=3, 
    key='start_row_num',
    help="The row number in the CSV file that contains the column headers (e.g., 'Date', 'Time', 'UA'). The default is Row 3." 
)

selected_date_format = st.sidebar.selectbox(
    "Date Format for Parsing",
    options=["DD/MM/YYYY", "YYYY-MM-DD"],
    index=0, # Default to DD/MM/YYYY
    key='selected_date_format',
    help="Select the format used for the Date column in your CSV file."
)

# --- Function to Process Data ---
def process_uploaded_files(uploaded_files, columns_config, header_index, date_format_string, separator):
    """
    Reads multiple CSV files, extracts configured columns, cleans PSum data, 
    and returns a dictionary of DataFrames.
    
    Accepts the 0-based header index, datetime format string, and the separator.
    """
    processed_data = {}
    
    # Ensure all required columns are unique
    if len(set(columns_config.keys())) != 3:
        st.error("Error: Date, Time, and PSum must be extracted from three unique column indices.")
        return {}

    col_indices = list(columns_config.keys())
    
    
    for uploaded_file in uploaded_files:
        filename = uploaded_file.name
        
        try:
            # 1. Read the CSV using the specified header row (header_index is 0-based)
            # --- USE USER-DEFINED SEPARATOR HERE ---
            df_full = pd.read_csv(
                uploaded_file, 
                header=header_index, 
                encoding='ISO-8859-1', 
                low_memory=False,
                sep=separator # Use the selected separator
            )
            
            # 2. Check if DataFrame has enough columns
            max_index = max(col_indices)
            if df_full.shape[1] < max_index + 1:
                 # Inform the user that the file likely didn't parse correctly (common delimiter issue)
                 st.error(f"File **{filename}** failed to read data correctly. It only has {df_full.shape[1]} columns. This usually means the **CSV Delimiter** ('{separator}') is incorrect. Please try changing the separator in the sidebar (e.g., to ';' or '\\t').")
                 continue

            # 3. Extract only the required columns by their index (iloc)
            df_extracted = df_full.iloc[:, col_indices].copy()
            
            # 4. Rename the columns to the final names for output
            temp_cols = {
                columns_config[k]: v for k, v in COLUMNS_TO_EXTRACT.items()
            }
            df_extracted.columns = temp_cols.values()
            
            # 5. Data Cleaning: Convert PSum to numeric, handling potential errors
            if PSUM_OUTPUT_NAME in df_extracted.columns:
                df_extracted[PSUM_OUTPUT_NAME] = pd.to_numeric(
                    df_extracted[PSUM_OUTPUT_NAME], 
                    errors='coerce' # Convert non-numeric values to NaN
                )

            # 6. Format Date and Time columns separately after parsing for correction
            # Combine the two raw columns ('Date' and 'Time') from the CSV for reliable datetime parsing.
            combined_dt_str = df_extracted['Date'].astype(str) + ' ' + df_extracted['Time'].astype(str)

            datetime_series = pd.to_datetime(
                combined_dt_str, 
                errors='coerce',
                # Use the selected date format string for parsing
                format=date_format_string 
            )
            
            # --- CHECK: Verify successful datetime parsing ---
            valid_dates_count = datetime_series.count()
            if valid_dates_count == 0:
                st.warning(f"File **{filename}**: No valid dates could be parsed. Check your 'Date Format for Parsing' setting (**{selected_date_format}**) and ensure the 'Date' and 'Time' columns ({date_col_str.upper()} and {time_col_str.upper()}) contain valid data starting from Row {start_row_num}.")
                continue
            # ---------------------------------------------------

            # GUARANTEE SEPARATION: Create a new DataFrame explicitly with separated columns
            df_final = pd.DataFrame({
                'Date': datetime_series.dt.strftime('%d/%m/%Y'), # Output Date is consistently DD/MM/YYYY
                'Time': datetime_series.dt.strftime('%H:%M:%S'),
                PSUM_OUTPUT_NAME: df_extracted[PSUM_OUTPUT_NAME] # Keep the PSum data from the original extracted DF
            })

            # 7. Clean the filename for the Excel sheet name
            sheet_name = filename.replace('.csv', '').replace('.', '_').strip()[:31]
            
            # Use the new, explicitly constructed DataFrame for the output
            processed_data[sheet_name] = df_final
            
        except Exception as e:
            # Catch all other unexpected exceptions
            st.error(f"Error processing file **{filename}**. An unexpected error occurred. Error: {e}")
            continue
            
    return processed_data


# --- Function to Generate Excel File for Download ---
@st.cache_data
def to_excel(data_dict):
    """
    Takes a dictionary of DataFrames and writes them to an in-memory Excel file.
    The index is NOT included (index=False).
    
    Explicitly sets column formats to text using xlsxwriter to prevent merging 
    of Date and Time columns by Excel.
    """
    output = BytesIO()
    # Use pandas ExcelWriter with 'xlsxwriter' engine
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet_name, df in data_dict.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

            # --- Explicitly set column formats to Text (Crucial Fix) ---
            
            # Get the xlsxwriter workbook and worksheet objects.
            workbook = writer.book
            worksheet = writer.sheets[sheet_name]
            
            # Define a text format (num_format: '@')
            text_format = workbook.add_format({'num_format': '@'})
            
            # Find column indices and apply the text format
            try:
                if 'Date' in df.columns:
                    date_col_index = df.columns.get_loc('Date')
                    # Apply text format to the entire column
                    worksheet.set_column(date_col_index, date_col_index, 12, text_format) 
                
                if 'Time' in df.columns:
                    time_col_index = df.columns.get_loc('Time')
                    # Apply text format to the entire column
                    worksheet.set_column(time_col_index, time_col_index, 10, text_format)
            except Exception as e:
                # Log any errors during explicit formatting but don't stop execution
                print(f"Error applying explicit xlsxwriter formats: {e}")
            # --------------------------------------------------------
            
    output.seek(0)
    return output.getvalue()


# --- Main Streamlit Logic ---
if __name__ == "__main__":
    
    # Try to convert column letters to 0-based indices
    try:
        date_col_index = excel_col_to_index(date_col_str)
        time_col_index = excel_col_to_index(time_col_str)
        ps_um_col_index = excel_col_to_index(ps_um_col_str)
        
        # Define the columns to extract
        COLUMNS_TO_EXTRACT = {
            date_col_index: 'Date',
            time_col_index: 'Time',
            ps_um_col_index: PSUM_OUTPUT_NAME
        }
        
    except ValueError as e:
        st.error(f"Configuration Error: Invalid column letter entered: {e}. Please use valid Excel column notation (e.g., A, C, AA).")
        st.stop()
    
    # 0-based index for Pandas header argument (Row Number - 1)
    header_row_index = int(start_row_num) - 1
    
    # Get the precise format string for parsing datetime objects
    date_format_string = DATE_FORMAT_MAP.get(selected_date_format)

    # File Uploader
    uploaded_files = st.file_uploader(
        "Choose up to 10 CSV files", 
        type=["csv"], 
        accept_multiple_files=True
    )
    
    # Limit to 10 files
    if uploaded_files and len(uploaded_files) > 10:
        st.warning(f"You have uploaded {len(uploaded_files)} files. Only the first 10 will be processed.")
        uploaded_files = uploaded_files[:10]

    # Processing and Download Button
    if uploaded_files:
        
        # Display the column letters and settings being used for user confirmation
        st.info(f"Processing {len(uploaded_files)} file(s) using columns: Date: {date_col_str.upper()}, Time: {time_col_str.upper()}, PSum: {ps_um_col_str.upper()}. Data reading starts from **Row {start_row_num}** using **{selected_date_format}** format and delimiter **'{delimiter_input}'**.")
        
        # 1. Process data 
        processed_data_dict = process_uploaded_files(
            uploaded_files, 
            COLUMNS_TO_EXTRACT, 
            header_row_index, 
            date_format_string,
            delimiter_input # Pass the new configuration
        )
        
        if processed_data_dict:
            
            # --- CONSOLIDATED RAW DATA SECTION ---
            st.header("Consolidated Raw Data Output")
            
            # Display a preview of the first processed file
            first_sheet_name = next(iter(processed_data_dict))
            st.subheader(f"Preview of: {first_sheet_name}")
            # Show a preview with the now separate Date and Time columns
            st.dataframe(processed_data_dict[first_sheet_name].head())
            st.success("Selected columns extracted and consolidated successfully!")

            # --- File Name Customization ---
            file_names_without_ext = [f.name.rsplit('.', 1)[0] for f in uploaded_files]
            
            if len(file_names_without_ext) > 1:
                first_name = file_names_without_ext[0]
                if len(first_name) > 20:
                    first_name = first_name[:17] + "..."
                default_filename = f"{first_name}_and_{len(file_names_without_ext) - 1}_More_Consolidated.xlsx"
            elif file_names_without_ext:
                default_filename = f"{file_names_without_ext[0]}_Consolidated.xlsx"
            else:
                default_filename = "EnergyAnalyser_Consolidated_Data.xlsx"


            custom_filename = st.text_input(
                "Output Excel Filename:",
                value=default_filename,
                key="output_filename_input_raw",
                help="Enter the name for the final Excel file with raw extracted data."
            )
            
            # Generate Excel file for raw data
            excel_data = to_excel(processed_data_dict)
            
            # Download Button for raw data
            st.download_button(
                label="📥 Download Consolidated Data (Date, Time, PSum)",
                data=excel_data,
                file_name=custom_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Click to download the Excel file with one sheet per uploaded CSV file."
            )

        else:
            st.error("No data could be successfully processed. Please review the error messages above and adjust the column letters if necessary.")
