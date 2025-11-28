import streamlit as st
import pandas as pd
import numpy as np

# --- Configuration ---
# The raw data files have two header rows. The second row (index 1) contains the usable column names.
HEADER_ROW_INDEX = 1 
# We select columns by name for robustness. These columns are chosen based on
# common requirements for energy analysis and the structure of your raw data snippets.
REQUIRED_COLUMNS = ['Time', 'Total Active Power Demand(W)']

def process_msb_data(file_path, msb_name):
    """
    Reads and processes the MSB raw data files.
    This function is corrected to use column names and the appropriate header index.
    """
    st.write(f"--- Processing {msb_name} Data ---")
    
    try:
        # Use header=HEADER_ROW_INDEX (1) to correctly set the column names 
        # based on the row that starts with 'Date,Time,UA,UB...'.
        # We explicitly skip the ProductSN row (Row 0, index 0).
        df_raw = pd.read_csv(file_path, header=HEADER_ROW_INDEX)
        
        # --- Column Validation and Selection ---
        
        # Check if all required columns are present in the loaded DataFrame
        missing_cols = [col for col in REQUIRED_COLUMNS if col not in df_raw.columns]
        if missing_cols:
            st.error(f"Error for {msb_name}: Missing expected columns: {missing_cols}. Please check the file structure.")
            st.warning(f"Available Columns: {list(df_raw.columns)[:20]}...")
            return None

        # Select the necessary columns by name (not index)
        df = df_raw[REQUIRED_COLUMNS].copy()
        
        # Rename the Power column to be consistent across charts
        df.rename(columns={'Total Active Power Demand(W)': 'Power (W)', 'Time': 'Timestamp'}, inplace=True)
        
        # Convert power from W to kW
        df['Power (kW)'] = df['Power (W)'] / 1000.0
        
        st.success(f"Successfully loaded and processed {msb_name}. Rows: {len(df)}")
        st.dataframe(df.head())
        
        return df

    except FileNotFoundError:
        st.error(f"File not found for {msb_name} at path: {file_path}")
        return None
    except Exception as e:
        # Catch any other unexpected error and report it clearly
        st.error(f"An unexpected error occurred while processing {msb_name}: {e}")
        return None

def main():
    """Main Streamlit application logic."""
    st.set_page_config(layout="wide", page_title="Energy Data Analyser (MSB)")
    st.title("Energy Data Analyser (MSB)")
    st.markdown("Load and analyze raw energy data from MSB devices.")

    # --- File Paths (Update these based on your deployment location) ---
    msb_files = {
        'MSB-1': 'raw data MSB 1.csv',
        'MSB-2': 'raw data MSB 2.csv',
        'MSB-3': 'raw data MSB 3.csv',
    }
    
    # --- Data Processing and Visualization ---
    
    all_dfs = {}
    
    for msb_name, file_path in msb_files.items():
        # NOTE: Using the accessible filenames from the environment
        if msb_name == 'MSB-1':
            file_path = 'raw data MSB 1.csv'
        elif msb_name == 'MSB-2':
            file_path = 'raw data MSB 2.csv'
        elif msb_name == 'MSB-3':
            file_path = 'raw data MSB 3.csv'
            
        df = process_msb_data(file_path, msb_name)
        if df is not None:
            all_dfs[msb_name] = df
            
    if all_dfs:
        st.header("Combined Daily Power Overview (kW)")

        # Simple concatenation for comparison
        combined_data = pd.DataFrame()
        
        # Plotting the data if processing was successful
        for msb_name, df in all_dfs.items():
            if 'Power (kW)' in df.columns and 'Timestamp' in df.columns:
                # Group by Timestamp (assuming the raw data is recorded at specific intervals)
                # and plot the average power for those intervals
                df_plot = df.set_index('Timestamp')['Power (kW)'].rename(msb_name)
                combined_data = pd.concat([combined_data, df_plot], axis=1)

        if not combined_data.empty:
            # Drop rows with any NaN values that could result from inconsistent timestamps
            combined_data = combined_data.dropna()
            
            st.line_chart(combined_data)
            st.subheader("Data Summary (First 500 rows only)")
            st.dataframe(combined_data.head(500))
        else:
            st.info("No data available to combine and display.")
            
if __name__ == "__main__":
    main()
