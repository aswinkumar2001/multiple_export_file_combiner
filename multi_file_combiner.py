import streamlit as st
import pandas as pd
import io
from datetime import datetime
import zipfile
from io import BytesIO
import os
import numpy as np

# Set page configuration
st.set_page_config(page_title="ZIP File Combiner", layout="wide")
st.title("ZIP File Combiner for CSV Files")
st.write("""
Upload a ZIP file containing multiple CSV files. Each CSV should have a 'Timestamp' column followed by meter columns. 
The app will extract, process, and combine them into one Excel file with a single 'Timestamp' column and all unique meters.
""")

# File uploader for ZIP file
uploaded_zip = st.file_uploader("Upload ZIP File", type=["zip"])

if uploaded_zip:
    try:
        # Read the ZIP file
        zip_content = BytesIO(uploaded_zip.read())
        with zipfile.ZipFile(zip_content, 'r') as zip_ref:
            # Get list of CSV files in ZIP, skipping macOS metadata files
            csv_files = [
                f for f in zip_ref.namelist()
                if f.lower().endswith('.csv')
                and 'MACOSX' not in f.upper()
                and not os.path.basename(f).startswith('.')
            ]
            
            if not csv_files:
                st.error("No valid CSV files found in the uploaded ZIP. Note: macOS metadata files are skipped.")
            else:
                dfs = []
                file_info = []  # To store information about each file
                
                # Add a progress bar for processing files
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                for i, csv_name in enumerate(csv_files):
                    status_text.text(f"Processing {i+1}/{len(csv_files)}: {os.path.basename(csv_name)}")
                    progress_bar.progress((i + 1) / len(csv_files))
                    
                    try:
                        # Try multiple encodings to handle different CSV formats
                        for encoding in ['utf-8', 'latin1', 'cp1252', 'iso-8859-1']:
                            try:
                                with zip_ref.open(csv_name) as csv_file:
                                    df = pd.read_csv(csv_file, encoding=encoding)
                                break
                            except UnicodeDecodeError:
                                continue
                        else:
                            st.warning(f"Could not read {csv_name} with any supported encoding. Skipping.")
                            continue
                        
                        # Ensure 'Timestamp' is the first column (case-insensitive check)
                        timestamp_col = None
                        for col in df.columns:
                            if 'timestamp' in col.lower():
                                timestamp_col = col
                                break
                        
                        if timestamp_col is None:
                            st.warning(f"File {csv_name} does not have a 'Timestamp' column. Skipping.")
                            continue
                        
                        # Make sure Timestamp is the first column
                        cols = df.columns.tolist()
                        cols.remove(timestamp_col)
                        cols.insert(0, timestamp_col)
                        df = df[cols]
                        
                        # Parse Timestamp to datetime with multiple format support
                        def parse_timestamp(ts):
                            if pd.isna(ts):
                                return None
                                
                            ts_str = str(ts).strip()
                            formats = [
                                "%A, %B %d, %Y %H:%M",  # Original format: Monday, January 01, 2024 00:00
                                "%Y-%m-%d %H:%M:%S",     # ISO format
                                "%m/%d/%Y %H:%M",        # US format
                                "%d/%m/%Y %H:%M",        # European format
                                "%Y-%m-%d %H:%M",        # ISO without seconds
                                "%d-%m-%Y %H:%M",        # Another common format
                            ]
                            
                            for fmt in formats:
                                try:
                                    return datetime.strptime(ts_str, fmt)
                                except ValueError:
                                    continue
                            return None
                        
                        df[timestamp_col] = df[timestamp_col].apply(parse_timestamp)
                        
                        # Drop rows with invalid timestamps
                        initial_count = len(df)
                        df = df.dropna(subset=[timestamp_col])
                        if len(df) < initial_count:
                            st.info(f"File {csv_name}: Dropped {initial_count - len(df)} rows with invalid timestamps.")
                        
                        if df.empty:
                            st.warning(f"File {csv_name} has no valid timestamps. Skipping.")
                            continue
                        
                        # Rename meter columns: remove " - Consumption Recorded (MWh)"
                        new_columns = [timestamp_col] + [col.replace(" - Consumption Recorded (MWh)", "") for col in df.columns[1:]]
                        df.columns = new_columns
                        
                        # Set index to Timestamp
                        df = df.set_index(timestamp_col)
                        
                        # Sort by timestamp
                        df = df.sort_index()
                        
                        # Store file information
                        time_range = f"{df.index.min().strftime('%Y-%m-%d')} to {df.index.max().strftime('%Y-%m-%d')}"
                        file_info.append({
                            'filename': os.path.basename(csv_name),
                            'time_range': time_range,
                            'meters': list(df.columns),
                            'rows': len(df)
                        })
                        
                        dfs.append(df)
                        
                    except pd.errors.ParserError as pe:
                        st.warning(f"Parsing error in {csv_name}: {str(pe)}. File may not be a valid CSV. Skipping.")
                    except Exception as e:
                        st.warning(f"Error processing {csv_name}: {str(e)}. Skipping.")
                
                if not dfs:
                    st.error("No valid CSV files processed.")
                else:
                    status_text.text("Combining data from all files...")
                    
                    # Combine all DataFrames along the time axis (vertically)
                    # First, ensure all DataFrames have the same column names
                    all_columns = set()
                    for df in dfs:
                        all_columns.update(df.columns)
                    
                    all_columns = sorted(list(all_columns))
                    
                    # Reindex each DataFrame to have all columns
                    reindexed_dfs = []
                    for df in dfs:
                        # Add missing columns with NaN values
                        for col in all_columns:
                            if col not in df.columns:
                                df[col] = np.nan
                        # Reorder columns to match the master list
                        df = df[all_columns]
                        reindexed_dfs.append(df)
                    
                    # Concatenate along the time axis (vertically)
                    combined_df = pd.concat(reindexed_dfs, axis=0)
                    
                    # Sort by timestamp
                    combined_df = combined_df.sort_index()
                    
                    # Remove duplicate timestamps (keep last occurrence)
                    combined_df = combined_df[~combined_df.index.duplicated(keep='last')]
                    
                    # Reset index to bring Timestamp back as column
                    combined_df = combined_df.reset_index()
                    combined_df.rename(columns={'index': 'Timestamp'}, inplace=True)
                    
                    # Display file information
                    st.subheader("Processed Files Summary")
                    file_info_df = pd.DataFrame(file_info)
                    st.dataframe(file_info_df)
                    
                    # Display combined data information
                    st.subheader("Combined Data Information")
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Total Time Range", 
                                 f"{combined_df['Timestamp'].min().strftime('%Y-%m-%d')} to {combined_df['Timestamp'].max().strftime('%Y-%m-%d')}")
                    with col2:
                        st.metric("Number of Meters", len(combined_df.columns) - 1)
                    with col3:
                        st.metric("Total Rows", len(combined_df))
                    
                    # Display preview
                    st.subheader("Preview of Combined Data")
                    st.dataframe(combined_df.head(10))  # Show first 10 rows
                    
                    # Show data completeness
                    st.subheader("Data Completeness by Meter")
                    completeness = {}
                    for col in combined_df.columns[1:]:  # Skip Timestamp column
                        completeness[col] = {
                            'Total Values': len(combined_df),
                            'Non-Null Values': combined_df[col].notna().sum(),
                            'Completeness %': round(combined_df[col].notna().sum() / len(combined_df) * 100, 2)
                        }
                    
                    completeness_df = pd.DataFrame(completeness).T
                    st.dataframe(completeness_df)
                    
                    # Function to convert DataFrame to Excel
                    def to_excel(df):
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='xlsxwriter', datetime_format='YYYY-MM-DD HH:MM') as writer:
                            df.to_excel(writer, index=False, sheet_name='Combined_Data')
                            
                            # Add summary sheet
                            file_info_df.to_excel(writer, index=False, sheet_name='File_Summary')
                            completeness_df.to_excel(writer, sheet_name='Data_Completeness')
                            
                            # Get workbook and worksheet objects
                            workbook = writer.book
                            worksheet = writer.sheets['Combined_Data']
                            
                            # Add a header format
                            header_format = workbook.add_format({
                                'bold': True,
                                'text_wrap': True,
                                'valign': 'top',
                                'fg_color': '#D7E4BC',
                                'border': 1
                            })
                            
                            # Write the column headers with the defined format
                            for col_num, value in enumerate(df.columns.values):
                                worksheet.write(0, col_num, value, header_format)
                                
                            # Adjust column widths
                            for i, col in enumerate(df.columns):
                                max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
                                worksheet.set_column(i, i, min(max_len, 50))
                        return output.getvalue()
                    
                    # Download button
                    excel_data = to_excel(combined_df)
                    st.download_button(
                        label="Download Combined Excel File",
                        data=excel_data,
                        file_name="combined_meter_data.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        help="Download the combined data as an Excel file with multiple sheets"
                    )
                    
                    status_text.text("Processing complete!")
                    progress_bar.empty()
                    
    except zipfile.BadZipFile:
        st.error("The uploaded file is not a valid ZIP file.")
    except Exception as e:
        st.error(f"An unexpected error occurred: {str(e)}")
else:
    st.info("Please upload a ZIP file containing CSV files.")
