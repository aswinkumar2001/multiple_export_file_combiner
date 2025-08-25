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
                all_data = []  # Store all data as dictionaries to preserve exact values
                file_info = []  # To store information about each file
                all_meters = set()  # Track all unique meter names
                timestamp_format = "%A, %B %d, %Y %H:%M"  # Thursday, January 02, 2025 02:30
                
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
                                    # Read without parsing dates to preserve exact values
                                    df = pd.read_csv(csv_file, encoding=encoding, dtype=str)
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
                        
                        # Parse Timestamp to datetime with the specific format
                        parsed_timestamps = []
                        invalid_timestamps = []
                        
                        for ts in df[timestamp_col]:
                            try:
                                # Try the specific format first
                                parsed_ts = datetime.strptime(ts, timestamp_format)
                                parsed_timestamps.append(parsed_ts)
                            except ValueError:
                                # If the specific format fails, try other common formats
                                try:
                                    # Try without zero-padded day
                                    parsed_ts = datetime.strptime(ts, "%A, %B %d, %Y %H:%M")
                                    parsed_timestamps.append(parsed_ts)
                                except ValueError:
                                    invalid_timestamps.append(ts)
                                    parsed_timestamps.append(None)
                        
                        # Add parsed timestamps to dataframe
                        df['Parsed_Timestamp'] = parsed_timestamps
                        
                        # Drop rows with invalid timestamps
                        initial_count = len(df)
                        df = df[df['Parsed_Timestamp'].notna()]
                        if len(df) < initial_count:
                            st.warning(f"File {csv_name}: Dropped {initial_count - len(df)} rows with invalid timestamps.")
                            if invalid_timestamps:
                                st.write(f"Invalid timestamp examples: {invalid_timestamps[:3]}")
                        
                        if df.empty:
                            st.warning(f"File {csv_name} has no valid timestamps. Skipping.")
                            continue
                        
                        # Rename meter columns: remove " - Consumption Recorded (MWh)"
                        new_columns = ['Parsed_Timestamp'] + [col.replace(" - Consumption Recorded (MWh)", "") for col in df.columns[1:-1]]
                        df.columns = new_columns
                        
                        # Store all meter names
                        meter_cols = list(df.columns[1:])
                        all_meters.update(meter_cols)
                        
                        # Convert numeric columns to appropriate types, preserving zeros
                        for col in meter_cols:
                            # Convert to numeric, but preserve all values including zeros
                            df[col] = pd.to_numeric(df[col], errors='coerce')
                            # Replace NaN with 0 only if the entire column is numeric
                            if df[col].dtype in [np.float64, np.int64]:
                                df[col] = df[col].fillna(0)
                        
                        # Store file information
                        time_range = f"{df['Parsed_Timestamp'].min().strftime('%Y-%m-%d')} to {df['Parsed_Timestamp'].max().strftime('%Y-%m-%d')}"
                        file_info.append({
                            'filename': os.path.basename(csv_name),
                            'time_range': time_range,
                            'meters': meter_cols,
                            'rows': len(df)
                        })
                        
                        # Convert to list of dictionaries to preserve exact values
                        for _, row in df.iterrows():
                            record = {'Timestamp': row['Parsed_Timestamp']}
                            for meter in meter_cols:
                                record[meter] = row[meter]
                            all_data.append(record)
                        
                    except pd.errors.ParserError as pe:
                        st.warning(f"Parsing error in {csv_name}: {str(pe)}. File may not be a valid CSV. Skipping.")
                    except Exception as e:
                        st.warning(f"Error processing {csv_name}: {str(e)}. Skipping.")
                        import traceback
                        st.write(traceback.format_exc())
                
                if not all_data:
                    st.error("No valid CSV files processed.")
                else:
                    status_text.text("Combining data from all files...")
                    
                    # Create final dataframe with all meters as columns
                    combined_df = pd.DataFrame(all_data)
                    
                    # Ensure all meter columns are present (fill with 0 if missing from some files)
                    for meter in all_meters:
                        if meter not in combined_df.columns:
                            combined_df[meter] = 0
                    
                    # Sort by timestamp
                    combined_df = combined_df.sort_values(by='Timestamp')
                    
                    # Check for duplicate timestamps
                    duplicate_count = combined_df.duplicated(subset=['Timestamp']).sum()
                    if duplicate_count > 0:
                        st.warning(f"Found {duplicate_count} duplicate timestamps. Keeping all values.")
                        # For duplicate timestamps, we'll keep all rows but show a warning
                    
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
                    st.subheader("Preview of Combined Data (First 10 Rows)")
                    st.dataframe(combined_df.head(10))
                    
                    # Display last few rows to verify data integrity
                    st.subheader("End of Combined Data (Last 10 Rows)")
                    st.dataframe(combined_df.tail(10))
                    
                    # Show data completeness
                    st.subheader("Data Completeness by Meter")
                    completeness = {}
                    for col in combined_df.columns[1:]:  # Skip Timestamp column
                        completeness[col] = {
                            'Total Values': len(combined_df),
                            'Zero Values': (combined_df[col] == 0).sum(),
                            'Non-Zero Values': (combined_df[col] != 0).sum(),
                            'Non-Zero %': round((combined_df[col] != 0).sum() / len(combined_df) * 100, 2)
                        }
                    
                    completeness_df = pd.DataFrame(completeness).T
                    st.dataframe(completeness_df)
                    
                    # Show timestamp frequency analysis
                    st.subheader("Timestamp Frequency Analysis")
                    time_diffs = combined_df['Timestamp'].diff().dropna()
                    if not time_diffs.empty:
                        most_common_diff = time_diffs.mode().iloc[0] if not time_diffs.mode().empty else None
                        st.write(f"Most common time interval: {most_common_diff}")
                        st.write(f"Time range: {combined_df['Timestamp'].min()} to {combined_df['Timestamp'].max()}")
                    
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
        import traceback
        st.write(traceback.format_exc())
else:
    st.info("Please upload a ZIP file containing CSV files.")
