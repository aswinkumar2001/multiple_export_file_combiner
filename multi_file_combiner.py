import streamlit as st
import pandas as pd
import io
from datetime import datetime
import zipfile
from io import BytesIO
import os

# Set page configuration
st.title("ZIP File Combiner for CSV Files")
st.write("Upload a ZIP file containing multiple CSV files. Each CSV should have a 'Timestamp' column followed by meter columns. The app will extract, process, and combine them into one Excel file with a single 'Timestamp' column and all unique meters.")

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
                for csv_name in csv_files:
                    try:
                        # Read CSV from ZIP with latin1 encoding to handle special characters
                        with zip_ref.open(csv_name) as csv_file:
                            df = pd.read_csv(csv_file, encoding='latin1')
                        
                        # Ensure 'Timestamp' is the first column
                        if df.columns[0] != 'Timestamp':
                            st.warning(f"File {csv_name} does not have 'Timestamp' as the first column. Skipping.")
                            continue
                        
                        # Parse Timestamp to datetime
                        def parse_timestamp(ts):
                            try:
                                return datetime.strptime(ts, "%A, %B %d, %Y %H:%M")
                            except ValueError:
                                return None
                        
                        df['Timestamp'] = df['Timestamp'].apply(parse_timestamp)
                        
                        # Drop rows with invalid timestamps
                        df = df.dropna(subset=['Timestamp'])
                        
                        if df.empty:
                            st.warning(f"File {csv_name} has no valid timestamps. Skipping.")
                            continue
                        
                        # Rename meter columns: remove " - Consumption Recorded (MWh)"
                        new_columns = ['Timestamp'] + [col.replace(" - Consumption Recorded (MWh)", "") for col in df.columns[1:]]
                        df.columns = new_columns
                        
                        # Set index to Timestamp
                        df = df.set_index('Timestamp')
                        
                        dfs.append(df)
                    except UnicodeDecodeError as ude:
                        st.warning(f"Encoding error in {csv_name}: {str(ude)}. Try checking file encoding. Skipping.")
                    except pd.errors.ParserError as pe:
                        st.warning(f"Parsing error in {csv_name}: {str(pe)}. File may not be a valid CSV. Skipping.")
                    except Exception as e:
                        st.warning(f"Error processing {csv_name}: {str(e)}. Skipping.")
                
                if not dfs:
                    st.error("No valid CSV files processed.")
                else:
                    # Concatenate all DataFrames on columns (axis=1), aligning on Timestamp index
                    combined_df = pd.concat(dfs, axis=1)
                    
                    # Reset index to bring Timestamp back as column
                    combined_df = combined_df.reset_index()
                    
                    # Display preview
                    st.write("Preview of Combined Data:")
                    st.dataframe(combined_df.head(10))  # Show first 10 rows
                    
                    # Function to convert DataFrame to Excel
                    def to_excel(df):
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='xlsxwriter', datetime_format='YYYY-MM-DD HH:MM') as writer:
                            df.to_excel(writer, index=False, sheet_name='Combined_Data')
                        return output.getvalue()
                    
                    # Download button
                    excel_data = to_excel(combined_df)
                    st.download_button(
                        label="Download Combined Excel",
                        data=excel_data,
                        file_name="combined_data.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
    except zipfile.BadZipFile:
        st.error("The uploaded file is not a valid ZIP file.")
    except Exception as e:
        st.error(f"An unexpected error occurred: {str(e)}")
else:
    st.info("Please upload a ZIP file containing CSV files.")
