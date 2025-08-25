import streamlit as st
import pandas as pd
import io
from datetime import datetime

# Set page configuration
st.title("Multi XLSX File Combiner")
st.write("Upload multiple XLSX files. Each file should have a 'Timestamp' column followed by meter columns. The app will combine them into one Excel file with a single 'Timestamp' column and all unique meters.")

# File uploader for multiple files
uploaded_files = st.file_uploader("Upload XLSX Files", type=["xlsx", "xls"], accept_multiple_files=True)

if uploaded_files:
    try:
        dfs = []
        for uploaded_file in uploaded_files:
            # Read the Excel file
            df = pd.read_excel(uploaded_file)
            
            # Ensure 'Timestamp' is the first column
            if df.columns[0] != 'Timestamp':
                st.warning(f"File {uploaded_file.name} does not have 'Timestamp' as the first column. Skipping.")
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
            
            # Rename meter columns: remove " - Consumption Recorded (MWh)"
            new_columns = ['Timestamp'] + [col.replace(" - Consumption Recorded (MWh)", "") for col in df.columns[1:]]
            df.columns = new_columns
            
            # Set index to Timestamp
            df = df.set_index('Timestamp')
            
            dfs.append(df)
        
        if not dfs:
            st.error("No valid files processed.")
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
    except Exception as e:
        st.error(f"An error occurred while processing the files: {str(e)}")
else:
    st.info("Please upload at least one XLSX file.")