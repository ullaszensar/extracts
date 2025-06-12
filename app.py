import streamlit as st
import pandas as pd
import io
from data_processor import DataProcessor
from utils import validate_excel_file, create_download_link

def main():
    st.title("Excel Data Processing Application")
    st.markdown("Upload two Excel files to extract demographic data and merge with table information")
    
    # Initialize session state
    if 'processed_data' not in st.session_state:
        st.session_state.processed_data = None
    if 'processing_complete' not in st.session_state:
        st.session_state.processing_complete = False
    
    # Create two columns for file uploads
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📊 Table Data File")
        st.markdown("Upload Excel file containing table information with storage_id")
        table_file = st.file_uploader(
            "Choose table data Excel file",
            type=['xlsx', 'xls'],
            key="table_file",
            help="This file should contain table names and their corresponding storage_id values"
        )
        
        if table_file:
            if validate_excel_file(table_file):
                st.success("✅ Valid Excel file uploaded")
                # Preview table data
                try:
                    df_preview = pd.read_excel(table_file, nrows=5)
                    st.write("**Preview (first 5 rows):**")
                    st.dataframe(df_preview)
                except Exception as e:
                    st.error(f"Error reading file: {str(e)}")
            else:
                st.error("❌ Invalid file format. Please upload an Excel file (.xlsx or .xls)")
    
    with col2:
        st.subheader("📋 Columns Data File")
        st.markdown("Upload Excel file containing demographic data with storage_id")
        columns_file = st.file_uploader(
            "Choose columns data Excel file",
            type=['xlsx', 'xls'],
            key="columns_file",
            help="This file should contain demographic information and storage_id for matching"
        )
        
        if columns_file:
            if validate_excel_file(columns_file):
                st.success("✅ Valid Excel file uploaded")
                # Preview columns data
                try:
                    df_preview = pd.read_excel(columns_file, nrows=5)
                    st.write("**Preview (first 5 rows):**")
                    st.dataframe(df_preview)
                except Exception as e:
                    st.error(f"Error reading file: {str(e)}")
            else:
                st.error("❌ Invalid file format. Please upload an Excel file (.xlsx or .xls)")
    
    # Processing section
    if table_file and columns_file:
        st.divider()
        
        # Configuration section
        st.subheader("⚙️ Processing Configuration")
        
        col_config1, col_config2 = st.columns(2)
        
        with col_config1:
            storage_id_col_table = st.text_input(
                "Storage ID column name in table data",
                value="storage_id",
                help="The column name that contains storage_id in the table data file"
            )
            
            table_name_col = st.text_input(
                "Table name column in table data",
                value="table_name",
                help="The column name that contains table names in the table data file"
            )
        
        with col_config2:
            storage_id_col_columns = st.text_input(
                "Storage ID column name in columns data",
                value="storage_id",
                help="The column name that contains storage_id in the columns data file"
            )
            
            demographic_keywords = st.text_area(
                "Demographic keywords (one per line)",
                value="age\ngender\nrace\nethnicity\nincome\neducation\nmarital_status\noccupation",
                help="Keywords to identify demographic columns in the data"
            )
        
        # Convert demographic keywords to list
        demographic_list = [keyword.strip().lower() for keyword in demographic_keywords.split('\n') if keyword.strip()]
        
        # Process button
        if st.button("🚀 Process Data", type="primary"):
            try:
                with st.spinner("Processing data..."):
                    # Initialize data processor
                    processor = DataProcessor(
                        storage_id_col_table=storage_id_col_table,
                        storage_id_col_columns=storage_id_col_columns,
                        table_name_col=table_name_col,
                        demographic_keywords=demographic_list
                    )
                    
                    # Process the data
                    result = processor.process_files(table_file, columns_file)
                    
                    if result['success']:
                        st.session_state.processed_data = result['data']
                        st.session_state.processing_complete = True
                        st.success("✅ Data processed successfully!")
                        st.rerun()
                    else:
                        st.error(f"❌ Processing failed: {result['error']}")
                        
            except Exception as e:
                st.error(f"❌ An unexpected error occurred: {str(e)}")
    
    # Results section
    if st.session_state.processing_complete and st.session_state.processed_data is not None:
        st.divider()
        st.subheader("📈 Processing Results")
        
        processed_df = st.session_state.processed_data
        
        # Display statistics
        col_stats1, col_stats2, col_stats3 = st.columns(3)
        
        with col_stats1:
            st.metric("Total Records", len(processed_df))
        
        with col_stats2:
            unique_tables = processed_df['table_name'].nunique() if 'table_name' in processed_df.columns else 0
            st.metric("Unique Tables", unique_tables)
        
        with col_stats3:
            demographic_cols = [col for col in processed_df.columns if any(keyword in col.lower() for keyword in demographic_list)]
            st.metric("Demographic Columns", len(demographic_cols))
        
        # Display the processed data
        st.write("**Processed Data:**")
        st.dataframe(processed_df, use_container_width=True)
        
        # Download section
        st.subheader("💾 Download Results")
        
        # Create Excel file for download
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            processed_df.to_excel(writer, sheet_name='Processed_Data', index=False)
        
        excel_data = output.getvalue()
        
        st.download_button(
            label="📥 Download Processed Data as Excel",
            data=excel_data,
            file_name="processed_demographic_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        # Option to reset and process new files
        if st.button("🔄 Process New Files"):
            st.session_state.processed_data = None
            st.session_state.processing_complete = False
            st.rerun()
    
    # Help section
    with st.expander("ℹ️ Help & Instructions"):
        st.markdown("""
        ### How to use this application:
        
        1. **Upload Table Data File**: This should contain table names and their corresponding storage_id values
        2. **Upload Columns Data File**: This should contain demographic information with storage_id for matching
        3. **Configure Processing**: 
           - Specify the column names for storage_id in both files
           - Specify the table name column in the table data file
           - Define demographic keywords to identify relevant columns
        4. **Process Data**: Click the process button to extract and merge data
        5. **Review Results**: Check the processed data and statistics
        6. **Download**: Save the results as an Excel file
        
        ### File Requirements:
        - Both files must be in Excel format (.xlsx or .xls)
        - Files must contain the specified storage_id columns for matching
        - Column names should be in the first row
        
        ### Demographic Detection:
        The application identifies demographic columns by searching for keywords in column names.
        You can customize these keywords in the configuration section.
        """)

if __name__ == "__main__":
    main()
