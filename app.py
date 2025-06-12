import streamlit as st
import pandas as pd
import io
from data_processor import DataProcessor
from utils import validate_excel_file, create_download_link
from report_generator import ReportGenerator

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
            "Choose table data file",
            type=['xlsx', 'xls', 'csv'],
            key="table_file",
            help="This file should contain table names and their corresponding storage_id values. Supports Excel (.xlsx, .xls) and CSV (.csv) formats"
        )
        
        if table_file:
            if validate_excel_file(table_file):
                try:
                    # Get full file info for row count
                    if table_file.name.lower().endswith('.csv'):
                        df_full = pd.read_csv(table_file)
                        df_preview = df_full.head(5)
                    else:
                        df_full = pd.read_excel(table_file)
                        df_preview = df_full.head(5)
                    
                    st.success(f"✅ Table file loaded successfully with **{len(df_full)} rows** and **{len(df_full.columns)} columns**")
                    st.write("**Preview (first 5 rows):**")
                    st.dataframe(df_preview)
                except Exception as e:
                    st.error(f"Error reading file: {str(e)}")
            else:
                st.error("❌ Invalid file format. Please upload an Excel (.xlsx, .xls) or CSV (.csv) file")
    
    with col2:
        st.subheader("📋 Columns Data File")
        st.markdown("Upload file containing demographic data with storage_id")
        columns_file = st.file_uploader(
            "Choose columns data file",
            type=['xlsx', 'xls', 'csv'],
            key="columns_file",
            help="This file should contain demographic information and storage_id for matching. Supports Excel (.xlsx, .xls) and CSV (.csv) formats"
        )
        
        if columns_file:
            if validate_excel_file(columns_file):
                try:
                    # Get full file info for row count
                    if columns_file.name.lower().endswith('.csv'):
                        df_full = pd.read_csv(columns_file)
                        df_preview = df_full.head(5)
                    else:
                        df_full = pd.read_excel(columns_file)
                        df_preview = df_full.head(5)
                    
                    st.success(f"✅ Columns file loaded successfully with **{len(df_full)} rows** and **{len(df_full.columns)} columns**")
                    st.write("**Preview (first 5 rows):**")
                    st.dataframe(df_preview)
                except Exception as e:
                    st.error(f"Error reading file: {str(e)}")
            else:
                st.error("❌ Invalid file format. Please upload an Excel (.xlsx, .xls) or CSV (.csv) file")
    
    # Processing section
    if columns_file:
        st.divider()
        
        # Configuration section
        st.subheader("⚙️ Processing Configuration")
        
        # Show file status and configuration
        if table_file:
            st.info("📊 Table data file uploaded - additional table information available")
        else:
            st.warning("📋 Processing with columns data file only - extracting demographic data from attr_description column")
        
        demographic_keywords = st.text_area(
            "Additional demographic keywords (one per line)",
            value="embossed name\nprimary name\nlegal name\ngender\ndob\nhome address\nbusiness address\nhome phone\nmobile phone\nservicing email",
            help="Extra keywords to identify demographic data. The system includes the standard demographic types by default."
        )
        
        # Fuzzy matching configuration
        st.subheader("🔍 Fuzzy Matching Configuration")
        
        fuzzy_col1, fuzzy_col2 = st.columns(2)
        
        with fuzzy_col1:
            fuzzy_algorithm = st.selectbox(
                "Fuzzy Matching Algorithm",
                options=["ratio", "partial_ratio", "token_sort_ratio", "token_set_ratio"],
                index=0,
                help="Choose the algorithm for fuzzy string matching"
            )
            
            algorithm_descriptions = {
                "ratio": "Basic string similarity comparison",
                "partial_ratio": "Best substring match similarity",
                "token_sort_ratio": "Sorted token comparison",
                "token_set_ratio": "Set-based token comparison"
            }
            
            st.info(f"**{fuzzy_algorithm}**: {algorithm_descriptions[fuzzy_algorithm]}")
        
        with fuzzy_col2:
            fuzzy_threshold = st.slider(
                "Accuracy Threshold (%)",
                min_value=50,
                max_value=100,
                value=80,
                step=5,
                help="Minimum similarity percentage for fuzzy matching (higher = more strict)"
            )
            
            st.metric("Current Threshold", f"{fuzzy_threshold}%")
        
        # Show built-in demographic data types
        with st.expander("📋 Built-in Demographic Data Types"):
            st.write("**Name Information:**")
            st.write("Embossed Name, Primary Name, Legal Name, DBA Name, Double Byte Name")
            
            st.write("**Personal Demographics:**")
            st.write("Gender, DOB (Date of Birth)")
            
            st.write("**Identification:**")
            st.write("Gov IDs, Government Identification")
            
            st.write("**Address Information:**")
            st.write("Home Address, Business Address, Alternate Address, Temporary Address, Other Address")
            
            st.write("**Phone Information:**")
            st.write("Home Phone, Business Phone, Mobile Phone, Attorney Phone, Fax, ANI Phone")
            
            st.write("**Email Information:**")
            st.write("Servicing Email, Estatement Email, Business Email, Other Email Address")
            
            st.write("**Preferences:**")
            st.write("Preference Language CD, Member Since Date")
        
        # Convert demographic keywords to list
        demographic_list = [keyword.strip().lower() for keyword in demographic_keywords.split('\n') if keyword.strip()]
        
        # Process button
        if st.button("🚀 Process Data", type="primary"):
            try:
                with st.spinner("Processing data..."):
                    # Initialize data processor
                    processor = DataProcessor(
                        demographic_keywords=demographic_list,
                        fuzzy_algorithm=fuzzy_algorithm,
                        fuzzy_threshold=fuzzy_threshold
                    )
                    
                    # Process the data (table_file can be None)
                    result = processor.process_files(table_file, columns_file)
                    
                    if result['success']:
                        st.session_state.processed_data = result['data']
                        st.session_state.processing_stats = result.get('stats', {})
                        st.session_state.fuzzy_algorithm = fuzzy_algorithm
                        st.session_state.fuzzy_threshold = fuzzy_threshold
                        st.session_state.table_file_provided = table_file is not None
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
        
        # Get processing statistics
        stats = st.session_state.get('processing_stats', {})
        
        # Display main statistics
        col_stats1, col_stats2, col_stats3, col_stats4 = st.columns(4)
        
        with col_stats1:
            st.metric("Total Input Rows", stats.get('original_columns_total', 0))
        
        with col_stats2:
            st.metric("Demographic Rows Extracted", stats.get('demographic_rows_extracted', 0))
        
        with col_stats3:
            st.metric("Non-Demographic Rows", stats.get('non_demographic_rows', 0))
        
        with col_stats4:
            extraction_pct = stats.get('extraction_percentage', 0)
            st.metric("Extraction Rate", f"{extraction_pct}%")
        
        # Additional statistics
        st.divider()
        col_extra1, col_extra2, col_extra3 = st.columns(3)
        
        with col_extra1:
            st.metric("Final Processed Records", len(processed_df))
        
        with col_extra2:
            unique_tables = processed_df['table_name'].nunique() if 'table_name' in processed_df.columns else 0
            st.metric("Unique Tables", unique_tables)
        
        with col_extra3:
            matched_records = stats.get('matched_records', 0)
            st.metric("Successfully Matched", matched_records)
        
        # Detailed breakdown section
        st.subheader("📊 Detailed Extraction Breakdown")
        
        breakdown_col1, breakdown_col2 = st.columns(2)
        
        with breakdown_col1:
            st.write("**Data Flow Summary:**")
            original_total = stats.get('original_columns_total', 0)
            extracted = stats.get('demographic_rows_extracted', 0)
            non_demo = stats.get('non_demographic_rows', 0)
            final_count = len(processed_df)
            
            st.write(f"• Started with: **{original_total}** total rows in columns file")
            st.write(f"• Found demographic data in: **{extracted}** rows")
            st.write(f"• Non-demographic rows: **{non_demo}** rows")
            st.write(f"• Final processed records: **{final_count}** rows")
            
            if extracted > 0:
                st.write(f"• Extraction efficiency: **{stats.get('extraction_percentage', 0)}%**")
        
        with breakdown_col2:
            st.write("**Detection Method:**")
            # Check if attr_description was used
            if 'attr_description' in processed_df.columns:
                st.write("✓ Used 'attr_description' column content")
                st.write("✓ Applied fuzzy matching algorithm")
                st.write(f"✓ Algorithm: **{st.session_state.get('fuzzy_algorithm', 'ratio')}**")
                st.write(f"✓ Threshold: **{st.session_state.get('fuzzy_threshold', 80)}%**")
            else:
                st.write("✓ Used column name pattern matching")
                demo_cols = stats.get('demographic_column_names', [])
                if demo_cols:
                    st.write("**Demographic Columns Found:**")
                    for i, col in enumerate(demo_cols, 1):
                        st.write(f"{i}. {col}")
                else:
                    st.write("No demographic columns identified")
        
        # Skip Storage ID matching results display as requested
        
        # Display the processed data with original columns only
        st.subheader("📋 Processed Data")
        
        # Show all original columns from the source file including table_name, attr_name, business_name, attr_description
        original_columns_only = processed_df[[col for col in processed_df.columns if col != 'matched']]
        
        # Show column preservation info
        original_col_count = len(original_columns_only.columns)
        st.info(f"Displaying {len(original_columns_only)} demographic records with all {original_col_count} original columns preserved")
        
        st.dataframe(original_columns_only, use_container_width=True)
        
        # Download section
        st.subheader("💾 Download Results")
        
        # Check if this is columns-only mode (no table file)
        table_file_provided = st.session_state.get('table_file_provided', False)
        
        if not table_file_provided:
            # Special export options for columns-only mode
            st.info("📂 Columns-only mode: Generate 20 separate Excel files for easy processing")
            
            download_col1, download_col2, download_col3 = st.columns(3)
            
            with download_col1:
                if st.button("📊 Generate 20 Excel Files", type="primary"):
                    with st.spinner("Creating 20 Excel files with all original columns preserved..."):
                        report_gen = ReportGenerator()
                        excel_files = report_gen.create_multiple_excel_files(processed_df)
                        st.session_state.excel_files = excel_files
                    
                    original_cols = len([col for col in processed_df.columns if col != 'matched'])
                    st.success(f"Generated {len(excel_files)} Excel files ready for download!")
                    st.info(f"Each file includes all {original_cols} columns: table_name, attr_name, business_name, attr_description and other original columns")
                    
                    # Display download buttons for each file
                    st.subheader("📥 Download Individual Files")
                    cols = st.columns(4)  # 4 columns for better layout
                    
                    for idx, (filename, file_data) in enumerate(excel_files):
                        col_idx = idx % 4
                        with cols[col_idx]:
                            st.download_button(
                                label=f"📄 Part {idx+1:02d}",
                                data=file_data,
                                file_name=filename,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key=f"download_{idx}",
                                help=f"Download {filename} with all original columns"
                            )
            
            with download_col2:
                # Create CSV file for download
                report_gen = ReportGenerator()
                csv_data = report_gen.create_csv_export(processed_df)
                
                st.download_button(
                    label="📄 Download CSV Data",
                    data=csv_data,
                    file_name="demographic_analysis_data.csv",
                    mime="text/csv",
                    help="Processed data in CSV format for easy import"
                )
            
            with download_col3:
                # Generate HTML analysis report
                if st.button("📊 Generate Analysis Report", type="secondary"):
                    with st.spinner("Generating comprehensive analysis report..."):
                        report_gen = ReportGenerator()
                        fuzzy_alg = st.session_state.get('fuzzy_algorithm', 'ratio')
                        fuzzy_thresh = st.session_state.get('fuzzy_threshold', 80)
                        
                        html_report = report_gen.generate_report(
                            processed_df, 
                            stats, 
                            fuzzy_alg, 
                            fuzzy_thresh
                        )
                        
                        st.session_state.html_report = html_report
                        st.success("Analysis report generated successfully!")
                        st.rerun()
        else:
            # Standard export options when table file is provided
            download_col1, download_col2, download_col3 = st.columns(3)
            
            with download_col1:
                # Create comprehensive Excel file for download
                report_gen = ReportGenerator()
                excel_data = report_gen.create_excel_export(processed_df, stats)
                
                st.download_button(
                    label="📥 Download Excel Report",
                    data=excel_data,
                    file_name="demographic_analysis_complete.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    help="Multi-sheet Excel file with processed data, statistics, and analysis"
                )
            
            with download_col2:
                # Create CSV file for download
                csv_data = report_gen.create_csv_export(processed_df)
                
                st.download_button(
                    label="📄 Download CSV Data",
                    data=csv_data,
                    file_name="demographic_analysis_data.csv",
                    mime="text/csv",
                    help="Processed data in CSV format for easy import"
                )
            
            with download_col3:
                # Generate HTML analysis report
                if st.button("📊 Generate Analysis Report", type="secondary"):
                    with st.spinner("Generating comprehensive analysis report..."):
                        report_gen = ReportGenerator()
                        fuzzy_alg = st.session_state.get('fuzzy_algorithm', 'ratio')
                        fuzzy_thresh = st.session_state.get('fuzzy_threshold', 80)
                        
                        html_report = report_gen.generate_report(
                            processed_df, 
                            stats, 
                            fuzzy_alg, 
                            fuzzy_thresh
                        )
                        
                        st.session_state.html_report = html_report
                        st.success("Analysis report generated successfully!")
                        st.rerun()
        
        # Display HTML report if generated
        if 'html_report' in st.session_state and st.session_state.html_report:
            st.divider()
            st.subheader("📋 Analysis Report Preview")
            
            # Download HTML report
            st.download_button(
                label="📄 Download HTML Analysis Report",
                data=st.session_state.html_report.encode('utf-8'),
                file_name="demographic_analysis_report.html",
                mime="text/html",
                help="Complete analysis report with charts and detailed statistics"
            )
            
            # Display report in expander
            with st.expander("📖 View Report Content", expanded=False):
                st.markdown("**Report Preview:**")
                st.info("📊 The complete HTML report includes interactive charts, detailed statistics, and comprehensive analysis tables. Download the HTML file to view the full interactive report.")
        
        # Option to reset and process new files
        if st.button("🔄 Process New Files"):
            st.session_state.processed_data = None
            st.session_state.processing_complete = False
            st.session_state.processing_stats = None
            if 'html_report' in st.session_state:
                del st.session_state.html_report
            st.rerun()
    
    # Help section
    with st.expander("ℹ️ Help & Instructions"):
        st.markdown("""
        ### How to use this application:
        
        1. **Upload Columns Data File** (Required): Excel or CSV file containing demographic information
        2. **Upload Table Data File** (Optional): Additional table information for reference
        3. **Configure Processing**: 
           - Add custom demographic keywords if needed
           - The system automatically detects demographic data from attr_description column content
        4. **Configure Fuzzy Matching**:
           - Select the fuzzy matching algorithm (ratio, partial_ratio, token_sort_ratio, token_set_ratio)
           - Set the accuracy threshold percentage (50-100%)
        5. **Process Data**: Extract demographic data using intelligent content analysis
        6. **Review Results**: View processed data with row counts and statistics
        7. **Export Results**:
           - Generate 20 separate Excel files for easy processing (columns-only mode)
           - Download CSV data for easy import
           - Generate comprehensive HTML analysis report with charts and graphs
        
        ### File Requirements:
        - Columns file: Excel (.xlsx, .xls) or CSV (.csv) format (Required)
        - Table file: Excel (.xlsx, .xls) or CSV (.csv) format (Optional)
        - Column names should be in the first row
        - attr_description column provides the best demographic detection results
        
        ### Demographic Detection:
        The application uses two methods to identify demographic data:
        1. **Primary Method**: If an 'attr_description' column exists, it searches for demographic keywords within the description text of each row
        2. **Fallback Method**: If no 'attr_description' column is found, it searches for keywords in column names
        
        You can customize the demographic keywords in the configuration section.
        """)

if __name__ == "__main__":
    main()
