import pandas as pd
import streamlit as st
from typing import Any, Optional
import base64
import io

def validate_excel_file(file) -> bool:
    """
    Validate if the uploaded file is a valid Excel file
    
    Args:
        file: Uploaded file object
        
    Returns:
        Boolean indicating if file is valid Excel format
    """
    if file is None:
        return False
    
    # Check file extension
    if not file.name.lower().endswith(('.xlsx', '.xls')):
        return False
    
    try:
        # Try to read the file to ensure it's valid Excel
        file.seek(0)  # Reset file pointer
        pd.read_excel(file, nrows=1)  # Read just one row to test
        file.seek(0)  # Reset file pointer again
        return True
    except Exception:
        return False

def create_download_link(df: pd.DataFrame, filename: str, link_text: str) -> str:
    """
    Create a download link for DataFrame as Excel file
    
    Args:
        df: DataFrame to download
        filename: Name for the downloaded file
        link_text: Text to display for the download link
        
    Returns:
        HTML string for download link
    """
    try:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Data')
        
        excel_data = output.getvalue()
        b64 = base64.b64encode(excel_data).decode()
        
        href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}">{link_text}</a>'
        return href
    
    except Exception as e:
        st.error(f"Error creating download link: {str(e)}")
        return ""

def display_dataframe_info(df: pd.DataFrame, title: str) -> None:
    """
    Display useful information about a DataFrame
    
    Args:
        df: DataFrame to analyze
        title: Title for the information display
    """
    st.subheader(title)
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric("Rows", len(df))
    
    with col2:
        st.metric("Columns", len(df.columns))
    
    with col3:
        null_count = df.isnull().sum().sum()
        st.metric("Null Values", null_count)
    
    # Show column information
    with st.expander("Column Information"):
        col_info = pd.DataFrame({
            'Column': df.columns,
            'Data Type': df.dtypes.astype(str),
            'Null Count': df.isnull().sum().values,
            'Null %': (df.isnull().sum() / len(df) * 100).round(2).values
        })
        st.dataframe(col_info, use_container_width=True)

def format_processing_error(error_msg: str) -> str:
    """
    Format error messages for better user experience
    
    Args:
        error_msg: Raw error message
        
    Returns:
        Formatted error message
    """
    if "not found" in error_msg.lower():
        return f"🔍 Column Error: {error_msg}"
    elif "empty" in error_msg.lower():
        return f"📄 File Error: {error_msg}"
    elif "read" in error_msg.lower():
        return f"📖 Reading Error: {error_msg}"
    else:
        return f"⚠️ Error: {error_msg}"

def safe_column_name(name: str) -> str:
    """
    Create a safe column name by removing special characters
    
    Args:
        name: Original column name
        
    Returns:
        Sanitized column name
    """
    import re
    # Replace spaces and special characters with underscores
    safe_name = re.sub(r'[^\w\s]', '', name)
    safe_name = re.sub(r'\s+', '_', safe_name)
    return safe_name.lower()

def get_sample_data_info() -> dict:
    """
    Provide information about expected data format
    
    Returns:
        Dictionary with sample data structure information
    """
    return {
        "table_data_expected_columns": [
            "storage_id", "table_name", "description"
        ],
        "columns_data_expected_columns": [
            "storage_id", "age", "gender", "race", "income", "education"
        ],
        "demographic_keywords": [
            "age", "gender", "race", "ethnicity", "income", "education",
            "marital_status", "occupation", "employment", "household"
        ]
    }

def validate_storage_ids(table_df: pd.DataFrame, columns_df: pd.DataFrame, 
                        table_id_col: str, columns_id_col: str) -> dict:
    """
    Validate storage IDs between two DataFrames
    
    Args:
        table_df: Table data DataFrame
        columns_df: Columns data DataFrame
        table_id_col: Storage ID column name in table data
        columns_id_col: Storage ID column name in columns data
        
    Returns:
        Dictionary with validation results
    """
    try:
        table_ids = set(table_df[table_id_col].dropna())
        column_ids = set(columns_df[columns_id_col].dropna())
        
        common_ids = table_ids.intersection(column_ids)
        table_only = table_ids - column_ids
        columns_only = column_ids - table_ids
        
        return {
            'valid': True,
            'total_table_ids': len(table_ids),
            'total_column_ids': len(column_ids),
            'common_ids': len(common_ids),
            'table_only_ids': len(table_only),
            'columns_only_ids': len(columns_only),
            'match_percentage': (len(common_ids) / max(len(column_ids), 1)) * 100
        }
    
    except Exception as e:
        return {
            'valid': False,
            'error': str(e)
        }
