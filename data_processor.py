import pandas as pd
import numpy as np
from typing import Dict, List, Any, Optional
import logging

class DataProcessor:
    """
    Handles the processing of Excel files to extract demographic data and merge with table information
    """
    
    def __init__(self, storage_id_col_table: str, storage_id_col_columns: str, 
                 table_name_col: str, demographic_keywords: List[str]):
        """
        Initialize the data processor with configuration parameters
        
        Args:
            storage_id_col_table: Column name for storage_id in table data
            storage_id_col_columns: Column name for storage_id in columns data
            table_name_col: Column name for table names in table data
            demographic_keywords: List of keywords to identify demographic columns
        """
        self.storage_id_col_table = storage_id_col_table
        self.storage_id_col_columns = storage_id_col_columns
        self.table_name_col = table_name_col
        self.demographic_keywords = [keyword.lower() for keyword in demographic_keywords]
        
    def process_files(self, table_file, columns_file) -> Dict[str, Any]:
        """
        Process the uploaded Excel files and return merged demographic data
        
        Args:
            table_file: Uploaded table data Excel file
            columns_file: Uploaded columns data Excel file
            
        Returns:
            Dictionary containing success status, data, and any error messages
        """
        try:
            # Read the Excel files
            table_df = self._read_excel_file(table_file, "table data")
            columns_df = self._read_excel_file(columns_file, "columns data")
            
            # Validate required columns exist
            validation_result = self._validate_columns(table_df, columns_df)
            if not validation_result['valid']:
                return {'success': False, 'error': validation_result['error']}
            
            # Extract demographic data from columns file
            demographic_data = self._extract_demographic_data(columns_df)
            
            if demographic_data.empty:
                return {'success': False, 'error': 'No demographic data found in columns file'}
            
            # Merge with table names
            merged_data = self._merge_with_table_names(demographic_data, table_df)
            
            return {'success': True, 'data': merged_data}
            
        except Exception as e:
            return {'success': False, 'error': f'Processing error: {str(e)}'}
    
    def _read_excel_file(self, file, file_type: str) -> pd.DataFrame:
        """
        Read Excel file and handle potential errors
        
        Args:
            file: File object to read
            file_type: Description of file type for error messages
            
        Returns:
            DataFrame containing the file data
        """
        try:
            # Reset file pointer to beginning
            file.seek(0)
            df = pd.read_excel(file)
            
            if df.empty:
                raise ValueError(f"{file_type} file is empty")
                
            return df
            
        except Exception as e:
            raise ValueError(f"Error reading {file_type} file: {str(e)}")
    
    def _validate_columns(self, table_df: pd.DataFrame, columns_df: pd.DataFrame) -> Dict[str, Any]:
        """
        Validate that required columns exist in both DataFrames
        
        Args:
            table_df: Table data DataFrame
            columns_df: Columns data DataFrame
            
        Returns:
            Dictionary with validation result and error message if any
        """
        # Check table data columns
        if self.storage_id_col_table not in table_df.columns:
            return {
                'valid': False, 
                'error': f"Column '{self.storage_id_col_table}' not found in table data. Available columns: {list(table_df.columns)}"
            }
        
        if self.table_name_col not in table_df.columns:
            return {
                'valid': False,
                'error': f"Column '{self.table_name_col}' not found in table data. Available columns: {list(table_df.columns)}"
            }
        
        # Check columns data
        if self.storage_id_col_columns not in columns_df.columns:
            return {
                'valid': False,
                'error': f"Column '{self.storage_id_col_columns}' not found in columns data. Available columns: {list(columns_df.columns)}"
            }
        
        return {'valid': True}
    
    def _extract_demographic_data(self, columns_df: pd.DataFrame) -> pd.DataFrame:
        """
        Extract demographic columns and data from the columns DataFrame
        
        Args:
            columns_df: Columns data DataFrame
            
        Returns:
            DataFrame containing only demographic data and storage_id
        """
        # Identify demographic columns
        demographic_cols = self._identify_demographic_columns(columns_df.columns)
        
        # Always include storage_id column
        cols_to_extract = [self.storage_id_col_columns] + demographic_cols
        
        # Filter DataFrame to include only relevant columns
        demographic_df = columns_df[cols_to_extract].copy()
        
        # Remove rows where storage_id is null
        demographic_df = demographic_df.dropna(subset=[self.storage_id_col_columns])
        
        # Remove rows where all demographic columns are null
        if demographic_cols:
            demographic_df = demographic_df.dropna(subset=demographic_cols, how='all')
        
        return demographic_df
    
    def _identify_demographic_columns(self, columns: List[str]) -> List[str]:
        """
        Identify which columns contain demographic information based on keywords
        
        Args:
            columns: List of column names to check
            
        Returns:
            List of column names that match demographic keywords
        """
        demographic_cols = []
        
        for col in columns:
            col_lower = col.lower()
            # Check if any demographic keyword is in the column name
            if any(keyword in col_lower for keyword in self.demographic_keywords):
                demographic_cols.append(col)
        
        return demographic_cols
    
    def _merge_with_table_names(self, demographic_df: pd.DataFrame, table_df: pd.DataFrame) -> pd.DataFrame:
        """
        Merge demographic data with table names based on storage_id
        
        Args:
            demographic_df: DataFrame containing demographic data
            table_df: DataFrame containing table names and storage_ids
            
        Returns:
            Merged DataFrame with table names added
        """
        # Create a mapping of storage_id to table_name
        table_mapping = table_df[[self.storage_id_col_table, self.table_name_col]].copy()
        table_mapping = table_mapping.drop_duplicates(subset=[self.storage_id_col_table])
        
        # Rename columns for merging
        table_mapping = table_mapping.rename(columns={
            self.storage_id_col_table: self.storage_id_col_columns,
            self.table_name_col: 'table_name'
        })
        
        # Merge demographic data with table names
        merged_df = demographic_df.merge(
            table_mapping,
            on=self.storage_id_col_columns,
            how='left'
        )
        
        # Add a flag for unmatched records
        merged_df['matched'] = merged_df['table_name'].notna()
        
        # Fill missing table names with a placeholder
        merged_df['table_name'] = merged_df['table_name'].fillna('No_Match_Found')
        
        # Reorder columns to put table_name and matched status first
        cols = ['table_name', 'matched', self.storage_id_col_columns]
        demographic_cols = [col for col in merged_df.columns if col not in cols]
        merged_df = merged_df[cols + demographic_cols]
        
        return merged_df
    
    def get_processing_summary(self, merged_df: pd.DataFrame) -> Dict[str, Any]:
        """
        Generate a summary of the processing results
        
        Args:
            merged_df: The merged DataFrame
            
        Returns:
            Dictionary containing processing statistics
        """
        total_records = len(merged_df)
        matched_records = merged_df['matched'].sum() if 'matched' in merged_df.columns else 0
        unmatched_records = total_records - matched_records
        
        demographic_cols = [col for col in merged_df.columns 
                          if col not in ['table_name', 'matched', self.storage_id_col_columns]]
        
        unique_tables = merged_df['table_name'].nunique() if 'table_name' in merged_df.columns else 0
        
        return {
            'total_records': total_records,
            'matched_records': int(matched_records),
            'unmatched_records': unmatched_records,
            'demographic_columns': len(demographic_cols),
            'unique_tables': unique_tables,
            'demographic_column_names': demographic_cols
        }
