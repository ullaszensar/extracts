import pandas as pd
import numpy as np
from typing import Dict, List, Any, Optional
import logging
from fuzzywuzzy import fuzz, process

class DataProcessor:
    """
    Handles the processing of Excel files to extract demographic data and merge with table information
    """
    
    def __init__(self, table_name_col: str = "table_name", demographic_keywords: List[str] = None,
                 fuzzy_algorithm: str = "ratio", fuzzy_threshold: int = 80):
        """
        Initialize the data processor with configuration parameters
        
        Args:
            table_name_col: Column name for table names in table data (optional)
            demographic_keywords: List of keywords to identify demographic columns
            fuzzy_algorithm: Algorithm for fuzzy matching ('ratio', 'partial_ratio', 'token_sort_ratio', 'token_set_ratio')
            fuzzy_threshold: Minimum similarity percentage (0-100) for fuzzy matching
        """
        self.table_name_col = table_name_col
        self.demographic_keywords = [keyword.lower() for keyword in demographic_keywords] if demographic_keywords else []
        self.fuzzy_algorithm = fuzzy_algorithm
        self.fuzzy_threshold = fuzzy_threshold
        
        # Define specific demographic data types as provided
        self.demographic_data_types = [
            # Name Information
            "embossed name", "embossed company name", "primary name", "secondary name",
            "legal name", "dba name", "double byte name",
            
            # Personal Demographics
            "gender", "dob", "date of birth", "birth date",
            
            # Identification
            "gov ids", "government ids", "government identification", "social security",
            "ssn", "tax id", "identification number",
            
            # Address Information
            "home address", "business address", "alternate address", "temporary address",
            "other address", "additional addresses", "mailing address", "billing address",
            "shipping address", "residential address", "work address",
            
            # Phone Information
            "home phone", "alternate home phone", "business phone", "alternate business phone",
            "mobile phone", "alternate mobile phone", "attorney phone", "fax", "ani phone",
            "other phone", "additional phone", "cell phone", "work phone", "office phone",
            "contact number", "telephone", "phone number",
            
            # Email Information
            "servicing email", "estatement email", "business email", "other email address",
            "additional email", "contact email", "work email", "personal email",
            "primary email", "secondary email",
            
            # Preferences and Dates
            "preference language cd", "language preference", "preferred language",
            "member since date", "membership date", "registration date", "enrollment date",
            "customer since", "account opened"
        ]
        
    def process_files(self, table_file, columns_file) -> Dict[str, Any]:
        """
        Process the uploaded files and return demographic data (table file is optional)
        
        Args:
            table_file: Uploaded table data file (optional, can be None)
            columns_file: Uploaded columns data file (required)
            
        Returns:
            Dictionary containing success status, data, and any error messages
        """
        try:
            # Read the columns file (required)
            columns_df = self._read_excel_file(columns_file, "columns data")
            
            # Store original data statistics
            original_columns_total = len(columns_df)
            original_table_total = 0
            
            # Read table file if provided
            table_df = None
            if table_file is not None:
                table_df = self._read_excel_file(table_file, "table data")
                original_table_total = len(table_df)
                
                # Validate required columns exist when table file is provided
                validation_result = self._validate_columns(table_df, columns_df)
                if not validation_result['valid']:
                    return {'success': False, 'error': validation_result['error']}
            else:
                # When no table file is provided, skip validation entirely
                pass
            
            # Extract demographic data from columns file
            demographic_data = self._extract_demographic_data(columns_df)
            
            if demographic_data.empty:
                return {'success': False, 'error': 'No demographic data found in columns file'}
            
            # Calculate extraction statistics
            demographic_rows_extracted = len(demographic_data)
            non_demographic_rows = original_columns_total - demographic_rows_extracted
            
            # Merge with table names if table file is provided
            if table_df is not None:
                merged_data = self._merge_with_table_names(demographic_data, table_df)
            else:
                # Use demographic data as-is without table name merging
                merged_data = demographic_data.copy()
                merged_data['table_name'] = 'N/A'
                merged_data['matched'] = True  # All records are considered "matched" since no table matching is needed
            
            # Generate processing summary
            processing_stats = self.get_processing_summary(merged_data)
            processing_stats['original_columns_total'] = original_columns_total
            processing_stats['original_table_total'] = original_table_total
            processing_stats['demographic_rows_extracted'] = demographic_rows_extracted
            processing_stats['non_demographic_rows'] = non_demographic_rows
            processing_stats['extraction_percentage'] = round((demographic_rows_extracted / original_columns_total) * 100, 2) if original_columns_total > 0 else 0
            
            return {
                'success': True, 
                'data': merged_data,
                'stats': processing_stats
            }
            
        except Exception as e:
            return {'success': False, 'error': f'Processing error: {str(e)}'}
    
    def _read_file(self, file, file_type: str) -> pd.DataFrame:
        """
        Read Excel or CSV file and handle potential errors
        
        Args:
            file: File object to read
            file_type: Description of file type for error messages
            
        Returns:
            DataFrame containing the file data
        """
        try:
            # Reset file pointer to beginning
            file.seek(0)
            
            # Determine file type and read accordingly
            if file.name.lower().endswith('.csv'):
                df = pd.read_csv(file)
            else:
                df = pd.read_excel(file)
            
            if df.empty:
                raise ValueError(f"{file_type} file is empty")
                
            return df
            
        except Exception as e:
            raise ValueError(f"Error reading {file_type} file: {str(e)}")
    
    def _read_excel_file(self, file, file_type: str) -> pd.DataFrame:
        """
        Backward compatibility function - now reads both Excel and CSV
        """
        return self._read_file(file, file_type)
    
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
    
    def _validate_columns_only(self, columns_df: pd.DataFrame) -> Dict[str, Any]:
        """
        Validate that required columns exist in columns DataFrame only
        When no table file is provided, we don't require storage_id column
        
        Args:
            columns_df: Columns data DataFrame
            
        Returns:
            Dictionary with validation result and error message if any
        """
        # When processing only columns file, we don't require storage_id
        # Just check that we have some data to work with
        if columns_df.empty:
            return {
                'valid': False,
                'error': "Columns data file is empty"
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
        # Check if attr_description column exists
        attr_desc_col = 'attr_description'
        if attr_desc_col not in columns_df.columns:
            # Fallback to original method if attr_description doesn't exist
            demographic_cols = self._identify_demographic_columns(columns_df.columns)
            cols_to_extract = [self.storage_id_col_columns] + demographic_cols
            demographic_df = columns_df[cols_to_extract].copy()
        else:
            # Use attr_description column to identify demographic rows
            demographic_rows_mask = self._identify_demographic_rows_by_description(columns_df, attr_desc_col)
            
            # Filter to demographic rows only
            demographic_df = columns_df[demographic_rows_mask].copy()
        
        # Remove rows where storage_id is null
        demographic_df = demographic_df.dropna(subset=[self.storage_id_col_columns])
        
        return demographic_df
    
    def _identify_demographic_columns(self, columns) -> List[str]:
        """
        Identify which columns contain demographic information based on keywords
        
        Args:
            columns: List or Index of column names to check
            
        Returns:
            List of column names that match demographic keywords
        """
        demographic_cols = []
        
        for col in columns:
            col_lower = str(col).lower()
            # Check if any demographic keyword is in the column name
            if any(keyword in col_lower for keyword in self.demographic_keywords):
                demographic_cols.append(col)
        
        return demographic_cols
    
    def _identify_demographic_rows_by_description(self, columns_df: pd.DataFrame, attr_desc_col: str) -> pd.Series:
        """
        Identify rows containing demographic information based on attr_description column content using fuzzy matching
        
        Args:
            columns_df: Columns data DataFrame
            attr_desc_col: Name of the attribute description column
            
        Returns:
            Boolean Series indicating which rows contain demographic data
        """
        # Create a boolean mask for demographic rows
        demographic_mask = pd.Series([False] * len(columns_df), index=columns_df.index)
        
        # Combine user keywords with comprehensive demographic data types
        all_keywords = list(set(self.demographic_keywords + self.demographic_data_types))
        
        # Check each row's attr_description for demographic keywords using fuzzy matching
        for idx, row in columns_df.iterrows():
            description = str(row.get(attr_desc_col, ''))
            
            if description and description.lower() != 'nan':
                # Use fuzzy matching to find demographic content
                is_demographic = self._fuzzy_match_demographic(description, all_keywords)
                demographic_mask.loc[idx] = is_demographic
        
        return demographic_mask
    
    def _fuzzy_match_demographic(self, text: str, keywords: List[str]) -> bool:
        """
        Use fuzzy matching to determine if text contains demographic information
        
        Args:
            text: Text to analyze
            keywords: List of demographic keywords to match against
            
        Returns:
            Boolean indicating if text matches demographic criteria
        """
        text_lower = text.lower()
        
        # Get the fuzzy matching function based on selected algorithm
        fuzzy_func = getattr(fuzz, self.fuzzy_algorithm, fuzz.ratio)
        
        # Check each keyword for fuzzy match
        for keyword in keywords:
            # Calculate similarity score
            score = fuzzy_func(text_lower, keyword.lower())
            
            if score >= self.fuzzy_threshold:
                return True
            
            # Also check if keyword is substring (for partial matches)
            if keyword.lower() in text_lower:
                return True
        
        # Use process.extractOne for best match
        best_match = process.extractOne(text_lower, keywords, scorer=fuzzy_func)
        if best_match and best_match[1] >= self.fuzzy_threshold:
            return True
        
        return False
    
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
        table_mapping = table_mapping.drop_duplicates(subset=self.storage_id_col_table)
        
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
