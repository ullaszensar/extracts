import pandas as pd
from typing import List, Dict, Any
from fuzzywuzzy import fuzz

class DataProcessor:
    """
    Handles the processing of Excel files to extract demographic data
    """
    
    def __init__(self, demographic_keywords: List[str] = None,
                 fuzzy_algorithm: str = "ratio", fuzzy_threshold: int = 80):
        """
        Initialize the data processor with configuration parameters
        
        Args:
            demographic_keywords: List of keywords to identify demographic columns
            fuzzy_algorithm: Algorithm for fuzzy matching ('ratio', 'partial_ratio', 'token_sort_ratio', 'token_set_ratio')
            fuzzy_threshold: Minimum similarity percentage (0-100) for fuzzy matching
        """
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
            
            # Contact Information
            "home phone", "business phone", "mobile phone", "cell phone", "work phone",
            "phone number", "telephone", "fax", "fax number",
            "home email", "business email", "work email", "personal email",
            "email address", "email", "e-mail",
            
            # Additional Demographics
            "member since date", "membership date", "registration date", "enrollment date",
            "customer since", "account opened"
        ]
        
    def process_files(self, table_file, columns_file) -> Dict[str, Any]:
        """
        Process the uploaded files and return demographic data
        
        Args:
            table_file: Uploaded table data file (optional, can be None)
            columns_file: Uploaded columns data file (required)
            
        Returns:
            Dictionary containing success status, data, and any error messages
        """
        try:
            # Read the columns file (required)
            columns_df = self._read_file(columns_file, "columns data")
            
            # Store original data statistics
            original_columns_total = len(columns_df)
            original_table_total = 0
            
            # Read table file if provided
            table_df = None
            if table_file is not None:
                table_df = self._read_file(table_file, "table data")
                original_table_total = len(table_df)
            
            # Extract demographic data from columns file
            demographic_data = self._extract_demographic_data(columns_df)
            
            if demographic_data.empty:
                return {'success': False, 'error': 'No demographic data found in columns file'}
            
            # Calculate extraction statistics
            demographic_rows_extracted = len(demographic_data)
            non_demographic_rows = original_columns_total - demographic_rows_extracted
            
            # Use demographic data as-is
            merged_data = demographic_data.copy()
            merged_data['table_name'] = 'N/A'
            merged_data['matched'] = True
            
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
            if file.name.lower().endswith('.csv'):
                return pd.read_csv(file)
            else:
                return pd.read_excel(file)
        except Exception as e:
            raise Exception(f"Error reading {file_type}: {str(e)}")
    
    def _extract_demographic_data(self, columns_df: pd.DataFrame) -> pd.DataFrame:
        """
        Extract demographic rows from the columns DataFrame while preserving all columns
        
        Args:
            columns_df: Columns data DataFrame
            
        Returns:
            DataFrame containing demographic rows with all original columns preserved
        """
        # First try to identify demographic rows by attr_description column
        if 'attr_description' in columns_df.columns:
            demographic_mask = self._identify_demographic_rows_by_description(columns_df, 'attr_description')
            
            if demographic_mask.any():
                # Return rows where demographic content was found, keeping ALL columns
                return columns_df[demographic_mask].copy()
        
        # Fallback: identify demographic rows by checking all column content
        demographic_mask = pd.Series([False] * len(columns_df), index=columns_df.index)
        
        for idx, row in columns_df.iterrows():
            # Check all text columns for demographic content
            for col in columns_df.columns:
                cell_value = str(row[col])
                if cell_value and cell_value.lower() != 'nan':
                    if self._fuzzy_match_demographic(cell_value, self.demographic_data_types + self.demographic_keywords):
                        demographic_mask[idx] = True
                        break
        
        if demographic_mask.any():
            # Return matching rows with ALL original columns preserved
            return columns_df[demographic_mask].copy()
        
        # If no demographic data found, return empty DataFrame
        return pd.DataFrame()
    
    def _identify_demographic_columns(self, columns) -> List[str]:
        """
        Identify which columns contain demographic information based on keywords
        
        Args:
            columns: List or Index of column names to check
            
        Returns:
            List of column names that match demographic keywords
        """
        demographic_cols = []
        
        # Combine predefined types with user-provided keywords
        all_keywords = self.demographic_data_types + self.demographic_keywords
        
        for col in columns:
            col_lower = str(col).lower()
            for keyword in all_keywords:
                if self._fuzzy_match_demographic(col_lower, [keyword]):
                    demographic_cols.append(col)
                    break
        
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
        # Combine predefined types with user-provided keywords
        all_keywords = self.demographic_data_types + self.demographic_keywords
        
        # Create boolean mask for demographic rows
        demographic_mask = pd.Series([False] * len(columns_df), index=columns_df.index)
        
        for idx, row in columns_df.iterrows():
            desc_text = str(row.get(attr_desc_col, ''))
            if desc_text and desc_text.lower() != 'nan':
                if self._fuzzy_match_demographic(desc_text, all_keywords):
                    demographic_mask[idx] = True
        
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
        
        for keyword in keywords:
            keyword_lower = keyword.lower()
            
            # Choose fuzzy matching algorithm
            if self.fuzzy_algorithm == "ratio":
                score = fuzz.ratio(text_lower, keyword_lower)
            elif self.fuzzy_algorithm == "partial_ratio":
                score = fuzz.partial_ratio(text_lower, keyword_lower)
            elif self.fuzzy_algorithm == "token_sort_ratio":
                score = fuzz.token_sort_ratio(text_lower, keyword_lower)
            elif self.fuzzy_algorithm == "token_set_ratio":
                score = fuzz.token_set_ratio(text_lower, keyword_lower)
            else:
                score = fuzz.ratio(text_lower, keyword_lower)
            
            if score >= self.fuzzy_threshold:
                return True
        
        return False
    
    def get_processing_summary(self, merged_df: pd.DataFrame) -> Dict[str, Any]:
        """
        Generate a summary of the processing results
        
        Args:
            merged_df: The merged DataFrame
            
        Returns:
            Dictionary containing processing statistics
        """
        # Get all original columns (excluding added processing columns)
        original_columns = [col for col in merged_df.columns if col not in ['table_name', 'matched']]
        
        summary = {
            'total_records': len(merged_df),
            'original_columns': original_columns,
            'original_column_count': len(original_columns),
            'demographic_columns': original_columns,  # All original columns are preserved
            'demographic_column_count': len(original_columns),
            'matched_records': merged_df['matched'].sum() if 'matched' in merged_df.columns else len(merged_df),
            'unmatched_records': len(merged_df) - (merged_df['matched'].sum() if 'matched' in merged_df.columns else len(merged_df)),
            'processing_algorithm': self.fuzzy_algorithm,
            'processing_threshold': self.fuzzy_threshold
        }
        
        return summary