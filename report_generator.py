import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import plotly.offline as pyo
from datetime import datetime
from jinja2 import Template
import base64
import io
from typing import Dict, Any, List

class ReportGenerator:
    """
    Generates comprehensive HTML analysis reports with charts and tables
    """
    
    def __init__(self):
        self.report_template = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Demographic Data Analysis Report</title>
    <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0;
            padding: 20px;
            background-color: #f5f5f5;
            color: #333;
        }
        .container {
            max-width: 1200px;
            margin: 0 auto;
            background-color: white;
            padding: 30px;
            border-radius: 10px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }
        .header {
            text-align: center;
            border-bottom: 3px solid #2E86AB;
            padding-bottom: 20px;
            margin-bottom: 30px;
        }
        .header h1 {
            color: #2E86AB;
            margin: 0;
            font-size: 2.5em;
        }
        .metadata {
            background-color: #f8f9fa;
            padding: 15px;
            border-radius: 8px;
            margin-bottom: 25px;
        }
        .section {
            margin-bottom: 40px;
        }
        .section h2 {
            color: #2E86AB;
            border-bottom: 2px solid #A23B72;
            padding-bottom: 10px;
            margin-bottom: 20px;
        }
        .metrics-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }
        .metric-card {
            background-color: #f8f9fa;
            padding: 20px;
            border-radius: 8px;
            text-align: center;
            border-left: 4px solid #F18F01;
        }
        .metric-card h3 {
            margin: 0 0 10px 0;
            color: #2E86AB;
        }
        .metric-card .value {
            font-size: 2em;
            font-weight: bold;
            color: #A23B72;
        }
        .chart-container {
            margin: 20px 0;
            background-color: white;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        .table-container {
            overflow-x: auto;
            margin: 20px 0;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            background-color: white;
            border-radius: 8px;
            overflow: hidden;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        th, td {
            padding: 12px;
            text-align: left;
            border-bottom: 1px solid #ddd;
        }
        th {
            background-color: #2E86AB;
            color: white;
            font-weight: bold;
        }
        tr:hover {
            background-color: #f5f5f5;
        }
        .summary-list {
            background-color: #f8f9fa;
            padding: 20px;
            border-radius: 8px;
            margin: 20px 0;
        }
        .summary-list ul {
            margin: 0;
            padding-left: 20px;
        }
        .summary-list li {
            margin-bottom: 8px;
        }
        .footer {
            text-align: center;
            margin-top: 40px;
            padding-top: 20px;
            border-top: 1px solid #ddd;
            color: #666;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Demographic Data Analysis Report</h1>
            <p>Generated on {{ report_date }}</p>
        </div>

        <div class="metadata">
            <h3>Processing Configuration</h3>
            <p><strong>Fuzzy Algorithm:</strong> {{ fuzzy_algorithm }}</p>
            <p><strong>Accuracy Threshold:</strong> {{ fuzzy_threshold }}%</p>
            <p><strong>Detection Method:</strong> {{ detection_method }}</p>
        </div>

        <div class="section">
            <h2>Executive Summary</h2>
            <div class="metrics-grid">
                <div class="metric-card">
                    <h3>Total Input Rows</h3>
                    <div class="value">{{ stats.original_columns_total }}</div>
                </div>
                <div class="metric-card">
                    <h3>Demographic Rows</h3>
                    <div class="value">{{ stats.demographic_rows_extracted }}</div>
                </div>
                <div class="metric-card">
                    <h3>Extraction Rate</h3>
                    <div class="value">{{ stats.extraction_percentage }}%</div>
                </div>
                <div class="metric-card">
                    <h3>Matched Records</h3>
                    <div class="value">{{ stats.matched_records }}</div>
                </div>
                <div class="metric-card">
                    <h3>Unique Tables</h3>
                    <div class="value">{{ stats.unique_tables }}</div>
                </div>
                <div class="metric-card">
                    <h3>Match Rate</h3>
                    <div class="value">{{ match_rate }}%</div>
                </div>
            </div>

            <div class="summary-list">
                <h3>Key Findings</h3>
                <ul>
                    <li>Processed {{ stats.original_columns_total }} total rows from input data</li>
                    <li>Successfully extracted {{ stats.demographic_rows_extracted }} demographic records ({{ stats.extraction_percentage }}% extraction rate)</li>
                    <li>{{ stats.matched_records }} records successfully matched with table information</li>
                    <li>Data spans across {{ stats.unique_tables }} unique tables</li>
                    <li>Overall match rate of {{ match_rate }}% achieved</li>
                </ul>
            </div>
        </div>

        <div class="section">
            <h2>Data Extraction Analysis</h2>
            <div class="chart-container">
                {{ extraction_chart }}
            </div>
        </div>



        <div class="section">
            <h2>Table Distribution</h2>
            <div class="chart-container">
                {{ table_distribution_chart }}
            </div>
        </div>

        <div class="section">
            <h2>Table Analysis Summary</h2>
            <div class="table-container">
                {{ table_analysis_table }}
            </div>
        </div>

        <div class="section">
            <h2>Algorithm Details</h2>
            <div class="algorithm-details">
                {{ algorithm_details }}
            </div>
        </div>

        {% if demographic_columns %}
        <div class="section">
            <h2>Demographic Data Types Found</h2>
            <div class="table-container">
                <table>
                    <thead>
                        <tr>
                            <th>#</th>
                            <th>Column Name</th>
                            <th>Data Type Category</th>
                        </tr>
                    </thead>
                    <tbody>
                        {% for idx, col in demographic_columns %}
                        <tr>
                            <td>{{ idx }}</td>
                            <td>{{ col }}</td>
                            <td>{{ col_categories.get(col, 'General Demographic') }}</td>
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>
        </div>
        {% endif %}

        <div class="section">
            <h2>Processing Statistics</h2>
            <div class="table-container">
                <table>
                    <thead>
                        <tr>
                            <th>Metric</th>
                            <th>Value</th>
                            <th>Description</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td>Total Input Rows</td>
                            <td>{{ stats.original_columns_total }}</td>
                            <td>Total number of rows in the uploaded columns data file</td>
                        </tr>
                        <tr>
                            <td>Demographic Rows Extracted</td>
                            <td>{{ stats.demographic_rows_extracted }}</td>
                            <td>Number of rows identified as containing demographic data</td>
                        </tr>
                        <tr>
                            <td>Non-Demographic Rows</td>
                            <td>{{ stats.non_demographic_rows }}</td>
                            <td>Number of rows that did not match demographic criteria</td>
                        </tr>
                        <tr>
                            <td>Final Processed Records</td>
                            <td>{{ final_records }}</td>
                            <td>Total records in the final processed dataset</td>
                        </tr>
                        <tr>
                            <td>Successfully Matched</td>
                            <td>{{ stats.matched_records }}</td>
                            <td>Records that were successfully matched with table information</td>
                        </tr>
                        <tr>
                            <td>Unmatched Records</td>
                            <td>{{ unmatched_records }}</td>
                            <td>Records that could not be matched with table information</td>
                        </tr>
                    </tbody>
                </table>
            </div>
        </div>

        <div class="section">
            <h2>Sample Data Preview</h2>
            <div class="table-container">
                {{ sample_data_table }}
            </div>
        </div>

        <div class="footer">
            <p>Report generated by Demographic Data Processing Application</p>
            <p>{{ report_date }}</p>
        </div>
    </div>
</body>
</html>
        """
    
    def generate_report(self, processed_data: pd.DataFrame, stats: Dict[str, Any], 
                       fuzzy_algorithm: str = "ratio", fuzzy_threshold: int = 80) -> str:
        """
        Generate comprehensive HTML analysis report
        
        Args:
            processed_data: The processed demographic data
            stats: Processing statistics
            fuzzy_algorithm: Algorithm used for fuzzy matching
            fuzzy_threshold: Threshold used for fuzzy matching
            
        Returns:
            HTML report as string
        """
        # Calculate additional metrics
        final_records = len(processed_data)
        unmatched_records = final_records - stats.get('matched_records', 0)
        match_rate = round((stats.get('matched_records', 0) / final_records * 100), 1) if final_records > 0 else 0
        
        # Determine detection method
        detection_method = "attr_description column analysis" if 'attr_description' in processed_data.columns else "Column name pattern matching"
        
        # Generate charts
        extraction_chart = self._create_extraction_chart(stats)
        matching_chart = self._create_matching_chart(stats, final_records)
        table_distribution_chart = self._create_table_distribution_chart(processed_data)
        
        # Get demographic columns information
        demographic_columns = self._get_demographic_columns_info(processed_data, stats)
        col_categories = self._categorize_columns(demographic_columns)
        
        # Create table analysis table
        table_analysis_table = self._create_table_analysis_table(processed_data)
        
        # Create algorithm details
        algorithm_details = self._create_algorithm_details(processed_data, fuzzy_algorithm, fuzzy_threshold)
        
        # Create sample data table
        sample_data_table = self._create_sample_data_table(processed_data)
        
        # Render template
        template = Template(self.report_template)
        
        report_html = template.render(
            report_date=datetime.now().strftime("%B %d, %Y at %I:%M %p"),
            stats=stats,
            fuzzy_algorithm=fuzzy_algorithm,
            fuzzy_threshold=fuzzy_threshold,
            detection_method=detection_method,
            match_rate=match_rate,
            final_records=final_records,
            unmatched_records=unmatched_records,
            extraction_chart=extraction_chart,
            matching_chart=matching_chart,
            table_distribution_chart=table_distribution_chart,
            table_analysis_table=table_analysis_table,
            algorithm_details=algorithm_details,
            demographic_columns=demographic_columns,
            col_categories=col_categories,
            sample_data_table=sample_data_table
        )
        
        return report_html
    
    def _create_extraction_chart(self, stats: Dict[str, Any]) -> str:
        """Create extraction analysis chart"""
        labels = ['Demographic Rows', 'Non-Demographic Rows']
        values = [
            stats.get('demographic_rows_extracted', 0),
            stats.get('non_demographic_rows', 0)
        ]
        colors = ['#2E86AB', '#A23B72']
        
        fig = go.Figure(data=[go.Pie(
            labels=labels,
            values=values,
            marker_colors=colors,
            hole=0.4,
            textinfo='label+percent+value',
            textfont_size=12
        )])
        
        fig.update_layout(
            title={
                'text': 'Data Extraction Breakdown',
                'x': 0.5,
                'xanchor': 'center',
                'font': {'size': 18}
            },
            height=400,
            margin=dict(t=60, b=20, l=20, r=20)
        )
        
        return pyo.plot(fig, output_type='div', include_plotlyjs=False)
    
    def _create_matching_chart(self, stats: Dict[str, Any], final_records: int) -> str:
        """Create matching performance chart"""
        matched = stats.get('matched_records', 0)
        unmatched = final_records - matched
        
        fig = go.Figure(data=[
            go.Bar(
                x=['Matched Records', 'Unmatched Records'],
                y=[matched, unmatched],
                marker_color=['#F18F01', '#C73E1D'],
                text=[f'{matched}', f'{unmatched}'],
                textposition='auto',
                textfont=dict(size=14, color='white')
            )
        ])
        
        fig.update_layout(
            title={
                'text': 'Storage ID Matching Results',
                'x': 0.5,
                'xanchor': 'center',
                'font': {'size': 18}
            },
            xaxis_title='Match Status',
            yaxis_title='Number of Records',
            height=400,
            margin=dict(t=60, b=60, l=60, r=20)
        )
        
        return pyo.plot(fig, output_type='div', include_plotlyjs=False)
    
    def _create_table_distribution_chart(self, processed_data: pd.DataFrame) -> str:
        """Create table distribution chart"""
        if 'table_name' not in processed_data.columns:
            return "<p>No table distribution data available</p>"
        
        table_counts = processed_data['table_name'].value_counts().head(10)
        
        fig = go.Figure(data=[
            go.Bar(
                x=table_counts.index,
                y=table_counts.values,
                marker_color='#2E86AB',
                text=table_counts.values,
                textposition='auto',
                textfont=dict(size=12, color='white')
            )
        ])
        
        fig.update_layout(
            title={
                'text': 'Top 10 Tables by Record Count',
                'x': 0.5,
                'xanchor': 'center',
                'font': {'size': 18}
            },
            xaxis_title='Table Name',
            yaxis_title='Number of Records',
            height=400,
            margin=dict(t=60, b=100, l=60, r=20),
            xaxis_tickangle=-45
        )
        
        return pyo.plot(fig, output_type='div', include_plotlyjs=False)
    
    def _create_table_analysis_table(self, processed_data: pd.DataFrame) -> str:
        """Create table analysis summary showing total fields and demographic data per table"""
        if 'table_name' not in processed_data.columns:
            return "<p>Table analysis not available - no table_name column found in data.</p>"
        
        # Group by table_name and count records
        table_stats = processed_data.groupby('table_name').agg({
            'attr_name': 'count',  # Total fields in each table
            'attr_description': lambda x: x.count()  # Non-null demographic descriptions
        }).reset_index()
        
        table_stats.columns = ['Table Name', 'Total Fields', 'Demographic Fields']
        
        # Calculate percentage of demographic fields per table
        table_stats['Demographic %'] = ((table_stats['Demographic Fields'] / table_stats['Total Fields']) * 100).round(1)
        
        # Sort by total fields descending
        table_stats = table_stats.sort_values('Total Fields', ascending=False)
        
        # Create HTML table
        html = """
        <table>
            <thead>
                <tr>
                    <th>Table Name</th>
                    <th>Total Fields</th>
                    <th>Demographic Fields</th>
                    <th>Demographic %</th>
                </tr>
            </thead>
            <tbody>
        """
        
        for _, row in table_stats.iterrows():
            html += f"""
                <tr>
                    <td>{row['Table Name']}</td>
                    <td>{row['Total Fields']}</td>
                    <td>{row['Demographic Fields']}</td>
                    <td>{row['Demographic %']}%</td>
                </tr>
            """
        
        html += """
            </tbody>
        </table>
        """
        
        return html
    
    def _create_algorithm_details(self, processed_data: pd.DataFrame, fuzzy_algorithm: str, fuzzy_threshold: int) -> str:
        """Create detailed algorithm explanation with examples"""
        
        # Algorithm descriptions
        algorithm_descriptions = {
            'ratio': 'Simple ratio comparison that calculates the overall similarity between two strings',
            'partial_ratio': 'Finds the best partial match within the strings, useful for substring matching',
            'token_sort_ratio': 'Sorts tokens alphabetically before comparison, ignoring word order',
            'token_set_ratio': 'Compares unique tokens between strings, ignoring duplicates and order'
        }
        
        algorithm_desc = algorithm_descriptions.get(fuzzy_algorithm, 'Advanced string similarity matching')
        
        # Create examples from actual data if available
        examples_html = ""
        if 'attr_description' in processed_data.columns:
            # Get some sample descriptions for examples
            sample_descriptions = processed_data['attr_description'].dropna().head(3).tolist()
            
            # Demographic keywords for examples
            demo_keywords = ['age', 'gender', 'race', 'ethnicity', 'sex', 'birth', 'demographic']
            
            examples_html = """
            <h4>Matching Examples from Your Data:</h4>
            <div class="examples-container">
            """
            
            for i, desc in enumerate(sample_descriptions[:2]):
                if len(str(desc)) > 10:  # Ensure it's a meaningful description
                    # Calculate similarity with a demographic keyword
                    from fuzzywuzzy import fuzz
                    best_match = ""
                    best_score = 0
                    
                    for keyword in demo_keywords:
                        if fuzzy_algorithm == 'ratio':
                            score = fuzz.ratio(str(desc).lower(), keyword)
                        elif fuzzy_algorithm == 'partial_ratio':
                            score = fuzz.partial_ratio(str(desc).lower(), keyword)
                        elif fuzzy_algorithm == 'token_sort_ratio':
                            score = fuzz.token_sort_ratio(str(desc).lower(), keyword)
                        else:  # token_set_ratio
                            score = fuzz.token_set_ratio(str(desc).lower(), keyword)
                        
                        if score > best_score:
                            best_score = score
                            best_match = keyword
                    
                    match_status = "✓ MATCH" if best_score >= fuzzy_threshold else "✗ NO MATCH"
                    color = "#28a745" if best_score >= fuzzy_threshold else "#dc3545"
                    
                    examples_html += f"""
                    <div class="example-item" style="border-left: 4px solid {color}; padding: 10px; margin: 10px 0; background: #f8f9fa;">
                        <strong>Text:</strong> "{str(desc)[:80]}..."<br>
                        <strong>Best Keyword Match:</strong> "{best_match}"<br>
                        <strong>Similarity Score:</strong> {best_score}% <span style="color: {color}; font-weight: bold;">({match_status})</span>
                    </div>
                    """
            
            examples_html += "</div>"
        
        html = f"""
        <div class="algorithm-info" style="background: #f8f9fa; padding: 20px; border-radius: 8px; margin: 20px 0;">
            <h3>Fuzzy Matching Algorithm: {fuzzy_algorithm.replace('_', ' ').title()}</h3>
            
            <div class="algorithm-description">
                <h4>How It Works:</h4>
                <p><strong>Algorithm:</strong> {algorithm_desc}</p>
                <p><strong>Threshold:</strong> {fuzzy_threshold}% (minimum similarity required for a match)</p>
                <p><strong>Process:</strong> The system analyzes the 'attr_description' column content and compares it against demographic keywords using fuzzy string matching. Any description with a similarity score above {fuzzy_threshold}% is classified as demographic data.</p>
            </div>
            
            <div class="keywords-list">
                <h4>Demographic Keywords Used:</h4>
                <p style="font-style: italic;">age, gender, race, ethnicity, sex, birth, demographic, population, ancestry, nationality, heritage, origin, background</p>
            </div>
            
            {examples_html}
            
            <div class="algorithm-benefits">
                <h4>Why Fuzzy Matching?</h4>
                <ul>
                    <li><strong>Handles Typos:</strong> Finds matches even with spelling errors</li>
                    <li><strong>Flexible Matching:</strong> Works with partial words and different formats</li>
                    <li><strong>Configurable Precision:</strong> Adjustable threshold allows fine-tuning sensitivity</li>
                    <li><strong>Context Aware:</strong> Considers word order and token relationships</li>
                </ul>
            </div>
        </div>
        """
        
        return html
    
    def _get_demographic_columns_info(self, processed_data: pd.DataFrame, stats: Dict[str, Any]) -> List[tuple]:
        """Get demographic columns information"""
        demo_cols = stats.get('demographic_column_names', [])
        if demo_cols:
            return [(i+1, col) for i, col in enumerate(demo_cols)]
        
        # If no specific demographic columns, show some key columns
        key_cols = [col for col in processed_data.columns 
                   if col not in ['storage_id', 'table_name', 'matched']][:10]
        return [(i+1, col) for i, col in enumerate(key_cols)]
    
    def _categorize_columns(self, demographic_columns: List[tuple]) -> Dict[str, str]:
        """Categorize demographic columns"""
        categories = {}
        
        for _, col_name in demographic_columns:
            col_lower = col_name.lower()
            if any(name_key in col_lower for name_key in ['name', 'embossed', 'primary', 'legal', 'dba']):
                categories[col_name] = 'Name Information'
            elif any(contact_key in col_lower for contact_key in ['phone', 'email', 'fax']):
                categories[col_name] = 'Contact Information'
            elif any(addr_key in col_lower for addr_key in ['address', 'location']):
                categories[col_name] = 'Address Information'
            elif any(id_key in col_lower for id_key in ['gender', 'dob', 'birth', 'gov', 'id']):
                categories[col_name] = 'Personal Demographics'
            elif any(pref_key in col_lower for pref_key in ['preference', 'language', 'member', 'since']):
                categories[col_name] = 'Preferences & Dates'
            else:
                categories[col_name] = 'General Demographic'
        
        return categories
    
    def _create_sample_data_table(self, processed_data: pd.DataFrame, max_rows: int = 10) -> str:
        """Create HTML table for sample data"""
        sample_df = processed_data.head(max_rows)
        
        # Limit columns to prevent wide tables
        display_cols = list(sample_df.columns)[:8]
        sample_df = sample_df[display_cols]
        
        html_table = sample_df.to_html(
            classes='table',
            table_id='sample-data',
            escape=False,
            index=False,
            border=0
        )
        
        return html_table
    
    def create_excel_export(self, processed_data: pd.DataFrame, stats: Dict[str, Any]) -> bytes:
        """
        Create Excel file with multiple sheets for export
        
        Args:
            processed_data: The processed demographic data
            stats: Processing statistics
            
        Returns:
            Excel file as bytes
        """
        output = io.BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Include all original columns from source file (table_name, attr_name, business_name, attr_description, etc.)
            original_data_only = processed_data[[col for col in processed_data.columns if col != 'matched']]
            
            # Main processed data with all original columns preserved
            original_data_only.to_excel(writer, sheet_name='Processed_Data', index=False)
            
            # Statistics summary
            stats_df = pd.DataFrame([
                ['Total Input Rows', stats.get('original_columns_total', 0)],
                ['Demographic Rows Extracted', stats.get('demographic_rows_extracted', 0)],
                ['Non-Demographic Rows', stats.get('non_demographic_rows', 0)],
                ['Final Processed Records', len(processed_data)],
                ['Successfully Matched', stats.get('matched_records', 0)],
                ['Unmatched Records', len(processed_data) - stats.get('matched_records', 0)],
                ['Extraction Percentage', f"{stats.get('extraction_percentage', 0)}%"],
                ['Unique Tables', stats.get('unique_tables', 0)]
            ], columns=['Metric', 'Value'])
            
            stats_df.to_excel(writer, sheet_name='Statistics', index=False)
            
            # Table distribution
            if 'table_name' in processed_data.columns:
                table_dist = processed_data['table_name'].value_counts().reset_index()
                table_dist.columns = ['Table Name', 'Record Count']
                table_dist.to_excel(writer, sheet_name='Table_Distribution', index=False)
            
            # Demographic columns info
            demo_cols = stats.get('demographic_column_names', [])
            if demo_cols:
                demo_df = pd.DataFrame(demo_cols, columns=['Demographic Column'])
                demo_df.to_excel(writer, sheet_name='Demographic_Columns', index=False)
        
        return output.getvalue()
    
    def create_csv_export(self, processed_data: pd.DataFrame) -> bytes:
        """
        Create CSV file for export with only original columns
        
        Args:
            processed_data: The processed demographic data
            
        Returns:
            CSV file as bytes
        """
        # Include all original columns from source file (table_name, attr_name, business_name, attr_description, etc.)
        original_data_only = processed_data[[col for col in processed_data.columns if col != 'matched']]
        
        output = io.StringIO()
        original_data_only.to_csv(output, index=False)
        return output.getvalue().encode('utf-8')
    
    def create_multiple_excel_files(self, processed_data: pd.DataFrame, records_per_file: int = None) -> List[tuple]:
        """
        Create multiple Excel files with demographic data split across them
        
        Args:
            processed_data: The processed demographic data
            records_per_file: Number of records per file (if None, splits into 20 files)
            
        Returns:
            List of tuples containing (filename, file_bytes)
        """
        if records_per_file is None:
            # Split into 20 files
            total_records = len(processed_data)
            records_per_file = max(1, total_records // 20)
            if total_records % 20 > 0:
                records_per_file += 1
        
        files = []
        total_records = len(processed_data)
        
        for i in range(0, total_records, records_per_file):
            file_num = (i // records_per_file) + 1
            chunk_data = processed_data.iloc[i:i + records_per_file].copy()
            
            if chunk_data.empty:
                continue
            
            # Create Excel file for this chunk
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # Include all original columns from source file (table_name, attr_name, business_name, attr_description, etc.)
                original_data_only = chunk_data[[col for col in chunk_data.columns if col != 'matched']]
                
                # Main data sheet with all original columns preserved
                original_data_only.to_excel(writer, sheet_name='Demographic_Data', index=False)
                
                # Summary sheet
                summary_data = {
                    'Metric': [
                        'Total Records in File',
                        'File Number',
                        'Records Range',
                        'Original Columns Preserved'
                    ],
                    'Value': [
                        len(chunk_data),
                        file_num,
                        f"{i+1} to {min(i + records_per_file, total_records)}",
                        len([col for col in chunk_data.columns if col not in ['table_name', 'matched']])
                    ]
                }
                
                summary_df = pd.DataFrame(summary_data)
                summary_df.to_excel(writer, sheet_name='Summary', index=False)
            
            filename = f"demographic_data_part_{file_num:02d}.xlsx"
            files.append((filename, output.getvalue()))
        
        return files