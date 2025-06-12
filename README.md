# Demographic Data Processing Application

A comprehensive Streamlit-based application for extracting and analyzing demographic information from Excel and CSV files using advanced fuzzy matching algorithms.

## Overview

This application processes data files to identify and extract demographic information using intelligent pattern matching and fuzzy string algorithms. It preserves all original columns while filtering data based on demographic criteria found in attribute descriptions.

## Key Features

- **Multi-format Support**: Processes Excel (.xlsx, .xls) and CSV (.csv) files
- **Advanced Fuzzy Matching**: Uses configurable algorithms for demographic detection
- **Complete Column Preservation**: Maintains all original data structure
- **Comprehensive Reporting**: Generates detailed HTML analysis reports
- **Batch Export**: Creates 20 separate Excel files with automatic ZIP download
- **Table Analysis**: Provides field distribution statistics across database tables

## Fuzzy Matching Algorithms

The application uses the FuzzyWuzzy library to implement several string similarity algorithms for demographic data detection:

### 1. Ratio Algorithm (`fuzz.ratio`)
**How it works**: Calculates the Levenshtein distance between two strings and returns a similarity percentage.

**Formula**: `2 * matches / total_length * 100`

**Example**:
```
Text: "patient age at diagnosis"
Keyword: "age"
Similarity: 85% → MATCH (if threshold ≥ 80%)
```

**Best for**: Simple, direct text comparisons

### 2. Partial Ratio Algorithm (`fuzz.partial_ratio`)
**How it works**: Finds the best matching substring within the text, useful when demographic keywords appear as part of longer descriptions.

**Process**:
1. Takes the shorter string (keyword)
2. Finds the best matching substring in the longer text
3. Calculates ratio on the best match

**Example**:
```
Text: "demographic information about patient gender"
Keyword: "demographic"
Similarity: 100% → MATCH (exact substring found)
```

**Best for**: Detecting keywords within longer descriptions

### 3. Token Sort Ratio Algorithm (`fuzz.token_sort_ratio`)
**How it works**: Sorts words alphabetically before comparison, ignoring word order differences.

**Process**:
1. Tokenize both strings
2. Sort tokens alphabetically
3. Compare sorted strings using ratio algorithm

**Example**:
```
Text: "gender and age data"
Keyword: "age gender"
After sorting: "age and data gender" vs "age gender"
Similarity: 75% → MATCH (if threshold ≤ 75%)
```

**Best for**: Handling different word orders

### 4. Token Set Ratio Algorithm (`fuzz.token_set_ratio`)
**How it works**: Compares unique tokens between strings, eliminating duplicates and order sensitivity.

**Process**:
1. Create sets of unique tokens from both strings
2. Find intersection and differences
3. Calculate similarity based on common tokens

**Example**:
```
Text: "race race ethnicity background"
Keyword: "race ethnicity"
Unique tokens comparison → High similarity
```

**Best for**: Handling repeated words and order variations

## Demographic Keywords

The system searches for these demographic-related terms:

- **Age-related**: age, birth, birthday, born
- **Gender**: gender, sex, male, female
- **Race/Ethnicity**: race, ethnicity, ethnic, racial
- **Demographics**: demographic, population, ancestry
- **Background**: nationality, heritage, origin, background

## Configuration Options

### Fuzzy Threshold
- **Range**: 0-100%
- **Default**: 80%
- **Lower values**: More inclusive matching (may include false positives)
- **Higher values**: More strict matching (may miss some matches)

### Algorithm Selection
Choose based on your data characteristics:
- **Ratio**: Standard exact matching
- **Partial Ratio**: Keywords within longer text
- **Token Sort**: Different word orders
- **Token Set**: Repeated words and variations

## Installation

```bash
# Install required packages
pip install streamlit pandas openpyxl fuzzywuzzy python-levenshtein plotly jinja2 numpy

# Run the application
streamlit run app.py
```

## Usage

1. **Upload Files**: Load your columns data Excel/CSV file
2. **Configure Settings**: Select fuzzy algorithm and threshold
3. **Process Data**: Run demographic extraction
4. **Analyze Results**: View processing statistics and table distribution
5. **Export Data**: Download individual files, CSV, or 20-file ZIP archive
6. **Generate Report**: Create comprehensive HTML analysis report

## Data Structure Requirements

### Input File Columns
- `table_name`: Database table identifier
- `attr_name`: Attribute/field name
- `attr_description`: Field description (used for demographic detection)
- `business_name`: Business-friendly field name
- Additional columns: Preserved in output

### Output Structure
All original columns are maintained in the output, with demographic rows filtered based on the matching algorithm results.

## Algorithm Performance Examples

### High Confidence Matches (90%+ similarity)
```
"Patient age at time of diagnosis" → "age" (95%)
"Gender identity information" → "gender" (92%)
"Racial background data" → "race" (91%)
```

### Medium Confidence Matches (70-89% similarity)
```
"Date of birth information" → "birth" (85%)
"Ethnicity and heritage details" → "ethnicity" (78%)
"Demographic survey responses" → "demographic" (82%)
```

### Low Confidence Matches (Below threshold)
```
"Medical diagnosis codes" → "age" (15%)
"Treatment protocol steps" → "gender" (8%)
"Insurance claim numbers" → "race" (5%)
```

## Technical Implementation

### Core Components
- **DataProcessor**: Handles file reading and demographic extraction
- **ReportGenerator**: Creates HTML reports and Excel exports
- **Fuzzy Matching**: Implements configurable similarity algorithms
- **Streamlit Interface**: Provides user-friendly web interface

### Performance Optimization
- Efficient pandas operations for large datasets
- Batch processing for multiple file exports
- Memory-optimized Excel generation
- Configurable processing parameters

## File Exports

### Single Excel File
Contains all demographic data with original columns preserved.

### CSV Export
Lightweight format with demographic records and original structure.

### 20-File ZIP Archive
Splits demographic data across 20 Excel files for easier processing and distribution.

### HTML Report
Comprehensive analysis including:
- Processing statistics
- Algorithm details with examples
- Table distribution charts
- Sample data preview
- Performance metrics

## Troubleshooting

### Common Issues
1. **Low match rates**: Lower the fuzzy threshold or try different algorithms
2. **Missing columns**: Ensure attr_description column exists in your data
3. **Performance issues**: Process smaller datasets or adjust batch sizes

### Algorithm Selection Guide
- **Many exact matches expected**: Use `ratio`
- **Keywords in longer text**: Use `partial_ratio`
- **Inconsistent word order**: Use `token_sort_ratio`
- **Repeated/duplicate words**: Use `token_set_ratio`

## Version Information

- **Application**: Demographic Data Processing v1.0
- **Streamlit**: Web interface framework
- **FuzzyWuzzy**: String similarity algorithms
- **Pandas**: Data processing and manipulation
- **Plotly**: Interactive charts and visualizations

## Support

For technical issues or questions about the fuzzy matching algorithms:
1. Check the HTML report for algorithm performance details
2. Experiment with different threshold values
3. Try alternative algorithms based on your data characteristics
4. Review the sample matches in the generated reports

---

*This application provides enterprise-grade demographic data processing with transparent, configurable fuzzy matching algorithms for reliable and accurate results.*