# interface.py documentation: 

## Overview
A simplified PyQt6-based desktop application for file management and processing at Habib University. This application provides a clean, user-friendly interface for loading spreadsheet files and executing the data.py processing script with integrated CLO/PLO calculation and Excel report generation.

## Features
- **File Loading**: Support for Excel (.xlsx, .xls) and CSV (.csv) files
- **Background Processing**: Non-blocking file loading with progress indication
- **File Validation**: Automatic validation of file content and format
- **External Script Integration**: Integrated with data.py for file processing
- **CLO/PLO Calculations**: Automated calculation of Course Learning Outcomes and Program Learning Outcomes
- **Excel Report Generation**: Creates formatted Excel reports with color-coded performance indicators
- **Grade Calculation**: Calculates final grades with letter grade assignments
- **University Branding**: Clean interface with Habib University color scheme
- **UTF-8 Support**: Handles Unicode characters in processing scripts

## Requirements
```bash
pip install PyQt6 pandas openpyxl
```

## File Structure
```
project/
├── interface.py              # Main application file (updated name)
├── data.py                   # Data processing script
├── clo_plo_calculator.py     # CLO/PLO calculation functions
├── excel_exporter.py         # Excel report generation
└── Documentation.md          # This documentation
```

## Usage

### Running the Application
```bash
python interface.py
```

### Basic Workflow
1. **Browse Files**: Click "Browse Files" to select a spreadsheet file
2. **File Loading**: The application automatically validates and loads the file
3. **Process Files**: Once loaded successfully, click "Process Files" to run data.py
4. **View Results**: CLO/PLO scores and final grades appear in terminal
5. **Excel Export**: Formatted Excel report is automatically generated

### Supported File Formats
- **CSV Files**: `.csv`
- **Excel Files**: `.xlsx`, `.xls`

## Code Structure

### Main Classes

#### `HabibUniversityApp`
The main application window that handles the user interface and user interactions.

**Key Methods:**
- `init_ui()`: Sets up the user interface components
- `browse_file()`: Opens file dialog for file selection
- `load_file()`: Initiates file loading process
- `process_files()`: Executes data.py processing script
- `_process_calculation_results()`: Handles CLO/PLO calculations and Excel generation

#### `FileProcessor`
Background thread class for non-blocking file loading and validation.

**Key Methods:**
- `run()`: Processes the selected file and validates content
- **Signals:**
  - `finished(bool, str)`: Emitted when processing completes
  - `progress(int)`: Emitted to update progress bar

#### `DataProcessor`
Background thread class for running data.py script without freezing the UI.

**Key Methods:**
- `run()`: Executes data.py with the selected file path
- **Signals:**
  - `finished(bool, str)`: Emitted when data processing completes
  - `progress(str)`: Emitted to update status during processing

### UI Components
- **Title**: Application header with university branding
- **File Selection**: Browse button and selected file display
- **Process Button**: Main action button for script execution
- **Status Display**: Real-time feedback on operations
- **Progress Bar**: Visual indication during file loading

## Integration with Processing Modules

### data.py Integration
The application is integrated with data.py which performs the following operations:
- Loads Excel files from the 'Data' sheet
- Cleans hidden characters and empty rows/columns
- Removes rows with less than 2 characters
- Filters columns with less than 70% missing values
- Extracts CLO, PLO, and student score data
- Outputs structured JSON data

### CLO/PLO Calculator Integration
The application uses `clo_plo_calculator.py` for:
- **CLO Score Calculation**: Weighted average of assessment scores per CLO
- **PLO Score Calculation**: Weighted mapping from CLO scores to PLO scores
- **Final Grade Calculation**: Overall course grade based on assessment weights
- **Letter Grade Assignment**: Converts numerical grades to letter grades (A+, A, B+, etc.)

### Excel Exporter Integration
The application uses `excel_exporter.py` for:
- **Formatted Excel Reports**: Professional-looking spreadsheets with university branding
- **Color-Coded Performance**: Visual indicators for student performance
- **Summary Analytics**: Statistical summaries of CLO/PLO performance
- **Grade Distribution**: Overview of class performance by grade ranges

## Output and Reports

### Terminal Output
The application displays comprehensive results in the terminal:
```
🎯 CLO Scores:
Ahmed Ali 1912: {'CLO 1': 100.0, 'CLO 5': 100.0, 'CLO 4': 100.0, 'CLO 0': 0.0}

📊 PLO Scores:
Ahmed Ali 1912: {'PLO 2': 100.0, 'PLO 3': 100.0}

🧮 Final Grades:
Ahmed Ali 1912: 100.00% (A+)

📌 Total CLO Weights:
CLO 1: 15.0 %
```

### Excel Report Features
The generated Excel file contains:

#### Main Results Sheet
- **Student ID Column**: All student identifiers
- **CLO Columns**: Individual CLO scores for each student
- **PLO Columns**: Calculated PLO scores
- **Overall Grade Column**: Final percentage and letter grade
- **Color Coding**: 
  - 🟢 **Green**: Scores ≥ 70% (Good performance)
  - 🟡 **Yellow**: Scores 60-69% (Needs attention)
  - 🔴 **Red**: Scores < 60% (Requires intervention)

#### Summary Sheet
- **CLO Performance Summary**: Average scores, students above/below thresholds
- **PLO Performance Summary**: Statistical analysis of PLO achievement
- **Grade Distribution**: Count of students in each letter grade category

#### Calculation Method Sheet
- **Documentation**: Explains the grade calculation methodology
- **Formula Details**: Transparent calculation process
- **Weighting Information**: How assessments contribute to final grades

## Status Messages

The application provides real-time feedback through color-coded status messages:
- **Gray**: Ready state
- **Yellow**: Processing in progress
- **Green**: Success
- **Red**: Error
- **Blue**: Information/Processing script status

## Error Handling

The application handles several error conditions:
- **Unsupported file formats**: Shows error message
- **Empty files**: Validates and rejects empty datasets
- **File reading errors**: Catches pandas exceptions
- **Processing errors**: Displays subprocess errors
- **Excel export errors**: Graceful fallback with terminal-only output
- **Encoding errors**: Handles UTF-8 encoding for special characters

## Thread Safety

- File loading operations run in a separate thread (`FileProcessor`) 
- Data processing runs in a separate thread (`DataProcessor`)
- Main thread communicates with background threads through Qt signals
- Prevents UI freezing during long operations
- Proper cleanup of background threads on application close

## Grade Calculation Details

### Assessment Weighting
The application uses the following assessment weights:
- **CLO 1**: 15% (Q1 assessment)
- **CLO 4**: 15% (Quiz 3 assessment) 
- **CLO 5**: 35% (Quiz 2 assessment)
- **CLO 0**: 0% (Bonus points)
- **Total Active Weight**: 65%

### Grade Scale
Using Habib University's standard grading scale:
- **A+**: 95-100%
- **A**: 90-94%
- **A-**: 85-89%
- **B+**: 80-84%
- **B**: 75-79%
- **B-**: 70-74%
- **C+**: 67-69%
- **C**: 63-66%
- **C-**: 60-62%
- **F**: Below 60%

## Customization

### Colors and Styling
The application uses Habib University's color scheme:
- **Primary Purple**: `#6B2C91`
- **Hover Purple**: `#5A2478`
- **Pressed Purple**: `#4A1D63`
- **Background**: `#FFFFFF`

### Performance Color Coding
Excel report color scheme:
- **Green (`#90EE90`)**: Scores ≥ 70%
- **Yellow (`#FFFF99`)**: Scores 60-69%
- **Red (`#FFB6C1`)**: Scores < 60%

### Performance Color Coding
Excel report color scheme:
- **Green (`#90EE90`)**: Scores ≥ 70%
- **Yellow (`#FFFF99`)**: Scores 60-69%
- **Red (`#FFB6C1`)**: Scores < 60%

### Window Properties
Default window size and constraints can be modified in `init_ui()`:
```python
self.setMinimumSize(500, 350)  # Minimum width, height
self.resize(600, 450)          # Default width, height
```

## Troubleshooting

### Common Issues and Solutions

1. **"No module named 'PyQt6'"**: 
   ```bash
   pip install PyQt6
   ```

2. **"No module named 'pandas'"**: 
   ```bash
   pip install pandas openpyxl
   ```

3. **"Module 'clo_plo_calculator' not found"**:
   - Ensure all Python files are in the same directory
   - Check that `clo_plo_calculator.py` exists

4. **"Module 'excel_exporter' not found"**:
   - Ensure `excel_exporter.py` is in the same directory
   - Check file permissions

5. **Encoding errors with special characters**:
   - The data.py script now forces UTF-8 encoding
   - Special characters are replaced with ASCII equivalents

6. **"Processing script 'data.py' not found"**:
   - Ensure data.py is in the same directory as interface.py

7. **Excel files not loading**: 
   - Install openpyxl: `pip install openpyxl`
   - Ensure the Excel file has a sheet named 'Data'

8. **Process button stays disabled**: 
   - Check that file loaded successfully
   - Check status messages for errors

9. **Excel export fails but calculations succeed**:
   - Check file permissions in the current directory
   - Ensure no other process has the output file open
   - Full calculation results are still available in terminal

### Debug Information
The application prints debug information to console:
- File paths during processing
- Full data.py output with structured JSON
- CLO/PLO calculation results
- Excel export status
- Error messages and stack traces

## Data Flow

1. **File Selection**: User selects Excel/CSV file through file dialog
2. **File Validation**: Background validation of file format and content
3. **Data Processing**: data.py extracts and cleans raw data
4. **JSON Parsing**: Structured data extracted from data.py output
5. **Score Calculation**: CLO, PLO, and grade calculations performed
6. **Terminal Display**: Results printed to console with emojis and formatting
7. **Excel Generation**: Formatted report created with color coding and summaries
8. **User Notification**: Success dialog with file path information

## Future Enhancements

Potential improvements that can be added:
- **Batch Processing**: Multiple file selection and processing
- **Configuration Management**: Customizable assessment weights and grading scales
- **Export Options**: PDF reports, CSV summaries
- **Data Visualization**: Charts and graphs for performance analysis
- **Historical Tracking**: Semester-over-semester performance comparison
- **Email Integration**: Automated report distribution
- **Custom Color Schemes**: User-configurable performance indicators
- **Template Management**: Multiple assessment templates
- **Advanced Analytics**: Predictive performance modeling

## License
This application is developed for Habib University internal use.

## Version History
- **v1.0**: Initial PyQt6 interface with basic file processing
- **v2.0**: Added CLO/PLO calculations and Excel export functionality
- **v2.1**: Updated color scheme and improved error handling
- **Current**: Enhanced documentation and streamlined grade calculation workflow
