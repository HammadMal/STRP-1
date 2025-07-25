# interface.py documentation: Enhanced with Batch Processing and Dynamic CLO Detection

## Overview
An enhanced PyQt6-based desktop application for file management and processing at Habib University. This application provides a clean, user-friendly interface for loading spreadsheet files (single or batch) and executing the data.py processing script with integrated CLO/PLO calculation and Excel report generation. **Now supports both single file processing and batch folder processing with intelligent CLO detection.**

## Features
- **Single File Processing**: Support for individual Excel (.xlsx, .xls) files
- **Batch Folder Processing**: Process multiple Excel files from a selected folder automatically
- **Dynamic CLO Detection**: Automatically identifies CLOs defined in course structure (no hardcoded CLO limits)
- **Intelligent Excel Output**: Only displays CLOs that are actually defined in the course
- **Dual Mode Interface**: Clear distinction between single file and batch processing modes
- **Real-time Progress Tracking**: Visual progress indicators for both file validation and processing
- **Background Processing**: Non-blocking file loading and processing with progress indication
- **File Validation**: Automatic validation of file content and format (individual and batch)
- **External Script Integration**: Integrated with data.py for file processing
- **CLO/PLO Calculations**: Automated calculation of Course Learning Outcomes and Program Learning Outcomes
- **Excel Report Generation**: Creates formatted Excel reports with color-coded performance indicators appended to original files
- **Grade Calculation**: Calculates final grades with letter grade assignments
- **Error Resilience**: In batch mode, individual file failures don't stop the entire process
- **Comprehensive Reporting**: Detailed success/failure reporting for batch operations
- **University Branding**: Clean interface with Habib University color scheme
- **UTF-8 Support**: Handles Unicode characters in processing scripts

## Requirements
```bash
pip install PyQt6 pandas openpyxl
```

## File Structure
```
project/
├── interface.py              # Main application file (enhanced with batch processing)
├── data.py                   # Data processing script (with dynamic CLO detection)
├── clo_plo_calculator.py     # CLO/PLO calculation functions
├── excel_exporter.py         # Excel report generation (dynamic CLO columns)
└── Documentation.md          # This documentation
```

## Key Enhancements in Latest Version

### Dynamic CLO Detection
- **Flexible CLO Structure**: No longer hardcoded to CLO 1-5; adapts to any number of CLOs
- **Course-Specific Output**: Excel reports only show CLOs that are actually defined in the course
- **Validation Logic**: Filters out invalid CLO entries and ensures proper course structure parsing
- **Comprehensive Debug Output**: Shows exactly which CLOs are found and used in processing

### Intelligent Excel Generation
- **Adaptive Columns**: Excel output automatically adjusts to show only relevant CLOs and PLOs
- **No Phantom CLOs**: Eliminates display of CLOs with zero scores that don't actually exist in the course
- **Debug Tracing**: Real-time feedback showing which CLOs are detected and included

## Usage

### Running the Application
```bash
python interface.py
```

### Basic Workflow - Single File Mode
1. **Select Single File**: Click "📄 Select Single File" to choose an individual Excel file
2. **File Loading**: The application automatically validates and loads the file
3. **CLO Detection**: System dynamically identifies all CLOs defined in the course structure
4. **Process File**: Once loaded successfully, click "🚀 Process Files" to run data.py
5. **View Results**: CLO/PLO scores and final grades appear in terminal with debug information
6. **Excel Export**: Formatted Excel report is automatically appended to the original file with only relevant CLOs

### Enhanced Workflow - Batch Processing Mode
1. **Select Folder**: Click "📁 Select Folder (Batch)" to choose a folder containing multiple Excel files
2. **File Discovery**: The application automatically scans the folder for all Excel files (.xlsx, .xls)
3. **File Validation**: Each Excel file is validated for compatibility and structure
4. **CLO Analysis**: Each file's CLO structure is analyzed independently
5. **Batch Processing**: Click "🚀 Process Files" to process all valid files in sequence
6. **Monitor Progress**: Watch real-time progress in the "Batch Progress" tab
7. **View Results**: Individual file results appear in real-time, with comprehensive terminal output
8. **Excel Export**: Each file gets its own CLO/PLO results sheet appended to the original file
9. **Summary Report**: Final summary dialog shows overall batch processing results

### Supported File Formats
- **Excel Files**: `.xlsx`, `.xls` (both single and batch mode)
- **CSV Files**: Not supported for result appending (Excel format required)

## Code Structure

### Main Classes

#### `HabibUniversityApp`
The main application window that handles the user interface and user interactions for both single and batch processing.

**Key Methods:**
- `init_ui()`: Sets up the enhanced user interface with dual mode support
- `browse_single_file()`: Opens file dialog for single file selection
- `browse_folder()`: Opens folder dialog for batch processing
- `load_file()`: Processes single selected file
- `load_folder()`: Validates all Excel files in selected folder
- `process_files()`: Routes to single or batch processing based on mode
- `_process_single_file()`: Executes data.py for single file
- `_process_batch_files()`: Executes batch processing for multiple files
- `_process_calculation_results()`: Handles CLO/PLO calculations and Excel generation

#### `FileProcessor`
Background thread class for non-blocking single file loading and validation.

**Key Methods:**
- `run()`: Processes the selected file and validates content
- **Signals:**
  - `finished(bool, str)`: Emitted when processing completes
  - `progress(int)`: Emitted to update progress bar

#### `BatchFileProcessor`
Background thread class for discovering and validating multiple Excel files in a folder.

**Key Methods:**
- `run()`: Scans folder for Excel files and validates each one
- **Signals:**
  - `finished(bool, str, list)`: Emitted when folder scanning completes
  - `progress(int)`: Emitted to update progress bar
  - `file_progress(str)`: Emitted to show current file being validated

#### `DataProcessor`
Background thread class for running data.py script on a single file without freezing the UI.

**Key Methods:**
- `run()`: Executes data.py with the selected file path
- **Signals:**
  - `finished(bool, str)`: Emitted when data processing completes
  - `progress(str)`: Emitted to update status during processing

#### `BatchDataProcessor`
Background thread class for processing multiple files in sequence.

**Key Methods:**
- `run()`: Processes all files in the batch sequentially
- `_process_single_file_results()`: Handles CLO/PLO calculations for each individual file
- **Signals:**
  - `finished(bool, str, dict)`: Emitted when batch processing completes
  - `progress(int)`: Emitted to update overall progress
  - `file_progress(str)`: Emitted to show current file being processed
  - `file_completed(str, bool, str)`: Emitted when each individual file completes

### Enhanced UI Components
- **Title**: Application header with university branding
- **Dual File Selection**: Separate buttons for single file and folder selection
- **Mode Indicator**: Clear display of current processing mode (single/batch)
- **Tabbed Status Display**: 
  - **Status Tab**: General application status and single file progress
  - **Batch Progress Tab**: Real-time batch processing progress and individual file results
- **Process Button**: Main action button for script execution (works for both modes)
- **Enhanced File Display**: Shows selected file(s) with appropriate formatting
- **Progress Bars**: Visual indication during file loading and processing
- **Scrollable Results**: Batch processing results with individual file status

## Integration with Processing Modules

### data.py Integration (Enhanced)
The application is integrated with data.py which performs the following operations:
- **Dynamic CLO Detection**: Scans Excel structure to identify all defined CLOs (not hardcoded)
- **Flexible Structure Parsing**: Adapts to different CLO numbering schemes (1-4, 1-5, 1-3, etc.)
- **Validation Logic**: Ensures CLO descriptions are meaningful and properly formatted
- **Course Structure Analysis**: Distinguishes between CLO definitions and PLO mappings
- Loads Excel files from the 'Data' sheet
- Cleans hidden characters and empty rows/columns
- Removes rows with less than 2 characters
- Filters columns with less than 70% missing values
- Extracts CLO, PLO, and student score data
- Outputs structured JSON data with comprehensive CLO definitions

### CLO/PLO Calculator Integration
The application uses `clo_plo_calculator.py` for:
- **CLO Score Calculation**: Weighted average of assessment scores per CLO
- **PLO Score Calculation**: Weighted mapping from CLO scores to PLO scores
- **Final Grade Calculation**: Overall course grade based on assessment weights
- **Letter Grade Assignment**: Converts numerical grades to letter grades (A+, A, B+, etc.)

### Excel Exporter Integration (Enhanced)
The application uses `excel_exporter.py` for:
- **Dynamic Column Generation**: Creates Excel columns only for CLOs that exist in the course
- **Intelligent CLO Detection**: Uses course definition data rather than hardcoded CLO lists
- **Formatted Excel Reports**: Professional-looking spreadsheets with university branding
- **Color-Coded Performance**: Visual indicators for student performance
- **Result Appending**: Adds results to original files rather than creating new files
- **Summary Analytics**: Statistical summaries of CLO/PLO performance
- **Grade Distribution**: Overview of class performance by grade ranges
- **Debug Information**: Detailed logging of which CLOs are included and why

## Output and Reports

### Enhanced Debug Output
The application now provides detailed debug information during processing:
```
🔍 CLOs found in course definition: ['CLO 1', 'CLO 2', 'CLO 3', 'CLO 4']
🔍 PLOs found in course mapping: ['PLO 1', 'PLO 2', 'PLO 3']
📝 Excluding CLO 0 from Excel output (bonus points)
📊 Final CLOs for Excel: ['CLO 1', 'CLO 2', 'CLO 3', 'CLO 4']
📊 Final PLOs for Excel: ['PLO 1', 'PLO 2', 'PLO 3']
```

### Terminal Output - Single File Mode
The application displays comprehensive results in the terminal for each file:
```
🎯 CLO Scores:
7724: {'CLO 1': 85.88, 'CLO 2': 90.53, 'CLO 3': 82.89, 'CLO 4': 85.11, 'CLO 0': 100.0}

📊 PLO Scores:
7724: {'PLO 1': 88.2, 'PLO 3': 82.89, 'PLO 2': 85.11}

🧮 Final Grades:
7724: 87.55% (A-)

📌 Total CLO Weights:
CLO 1: 2125.0 %
CLO 2: 2375.0 %
CLO 3: 2250.0 %
CLO 4: 2250.0 %
```

### Terminal Output - Batch Mode
For batch processing, results are organized by file:
```
==================================================
📁 Results for: Class_Section_A.xlsx
==================================================

🔍 CLOs found in course definition: ['CLO 1', 'CLO 2', 'CLO 3', 'CLO 4']
📊 Final CLOs for Excel: ['CLO 1', 'CLO 2', 'CLO 3', 'CLO 4']

🎯 CLO Scores:
Student 1: {'CLO 1': 85.0, 'CLO 2': 92.0, 'CLO 3': 78.0, 'CLO 4': 88.0}

📊 PLO Scores:
Student 1: {'PLO 1': 88.5, 'PLO 2': 85.0, 'PLO 3': 83.0}

🧮 Final Grades:
Student 1: 85.50% (A-)

✅ Results appended to: Class_Section_A.xlsx

==================================================
📁 Results for: Class_Section_B.xlsx
==================================================
[... continues for each file ...]

============================================================
📊 BATCH PROCESSING SUMMARY
============================================================

📊 Batch Processing Complete!

✅ Successfully processed: 8 files
❌ Failed: 2 files

Successful files:
• Class_Section_A.xlsx
• Class_Section_B.xlsx
• Class_Section_C.xlsx
[... etc ...]

Failed files:
• Corrupted_File.xlsx
• Invalid_Format.xlsx
```

### Excel Report Features (Enhanced)
Each processed file gets a new "CLO PLO Results" sheet with:

#### Main Results Sheet
- **Student ID Column**: All student identifiers
- **Dynamic CLO Columns**: Only CLOs that are actually defined in the course (excludes CLO 0)
- **Dynamic PLO Columns**: Only PLOs that are mapped in the course structure
- **Overall Grade Column**: Final percentage and letter grade
- **Adaptive Layout**: Column count varies based on actual course structure
- **Color Coding**: 
  - 🟢 **Green**: Scores ≥ 70% (Good performance)
  - 🟡 **Yellow**: Scores 60-69% (Needs attention)
  - 🔴 **Red**: Scores < 60% (Requires intervention)

## Enhanced Status Messages

The application provides real-time feedback through color-coded status messages:
- **Gray**: Ready state
- **Yellow**: Processing in progress (single file or batch)
- **Green**: Success (file loaded or processing complete)
- **Red**: Error (file issues or processing failures)
- **Blue**: Information/Processing script status

### CLO Detection Status Messages
- **CLO Analysis**: "🔍 CLOs found in course definition: ['CLO 1', 'CLO 2', 'CLO 3']"
- **PLO Mapping**: "🔍 PLOs found in course mapping: ['PLO 1', 'PLO 2']"
- **Excel Preparation**: "📊 Final CLOs for Excel: ['CLO 1', 'CLO 2', 'CLO 3']"

### Batch-Specific Status Messages
- **Folder Scanning**: "Scanning folder for Excel files..."
- **File Validation**: "Validating: filename.xlsx"
- **Batch Progress**: "Processing: filename.xlsx (3/10)"
- **Completion**: "✅ Batch processing completed!"

## Error Handling

The enhanced application handles several error conditions:

### Single File Mode
- **Unsupported file formats**: Shows error message
- **Empty files**: Validates and rejects empty datasets
- **Invalid CLO structure**: Handles courses with non-standard CLO definitions
- **Missing CLO descriptions**: Filters out invalid or incomplete CLO entries
- **File reading errors**: Catches pandas exceptions
- **Processing errors**: Displays subprocess errors
- **Excel export errors**: Graceful fallback with terminal-only output

### Batch Mode
- **No Excel files in folder**: Clear error message
- **Individual file failures**: Continues processing other files
- **Mixed CLO structures**: Each file's CLO structure is analyzed independently
- **Partial batch success**: Reports successful and failed files separately
- **Folder access errors**: Handles permission and path issues
- **Mixed file types**: Automatically filters for Excel files only

### Enhanced Error Resilience
- **Individual File Isolation**: In batch mode, one corrupted file doesn't stop the entire process
- **CLO Structure Validation**: Invalid CLO definitions don't crash the application
- **Detailed Error Reporting**: Specific error messages for each failed file
- **Graceful Degradation**: Processing continues even if Excel appending fails for some files
- **Progress Preservation**: Progress tracking continues even when individual files fail

## Thread Safety

- File loading operations run in separate threads (`FileProcessor`, `BatchFileProcessor`) 
- Data processing runs in separate threads (`DataProcessor`, `BatchDataProcessor`)
- Main thread communicates with background threads through Qt signals
- Prevents UI freezing during long operations (especially important for batch processing)
- Proper cleanup of background threads on application close
- Thread termination handling for safe application exit

## Grade Calculation Details

### Dynamic Assessment Weighting
The application now dynamically calculates weights based on the actual course structure:
- **Flexible CLO Weights**: Calculated based on actual assessments in the course
- **Adaptive Calculation**: No hardcoded CLO weightings; adapts to course design
- **CLO 0 Exclusion**: Bonus points (CLO 0) automatically excluded from final grades
- **Comprehensive Coverage**: All defined CLOs included in final grade calculation

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

### Window Properties
Default window size and constraints can be modified in `init_ui()`:
```python
self.setMinimumSize(700, 500)  # Minimum width, height (increased for enhanced UI)
self.resize(800, 600)          # Default width, height (increased for batch features)
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
   - Check that file(s) loaded successfully
   - Check status messages for errors

9. **Excel export fails but calculations succeed**:
   - Check file permissions in the current directory
   - Ensure no other process has the output file(s) open
   - Full calculation results are still available in terminal

### CLO Detection Specific Issues

10. **"No CLOs found in course structure"**:
    - Check that your Excel file has CLO definitions in the correct format
    - Ensure CLO descriptions are meaningful (longer than 10 characters)
    - Verify that cells start with "CLO" followed by a number

11. **"Wrong number of CLOs in Excel output"**:
    - Check the debug output to see which CLOs were detected
    - Verify that all CLO definitions have proper descriptions
    - CLO 0 is automatically excluded from Excel output (bonus points)

12. **"CLO descriptions appear wrong"**:
    - Ensure CLO definitions are in separate rows from PLO mappings
    - Check that there are no empty rows between CLO definitions
    - Verify the Excel file structure matches expected format

### Batch Processing Specific Issues

13. **"No Excel files found in the selected folder"**:
    - Ensure the folder contains .xlsx or .xls files
    - Check that files are not corrupted or password-protected

14. **Some files in batch fail to process**:
    - This is normal - check the Batch Progress tab for specific errors
    - Failed files are reported separately and don't stop the batch
    - Different files may have different CLO structures (this is supported)

15. **Batch processing takes a long time**:
    - This is expected for large numbers of files
    - Monitor progress in the Batch Progress tab
    - Individual file progress is shown in real-time

16. **Memory issues with large batches**:
    - Consider processing smaller batches (50-100 files at a time)
    - Close other applications to free up memory
    - Large files (>50MB each) may require additional system resources

## Enhanced Data Flow

### Single File Mode (Enhanced)
1. **File Selection**: User selects Excel file through file dialog
2. **File Validation**: Background validation of file format and content
3. **CLO Structure Analysis**: Dynamic identification of all defined CLOs in the course
4. **Data Processing**: data.py extracts and cleans raw data with flexible CLO detection
5. **JSON Parsing**: Structured data extracted from data.py output
6. **Score Calculation**: CLO, PLO, and grade calculations performed for detected CLOs
7. **Terminal Display**: Results printed to console with emojis, formatting, and debug info
8. **Excel Generation**: Formatted report appended to original file with only relevant CLOs
9. **User Notification**: Success dialog with file path information

### Batch Mode (Enhanced)
1. **Folder Selection**: User selects folder containing multiple Excel files
2. **File Discovery**: Background scanning for all Excel files in folder
3. **Batch Validation**: Each file validated individually
4. **Individual CLO Analysis**: Each file's CLO structure analyzed independently
5. **Batch Processing**: Sequential processing of all valid files
6. **Individual Processing**: Each file goes through the complete single-file workflow
7. **Progress Tracking**: Real-time updates for overall progress and current file
8. **Result Aggregation**: Success/failure tracking for each file with CLO-specific details
9. **Batch Summary**: Comprehensive summary dialog and terminal output

## Performance Considerations

### Single File Processing
- **Typical Processing Time**: 2-10 seconds per file (depending on size and complexity)
- **CLO Detection Overhead**: Minimal additional time for dynamic CLO analysis
- **Memory Usage**: Low - single file processed at a time
- **UI Responsiveness**: Non-blocking background processing

### Batch Processing
- **Processing Time**: Scales linearly with number of files (2-10 seconds × number of files)
- **CLO Structure Independence**: Each file analyzed independently for optimal accuracy
- **Memory Usage**: Moderate - processes files sequentially to manage memory
- **Progress Feedback**: Real-time progress updates prevent perceived freezing
- **Error Isolation**: Individual file failures don't impact overall batch

### Optimization Tips
- **File Organization**: Group similar files in folders for efficient batch processing
- **System Resources**: Close unnecessary applications when processing large batches
- **File Size**: Extremely large files (>100MB) may require additional processing time
- **Network Drives**: Local file processing is faster than network drives
- **CLO Consistency**: Files with similar CLO structures process slightly faster

## Future Enhancements

Potential improvements that can be added:
- **CLO Template Detection**: Automatic detection of different university CLO formats
- **Custom CLO Mapping**: User interface for manual CLO structure configuration
- **CLO Validation Rules**: Advanced validation for CLO description quality and completeness
- **PLO Mapping Validation**: Ensure all CLOs have proper PLO mappings
- **Parallel Processing**: Process multiple files simultaneously for faster batch operations
- **File Filtering**: Options to filter files by date, size, or naming patterns
- **Resume Capability**: Ability to resume interrupted batch processing
- **Advanced Progress**: Estimated time remaining for batch operations
- **Export Options**: Batch export to consolidated reports (PDF, CSV summaries)
- **Configuration Management**: Customizable assessment weights and grading scales
- **Template Management**: Multiple assessment templates for different courses
- **Historical Tracking**: Semester-over-semester performance comparison
- **Email Integration**: Automated report distribution
- **Advanced Analytics**: Predictive performance modeling across multiple files
- **File Comparison**: Compare results across different batches or semesters

## License
This application is developed for Habib University internal use.

## Version History
- **v1.0**: Initial PyQt6 interface with basic file processing
- **v2.0**: Added CLO/PLO calculations and Excel export functionality
- **v2.1**: Updated color scheme and improved error handling
- **v3.0**: **Enhanced with comprehensive batch processing capabilities**
  - Added folder selection and batch file processing
  - Implemented dual-mode interface (single/batch)
  - Added real-time progress tracking for batch operations
  - Enhanced error handling and reporting for batch mode
  - Improved UI with tabbed status display and batch progress monitoring
  - Added comprehensive batch summary and reporting features
- **v3.1**: **Dynamic CLO Detection and Intelligent Excel Output**
  - Implemented dynamic CLO detection (no longer hardcoded to CLO 1-5)
  - Enhanced Excel output to show only CLOs that exist in the course structure
  - Added comprehensive debug output for CLO detection process
  - Improved validation logic for CLO descriptions and course structure
  - Enhanced data.py with flexible CLO parsing capabilities
  - Updated excel_exporter.py with intelligent column generation
- **Current**: Enhanced documentation and streamlined grade calculation workflow with dynamic CLO detection and batch processing support
