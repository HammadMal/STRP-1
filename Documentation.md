# Userinterface.py documentation: 

## Overview
A simplified PyQt6-based desktop application for file management and processing at Habib University. This application provides a clean, user-friendly interface for loading spreadsheet files and executing the data.py processing script.

## Features
- **File Loading**: Support for Excel (.xlsx, .xls) and CSV (.csv) files
- **Background Processing**: Non-blocking file loading with progress indication
- **File Validation**: Automatic validation of file content and format
- **External Script Integration**: Integrated with data.py for file processing
- **University Branding**: Clean interface with Habib University color scheme
- **UTF-8 Support**: Handles Unicode characters in processing scripts

## Requirements
```bash
pip install PyQt6 pandas openpyxl
```

## File Structure
```
project/
├── UserInterface.py          # Main application file
├── data.py                   # Data processing script
└── Documentation.md          # This documentation
```

## Usage

### Running the Application
```bash
python UserInterface.py
```

### Basic Workflow
1. **Browse Files**: Click "Browse Files" to select a spreadsheet file
2. **File Loading**: The application automatically validates and loads the file
3. **Process Files**: Once loaded successfully, click "Process Files" to run data.py

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

## Integration with data.py

The application is integrated with data.py which performs the following operations:
- Loads Excel files from the 'Data' sheet
- Cleans hidden characters and empty rows/columns
- Removes rows with less than 2 characters
- Filters columns with less than 70% missing values
- Outputs the cleaned dataframe

### data.py Command Line Usage
The data.py script can be used in multiple ways:

```bash
# With UI (automatic)
# The UI passes the file path automatically

# Standalone with default path
python data.py

# With custom file path
python data.py "path/to/your/file.xlsx"
```

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
- **Encoding errors**: Handles UTF-8 encoding for special characters

## Thread Safety

- File loading operations run in a separate thread (`FileProcessor`) 
- Data processing runs in a separate thread (`DataProcessor`)
- Main thread communicates with background threads through Qt signals
- Prevents UI freezing during long operations

## Output Display

- Processing output is displayed in the status area
- Long outputs (>10 lines) are truncated in the UI for readability
- Full output is always printed to the console
- Console output is prefixed with "=== Data Processing Output ==="

## Customization

### Colors and Styling
The application uses Habib University's color scheme:
- **Primary Purple**: `#6B2C91`
- **Hover Purple**: `#5A2478`
- **Pressed Purple**: `#4A1D63`
- **Background**: `#FFFFFF`

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

3. **Encoding errors with special characters**:
   - The data.py script now forces UTF-8 encoding
   - Special characters are replaced with ASCII equivalents

4. **"Processing script 'data.py' not found"**:
   - Ensure data.py is in the same directory as UserInterface.py

5. **Excel files not loading**: 
   - Install openpyxl: `pip install openpyxl`
   - Ensure the Excel file has a sheet named 'Data'

6. **Process button stays disabled**: 
   - Check that file loaded successfully
   - Check status messages for errors

### Debug Information
The application prints debug information to console:
- File paths during processing
- Full data.py output
- Error messages and stack traces

## Future Enhancements

Potential improvements that can be added:
- Multiple file selection and batch processing
- Configuration file for processing script paths
- Processing history and logs
- Export functionality for processed results
- Progress indication for data.py execution
- Custom sheet name selection for Excel files
- Output file saving functionality

## License
This application is developed for Habib University internal use.
