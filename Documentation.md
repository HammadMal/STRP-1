# Userinterface.py documentation: 

## Overview
A simplified PyQt6-based desktop application for file management and processing at Habib University. This application provides a clean, user-friendly interface for loading spreadsheet files and executing processing scripts.

## Features
- **File Loading**: Support for Excel (.xlsx, .xls) and CSV (.csv) files
- **Background Processing**: Non-blocking file loading with progress indication
- **File Validation**: Automatic validation of file content and format
- **External Script Integration**: Ready-to-implement processing script execution
- **University Branding**: Clean interface with Habib University color scheme

## Requirements
```bash
pip install PyQt6 pandas openpyxl
```

## File Structure
```
project/
├── UserInterface.py          # Main application file
└── README.md               # This documentation
```

## Usage

### Running the Application
```bash
python UserInterface.py
```

### Basic Workflow
1. **Browse Files**: Click "Browse Files" to select a spreadsheet file
2. **File Loading**: The application automatically validates and loads the file
3. **Process Files**: Once loaded successfully, click "Process Files" to run processing scripts

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
- `process_files()`: Executes processing scripts (to be implemented)

#### `FileProcessor`
Background thread class for non-blocking file processing.

**Key Methods:**
- `run()`: Processes the selected file and validates content
- **Signals:**
  - `finished(bool, str)`: Emitted when processing completes
  - `progress(int)`: Emitted to update progress bar

### UI Components
- **Title**: Application header with university branding
- **File Selection**: Browse button and selected file display
- **Process Button**: Main action button for script execution
- **Status Display**: Real-time feedback on operations
- **Progress Bar**: Visual indication during file loading

## Implementation Guide

### Adding Processing Scripts

To implement your own processing script, modify the `process_files()` method:

```python
def process_files(self):
    """Process the loaded file with external script."""
    if not self.current_file_path:
        return
    
    import subprocess
    try:
        # Replace 'your_script.py' with your actual script
        result = subprocess.run([
            'python', 
            'your_processing_script.py', 
            self.current_file_path
        ], capture_output=True, text=True, check=True)
        
        # Update status with success message
        self.status_label.setText(f"Processing completed: {result.stdout}")
        self.status_label.setStyleSheet("/* success styling */")
        
    except subprocess.CalledProcessError as e:
        # Handle errors
        self.status_label.setText(f"Processing failed: {e.stderr}")
        self.status_label.setStyleSheet("/* error styling */")
```

### Processing Script Template

Your processing script should accept the file path as a command-line argument:

```python
#!/usr/bin/env python3
import sys
import pandas as pd

def main():
    if len(sys.argv) != 2:
        print("Usage: python script.py <file_path>")
        sys.exit(1)
    
    file_path = sys.argv[1]
    
    # Load the file
    if file_path.endswith('.csv'):
        df = pd.read_csv(file_path)
    else:
        df = pd.read_excel(file_path)
    
    # Your processing logic here
    print(f"Processed {len(df)} rows successfully")

if __name__ == "__main__":
    main()
```

## Customization

### Colors and Styling
The application uses Habib University's color scheme:
- **Primary Purple**: `#6B2C91`
- **Hover Purple**: `#5A2478`
- **Pressed Purple**: `#4A1D63`
- **Background**: `#FFFFFF`

To modify colors, update the `setStyleSheet()` calls in the button definitions.

### Window Properties
Default window size and constraints can be modified in `init_ui()`:
```python
self.setMinimumSize(500, 350)  # Minimum width, height
self.resize(600, 450)          # Default width, height
```

## Error Handling

The application handles several error conditions:
- **Unsupported file formats**: Shows error message
- **Empty files**: Validates and rejects empty datasets
- **File reading errors**: Catches pandas exceptions
- **Processing errors**: Will display subprocess errors (when implemented)

## Status Messages

The application provides real-time feedback through color-coded status messages:
- **Gray**: Ready state
- **Yellow**: Processing in progress
- **Green**: Success
- **Red**: Error
- **Blue**: Information/Processing script status

## Thread Safety

File loading operations run in a separate thread (`FileProcessor`) to prevent UI freezing. The main thread communicates with the background thread through Qt signals.

## Future Enhancements

Potential improvements that can be added:
- Multiple file selection and batch processing
- Configuration file for processing script paths
- Processing history and logs
- Export functionality for processed results
- Progress indication for external script execution

## Troubleshooting

### Common Issues
1. **"No module named 'PyQt6'"**: Install PyQt6 using pip
2. **"No module named 'pandas'"**: Install pandas using pip
3. **Excel files not loading**: Install openpyxl for Excel support
4. **Process button stays disabled**: Check that file loaded successfully

### Debug Information
The application prints debug information to console:
- File paths during processing
- Processing script execution details

## License
This application is developed for Habib University internal use.