#!/usr/bin/env python3
"""
Habib University CLO/PLO Mapping UI - Simplified Version
A basic PyQt6 application for file management and processing.
"""

import sys
import os
from pathlib import Path
from typing import Optional
import pandas as pd

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QMessageBox, QProgressBar
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal


class FileProcessor(QThread):
    """Background thread for processing files."""
    
    finished = pyqtSignal(bool, str)  # success, message
    progress = pyqtSignal(int)  # progress percentage
    
    def __init__(self, file_path: str):
        super().__init__()
        self.file_path = file_path
        
    def run(self):
        """Process the uploaded file."""
        try:
            self.progress.emit(25)
            
            file_ext = Path(self.file_path).suffix.lower()
            
            self.progress.emit(50)
            
            # Load file based on extension
            if file_ext == '.csv':
                df = pd.read_csv(self.file_path)
            elif file_ext in ['.xlsx', '.xls']:
                df = pd.read_excel(self.file_path)
            else:
                self.finished.emit(False, "Unsupported file format")
                return
            
            self.progress.emit(75)
            
            # Validate file content
            if df.empty:
                self.finished.emit(False, "File is empty")
                return
            
            self.progress.emit(100)
            
            rows, cols = df.shape
            message = f"File loaded successfully!\nRows: {rows}, Columns: {cols}"
            self.finished.emit(True, message)
            
        except Exception as e:
            self.finished.emit(False, f"Error processing file: {str(e)}")


class HabibUniversityApp(QMainWindow):
    """Main application window."""
    
    def __init__(self):
        super().__init__()
        self.current_file_path: Optional[str] = None
        self.file_processor: Optional[FileProcessor] = None
        self.init_ui()
    
    def init_ui(self):
        """Initialize the user interface."""
        self.setWindowTitle("Habib University - CLO/PLO Mapping")
        self.setMinimumSize(500, 350)
        self.resize(600, 450)
        
        # Create central widget
        central_widget = QWidget()
        central_widget.setStyleSheet("background-color: #FFFFFF;")
        self.setCentralWidget(central_widget)
        
        # Main layout
        layout = QVBoxLayout(central_widget)
        layout.setSpacing(20)
        layout.setContentsMargins(30, 30, 30, 30)
        
        # Title
        title = QLabel("Habib University CLO/PLO Mapping Tool")
        title.setStyleSheet("font-size: 24px; font-weight: bold; color: #6B2C91;")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("font-size: 18px; font-weight: bold; color: #333;")
        layout.addWidget(title)
        
        # File selection section
        file_layout = QHBoxLayout()
        
        self.file_label = QLabel("No file selected")
        self.file_label.setStyleSheet("padding: 8px; border: 1px solid #ccc; background: #f9f9f9;")
        
        self.browse_btn = QPushButton("Browse Files")
        self.browse_btn.clicked.connect(self.browse_file)
        self.browse_btn.setMinimumWidth(120)
        self.browse_btn.setStyleSheet("""
            QPushButton {
                background-color: #6B2C91;
                color: white;
                border: none;
                border-radius: 8px;
                padding: 8px 16px;
                font-weight: bold;
                min-height: 25px;
            }
            QPushButton:hover {
                background-color: #5A2478;
            }
            QPushButton:pressed {
                background-color: #4A1D63;
            }
            QPushButton:disabled {
                background-color: #CCCCCC;
                color: #666666;
            }
        """)
        
        file_layout.addWidget(self.file_label, 1)
        file_layout.addWidget(self.browse_btn)
        layout.addLayout(file_layout)
        
        # Supported formats info
        info = QLabel("Supported: Excel (.xlsx, .xls) and CSV (.csv)")
        info.setStyleSheet("color: #666; font-size: 12px;")
        layout.addWidget(info)
        
        # Process files button
        self.process_btn = QPushButton("Process Files")
        self.process_btn.clicked.connect(self.process_files)
        self.process_btn.setEnabled(False)  # Disabled until file is selected
        self.process_btn.setStyleSheet("""
            QPushButton {
                background-color: #6B2C91;
                color: white;
                border: none;
                border-radius: 8px;
                padding: 12px 16px;
                font-weight: bold;
                min-height: 30px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #5A2478;
            }
            QPushButton:pressed {
                background-color: #4A1D63;
            }
            QPushButton:disabled {
                background-color: #CCCCCC;
                color: #666666;
            }
        """)
        layout.addWidget(self.process_btn)
        
        # Status label
        self.status_label = QLabel("Ready")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status_label.setStyleSheet("padding: 12px; border: 1px solid #ddd; background: #f5f5f5;")
        layout.addWidget(self.status_label)
        
        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)
        
        # Add stretch to center content
        layout.addStretch()
    
    def browse_file(self):
        """Open file dialog to select a file."""
        file_path, _ = QFileDialog.getOpenFileName(
            self, 
            "Select File - Habib University",
            "",
            "Spreadsheet files (*.xlsx *.xls *.csv);;All files (*.*)"
        )
        
        if file_path:
            self.load_file(file_path)
    
    def load_file(self, file_path: str):
        """Load and process the selected file."""
        self.current_file_path = file_path
        file_name = Path(file_path).name
        
        # Update UI
        self.file_label.setText(f"Selected: {file_name}")
        self.status_label.setText("Processing file...")
        self.status_label.setStyleSheet("padding: 12px; border: 1px solid #ffc107; background: #fff3cd; color: #856404;")
        
        # Show progress bar
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        
        # Disable buttons
        self.browse_btn.setEnabled(False)
        self.process_btn.setEnabled(False)
        
        # Start processing
        self.file_processor = FileProcessor(file_path)
        self.file_processor.finished.connect(self.on_file_processed)
        self.file_processor.progress.connect(self.progress_bar.setValue)
        self.file_processor.start()
    
    def on_file_processed(self, success: bool, message: str):
        """Handle file processing completion."""
        # Hide progress bar
        self.progress_bar.setVisible(False)
        
        # Re-enable browse button
        self.browse_btn.setEnabled(True)
        
        # Update status
        if success:
            self.status_label.setText(message)
            self.status_label.setStyleSheet("padding: 12px; border: 1px solid #28a745; background: #d4edda; color: #155724;")
            # Enable process button after successful file load
            self.process_btn.setEnabled(True)
        else:
            self.status_label.setText(f"Error: {message}")
            self.status_label.setStyleSheet("padding: 12px; border: 1px solid #dc3545; background: #f8d7da; color: #721c24;")
            # Keep process button disabled on error
            self.process_btn.setEnabled(False)
        
        # Clean up
        if self.file_processor:
            self.file_processor.deleteLater()
            self.file_processor = None
    
    def process_files(self):
        """Process the loaded file with external script."""
        if not self.current_file_path:
            self.status_label.setText("No file selected for processing")
            self.status_label.setStyleSheet("padding: 12px; border: 1px solid #dc3545; background: #f8d7da; color: #721c24;")
            return
        
        # TODO: Implement the actual processing script here
        # For now, just show a placeholder message
        self.status_label.setText("Processing script will be implemented here...")
        self.status_label.setStyleSheet("padding: 12px; border: 1px solid #17a2b8; background: #d1ecf1; color: #0c5460;")
        
        # Placeholder for future script execution:
        # import subprocess
        # try:
        #     result = subprocess.run(['python', 'your_processing_script.py', self.current_file_path], 
        #                           capture_output=True, text=True, check=True)
        #     self.status_label.setText(f"Processing completed: {result.stdout}")
        # except subprocess.CalledProcessError as e:
        #     self.status_label.setText(f"Processing failed: {e.stderr}")
        
        print(f"Processing file: {self.current_file_path}")
    
    def closeEvent(self, event):
        """Handle application close."""
        if self.file_processor and self.file_processor.isRunning():
            self.file_processor.terminate()
            self.file_processor.wait()
        event.accept()


def main():
    """Main application entry point."""
    app = QApplication(sys.argv)
    app.setApplicationName("Habib University CLO/PLO Mapping Tool")
    app.setStyle("Fusion")
    
    window = HabibUniversityApp()
    window.show()
    
    sys.exit(app.exec())


if __name__ == "__main__":
    main()