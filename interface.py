#!/usr/bin/env python3
"""
Habib University CLO/PLO Mapping UI - Clean Version
A PyQt6 application for file management, processing, and Excel report generation.
"""

import sys
import os
from pathlib import Path
from typing import Optional
import pandas as pd
import subprocess
import json
from clo_plo_calculator import (
    calculate_clo_scores,
    calculate_plo_scores,
    calculate_grades,
    get_letter_grade,
    get_total_clo_weights
)
from excel_exporter import export_clo_plo_results

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


class DataProcessor(QThread):
    """Background thread for running data.py script."""
    
    finished = pyqtSignal(bool, str)  # success, message
    progress = pyqtSignal(str)  # progress updates
    
    def __init__(self, file_path: str, script_path: str = "data.py"):
        super().__init__()
        self.file_path = file_path
        self.script_path = script_path
        
    def run(self):
        """Run the data.py script with the file path."""
        try:
            self.progress.emit("Starting data processing...")
            
            # Check if data.py exists
            if not os.path.exists(self.script_path):
                self.finished.emit(False, f"Processing script '{self.script_path}' not found")
                return
            
            # Run the script directly with the file path as argument
            result = subprocess.run(
                [sys.executable, self.script_path, self.file_path],
                capture_output=True,
                text=True,
                check=True,
                encoding='utf-8'  # Handle UTF-8 encoding
            )
            
            # Extract output
            output = result.stdout if result.stdout else "Processing completed successfully"
            self.finished.emit(True, output)
            
        except subprocess.CalledProcessError as e:
            error_msg = e.stderr if e.stderr else str(e)
            self.finished.emit(False, f"Processing failed: {error_msg}")
            
        except Exception as e:
            self.finished.emit(False, f"Unexpected error: {str(e)}")


class HabibUniversityApp(QMainWindow):
    """Main application window."""
    
    def __init__(self):
        super().__init__()
        self.current_file_path: Optional[str] = None
        self.file_processor: Optional[FileProcessor] = None
        self.data_processor: Optional[DataProcessor] = None
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
        title.setStyleSheet("font-size: 18px; font-weight: bold; color: #333;")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)
        
        # File selection section
        file_layout = QHBoxLayout()
        
        self.file_label = QLabel("No file selected")
        self.file_label.setStyleSheet("padding: 8px; border: 1px solid #ccc; background: #f9f9f9;")
        
        self.browse_btn = QPushButton("Browse Files")
        self.browse_btn.clicked.connect(self.browse_file)
        self.browse_btn.setMinimumWidth(120)
        self.browse_btn.setStyleSheet(self._get_button_style())
        
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
        self.process_btn.setStyleSheet(self._get_button_style())
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
    
    def _get_button_style(self):
        """Get consistent button styling."""
        return """
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
        """
    
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
        self._update_status("Processing file...", "processing")
        
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
            self._update_status(message, "success")
            self.process_btn.setEnabled(True)
        else:
            self._update_status(f"Error: {message}", "error")
            self.process_btn.setEnabled(False)
        
        # Clean up
        if self.file_processor:
            self.file_processor.deleteLater()
            self.file_processor = None
    
    def process_files(self):
        """Process the loaded file with data.py script."""
        if not self.current_file_path:
            self._update_status("No file selected for processing", "error")
            return
        
        # Update UI to show processing state
        self._update_status("Running data.py processing...", "processing")
        
        # Disable buttons during processing
        self.browse_btn.setEnabled(False)
        self.process_btn.setEnabled(False)
        
        # Start data processing in background thread
        self.data_processor = DataProcessor(self.current_file_path)
        self.data_processor.finished.connect(self.on_data_processed)
        self.data_processor.progress.connect(self._update_status)
        self.data_processor.start()
    
    def on_data_processed(self, success: bool, message: str):
        """Handle data processing completion."""
        # Re-enable buttons
        self.browse_btn.setEnabled(True)
        self.process_btn.setEnabled(True)
        
        if not success:
            self._update_status(f"Processing failed: {message}", "error")
            self._cleanup_processor()
            return
        
        # Print full output to console
        print("\n=== Data Processing Output ===")
        print(message)
        print("==============================\n")
        
        # Process the results
        try:
            self._process_calculation_results(message)
        except Exception as e:
            self._update_status(f"Calculation failed: {str(e)}", "error")
        
        self._cleanup_processor()
    
    def _process_calculation_results(self, message: str):
            """Process CLO/PLO calculations and create Excel output."""
            # Extract JSON block from message
            json_start = message.find("{")
            if json_start == -1:
                raise ValueError("No JSON found in output")

            json_data = message[json_start:]
            data_dict = json.loads(json_data)

            # Calculate scores
            clo_scores = calculate_clo_scores(data_dict["clo_assessments"], data_dict["student_scores"])
            plo_scores = calculate_plo_scores(clo_scores, data_dict["clo_to_plo"])
            grades = calculate_grades(data_dict["clo_assessments"], data_dict["student_scores"])
            clo_weights = get_total_clo_weights(data_dict["clo_assessments"])

            # Print to terminal
            self._print_scores_to_console(clo_scores, plo_scores, grades, clo_weights)

            # Create Excel output (now passing grades as parameter)
            try:
                output_file = export_clo_plo_results(clo_scores, plo_scores, grades, data_dict)
                self._update_status(f"CLO/PLO calculation complete. Excel file created: {output_file}", "success")
                self._show_success_dialog(output_file)
            except Exception as excel_error:
                print(f"\n❌ Excel export failed: {excel_error}")
                self._update_status("CLO/PLO calculation complete. See terminal output. (Excel export failed)", "warning")

    def _print_scores_to_console(self, clo_scores, plo_scores, grades, clo_weights):
        """Print CLO, PLO, Grades, and CLO weights to terminal."""
        print("\n🎯 CLO Scores:")
        for student, scores in clo_scores.items():
            print(f"{student}: {scores}")

        print("\n📊 PLO Scores:")
        for student, scores in plo_scores.items():
            print(f"{student}: {scores}")

        print("\n🧮 Final Grades:")
        for student, percent in grades.items():
            letter = get_letter_grade(percent)
            print(f"{student}: {percent:.2f}% ({letter})")

        print("\n📌 Total CLO Weights:")
        for clo, weight in clo_weights.items():
            print(f"{clo}: {weight} %")

    
    def _show_success_dialog(self, output_file: str):
        """Show success dialog with file path."""
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Icon.Information)
        msg.setWindowTitle("Export Complete")
        msg.setText(f"Results exported successfully to:\n{output_file}")
        msg.setStandardButtons(QMessageBox.StandardButton.Ok)
        msg.exec()
    
    def _update_status(self, message: str, status_type: str = "info"):
        """Update status label with appropriate styling."""
        styles = {
            "success": "padding: 12px; border: 1px solid #28a745; background: #d4edda; color: #155724;",
            "error": "padding: 12px; border: 1px solid #dc3545; background: #f8d7da; color: #721c24;",
            "warning": "padding: 12px; border: 1px solid #ffc107; background: #fff3cd; color: #856404;",
            "processing": "padding: 12px; border: 1px solid #ffc107; background: #fff3cd; color: #856404;",
            "info": "padding: 12px; border: 1px solid #ddd; background: #f5f5f5;"
        }
        
        self.status_label.setText(message)
        self.status_label.setStyleSheet(styles.get(status_type, styles["info"]))
    
    def _cleanup_processor(self):
        """Clean up data processor thread."""
        if self.data_processor:
            self.data_processor.deleteLater()
            self.data_processor = None
    
    def closeEvent(self, event):
        """Handle application close."""
        if self.file_processor and self.file_processor.isRunning():
            self.file_processor.terminate()
            self.file_processor.wait()
        if self.data_processor and self.data_processor.isRunning():
            self.data_processor.terminate()
            self.data_processor.wait()
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