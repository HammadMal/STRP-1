#!/usr/bin/env python3
"""
Habib University CLO/PLO Mapping UI - Enhanced with Batch Processing and Student ID Validation
A PyQt6 application for file management, processing, and Excel report generation.
Now supports both single file and folder (batch) processing with student ID validation.
"""

import sys
import os
from pathlib import Path
from typing import Optional, List
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
    QPushButton, QLabel, QFileDialog, QMessageBox, QProgressBar,
    QTextEdit, QTabWidget, QScrollArea
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal


class BatchFileProcessor(QThread):
    """Background thread for processing multiple files in a folder."""
    
    finished = pyqtSignal(bool, str, list)  # success, message, processed_files
    progress = pyqtSignal(int)  # progress percentage
    file_progress = pyqtSignal(str)  # current file being processed
    
    def __init__(self, folder_path: str):
        super().__init__()
        self.folder_path = folder_path
        
    def run(self):
        """Process all Excel files in the selected folder."""
        try:
            # Find all Excel files in folder
            excel_files = []
            folder = Path(self.folder_path)
            
            for ext in ['*.xlsx', '*.xls']:
                excel_files.extend(folder.glob(ext))
            
            if not excel_files:
                self.finished.emit(False, "No Excel files found in the selected folder", [])
                return
            
            valid_files = []
            
            # Validate each file
            for i, file_path in enumerate(excel_files):
                self.file_progress.emit(f"Validating: {file_path.name}")
                
                try:
                    # Quick validation - try to read the file
                    df = pd.read_excel(file_path, nrows=5)  # Just read first 5 rows for validation
                    if not df.empty:
                        valid_files.append(str(file_path))
                    
                except Exception as e:
                    print(f"⚠️ Skipping {file_path.name}: {e}")
                
                # Update progress
                progress = int(((i + 1) / len(excel_files)) * 100)
                self.progress.emit(progress)
            
            if not valid_files:
                self.finished.emit(False, "No valid Excel files found in the folder", [])
                return
            
            message = f"Found {len(valid_files)} valid Excel files ready for processing:\n"
            for file_path in valid_files:
                message += f"• {Path(file_path).name}\n"
            
            self.finished.emit(True, message, valid_files)
            
        except Exception as e:
            self.finished.emit(False, f"Error processing folder: {str(e)}", [])


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
                self.finished.emit(False, "CSV files are not supported for result appending. Please use Excel format (.xlsx or .xls)")
                return
            elif file_ext in ['.xlsx', '.xls']:
                df = pd.read_excel(self.file_path)
            else:
                self.finished.emit(False, "Unsupported file format. Please use Excel (.xlsx, .xls) format.")
                return
            
            self.progress.emit(75)
            
            # Validate file content
            if df.empty:
                self.finished.emit(False, "File is empty")
                return
            
            self.progress.emit(100)
            
            rows, cols = df.shape
            message = f"Excel file loaded successfully!\nRows: {rows}, Columns: {cols}\nResults will be appended to this file."
            self.finished.emit(True, message)
            
        except Exception as e:
            self.finished.emit(False, f"Error processing file: {str(e)}")


class BatchDataProcessor(QThread):
    """Background thread for processing multiple files."""
    
    finished = pyqtSignal(bool, str, dict)  # success, message, results_summary
    progress = pyqtSignal(int)  # overall progress
    file_progress = pyqtSignal(str)  # current file being processed
    file_completed = pyqtSignal(str, bool, str)  # file_name, success, message
    
    def __init__(self, file_paths: List[str], script_path: str = "data.py"):
        super().__init__()
        self.file_paths = file_paths
        self.script_path = script_path
        
    def run(self):
        """Process all files in the list."""
        total_files = len(self.file_paths)
        successful_files = []
        failed_files = []
        results_summary = {}
        
        try:
            # Check if data.py exists
            if not os.path.exists(self.script_path):
                self.finished.emit(False, f"Processing script '{self.script_path}' not found", {})
                return
            
            for i, file_path in enumerate(self.file_paths):
                file_name = Path(file_path).name
                self.file_progress.emit(f"Processing: {file_name}")
                
                try:
                    # Process individual file
                    result = subprocess.run(
                        [sys.executable, self.script_path, file_path],
                        capture_output=True,
                        text=True,
                        check=True,
                        encoding='utf-8',
                        errors='replace'  # Handle encoding errors gracefully
                    )
                    
                    output = result.stdout if result.stdout else "Processing completed successfully"
                    
                    # Process the calculation results for this file
                    try:
                        file_results = self._process_single_file_results(output, file_path)
                        results_summary[file_name] = file_results
                        successful_files.append(file_name)
                        self.file_completed.emit(file_name, True, f"Processed successfully")
                        
                    except Exception as calc_error:
                        failed_files.append(file_name)
                        # Check if it's a student ID validation error
                        if "Invalid student ID format" in str(calc_error):
                            self.file_completed.emit(file_name, False, f"Invalid student ID format")
                        else:
                            self.file_completed.emit(file_name, False, f"Calculation failed: {str(calc_error)}")
                
                except subprocess.CalledProcessError as e:
                    error_msg = e.stderr if e.stderr else str(e)
                    failed_files.append(file_name)
                    
                    # Check if it's a student ID validation error
                    if "Invalid student ID format" in error_msg:
                        self.file_completed.emit(file_name, False, f"Invalid student ID format - check file")
                    else:
                        self.file_completed.emit(file_name, False, f"Processing failed: {error_msg}")
                
                except Exception as e:
                    failed_files.append(file_name)
                    self.file_completed.emit(file_name, False, f"Unexpected error: {str(e)}")
                
                # Update overall progress
                progress = int(((i + 1) / total_files) * 100)
                self.progress.emit(progress)
            
            # Compile final summary
            summary_message = f"""
Batch Processing Complete!

Successfully processed: {len(successful_files)} files
Failed: {len(failed_files)} files

Successful files:
""" + "\n".join([f"- {f}" for f in successful_files])
            
            if failed_files:
                summary_message += f"\n\nFailed files:\n" + "\n".join([f"- {f}" for f in failed_files])
            
            self.finished.emit(True, summary_message, results_summary)
            
        except Exception as e:
            self.finished.emit(False, f"Batch processing failed: {str(e)}", {})
    
    def _process_single_file_results(self, message: str, file_path: str):
        """Process CLO/PLO calculations for a single file and append to Excel."""
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

        # Print to terminal for this file
        file_name = Path(file_path).name
        print(f"\n{'='*50}")
        print(f"Results for: {file_name}")
        print(f"{'='*50}")
        
        print("\nCLO Scores:")
        for student, scores in clo_scores.items():
            print(f"{student}: {scores}")

        print("\nPLO Scores:")
        for student, scores in plo_scores.items():
            print(f"{student}: {scores}")

        print("\nFinal Grades:")
        for student, percent in grades.items():
            letter = get_letter_grade(percent)
            print(f"{student}: {percent:.2f}% ({letter})")

        print("\nTotal CLO Weights:")
        for clo, weight in clo_weights.items():
            print(f"{clo}: {weight} %")

        # Append results to the original Excel file
        try:
            updated_file_path = export_clo_plo_results(clo_scores, plo_scores, grades, data_dict, file_path)
            print(f"Results appended to: {updated_file_path}")
            
            return {
                "students_count": len(clo_scores),
                "clo_count": len(set().union(*[scores.keys() for scores in clo_scores.values()])),
                "plo_count": len(set().union(*[scores.keys() for scores in plo_scores.values()])),
                "excel_updated": True
            }
            
        except Exception as excel_error:
            print(f"\nExcel append failed for {file_name}: {excel_error}")
            return {
                "students_count": len(clo_scores),
                "clo_count": len(set().union(*[scores.keys() for scores in clo_scores.values()])),
                "plo_count": len(set().union(*[scores.keys() for scores in plo_scores.values()])),
                "excel_updated": False,
                "excel_error": str(excel_error)
            }


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
                encoding='utf-8',
                errors='replace'  # Handle encoding errors gracefully
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
    """Main application window with enhanced batch processing capabilities and student ID validation."""
    
    def __init__(self):
        super().__init__()
        self.current_file_path: Optional[str] = None
        self.current_file_paths: List[str] = []
        self.processing_mode: str = "single"  # "single" or "batch"
        
        # Thread references
        self.file_processor: Optional[FileProcessor] = None
        self.batch_file_processor: Optional[BatchFileProcessor] = None
        self.data_processor: Optional[DataProcessor] = None
        self.batch_data_processor: Optional[BatchDataProcessor] = None
        
        self.init_ui()
    
    def init_ui(self):
        """Initialize the enhanced user interface."""
        self.setWindowTitle("Habib University - CLO/PLO Mapping (Enhanced with ID Validation)")
        self.setMinimumSize(700, 500)
        self.resize(800, 600)
        
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
        
        # Student ID format info
        id_info = QLabel("📧 Student IDs must be in format: hm08298@st.habib.edu.pk or hm08298")
        id_info.setStyleSheet("color: #666; font-size: 12px; padding: 8px; background: #f8f9fa; border-radius: 4px;")
        id_info.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(id_info)
        
        # Enhanced file selection section
        file_layout = QVBoxLayout()
        
        # Selection buttons row
        button_layout = QHBoxLayout()
        
        self.browse_file_btn = QPushButton("📄 Select Single File")
        self.browse_file_btn.clicked.connect(self.browse_single_file)
        self.browse_file_btn.setStyleSheet(self._get_button_style())
        
        self.browse_folder_btn = QPushButton("📁 Select Folder (Batch)")
        self.browse_folder_btn.clicked.connect(self.browse_folder)
        self.browse_folder_btn.setStyleSheet(self._get_button_style())
        
        button_layout.addWidget(self.browse_file_btn)
        button_layout.addWidget(self.browse_folder_btn)
        file_layout.addLayout(button_layout)
        
        # File/folder display
        self.file_label = QLabel("No files selected")
        self.file_label.setStyleSheet("padding: 12px; border: 1px solid #ccc; background: #f9f9f9; min-height: 60px;")
        self.file_label.setWordWrap(True)
        file_layout.addWidget(self.file_label)
        
        layout.addLayout(file_layout)
        
        # Supported formats info
        info = QLabel("Supported: Excel (.xlsx, .xls) - Results will be appended to original files")
        info.setStyleSheet("color: #666; font-size: 12px;")
        layout.addWidget(info)
        
        # Process files button
        self.process_btn = QPushButton("🚀 Process Files")
        self.process_btn.clicked.connect(self.process_files)
        self.process_btn.setEnabled(False)
        self.process_btn.setStyleSheet(self._get_button_style())
        layout.addWidget(self.process_btn)
        
        # Status section with tabs
        self.status_tabs = QTabWidget()
        
        # Status tab
        status_widget = QWidget()
        status_layout = QVBoxLayout(status_widget)
        
        self.status_label = QLabel("Ready - Please select Excel file(s) to begin")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status_label.setStyleSheet("padding: 12px; border: 1px solid #ddd; background: #f5f5f5;")
        status_layout.addWidget(self.status_label)
        
        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        status_layout.addWidget(self.progress_bar)
        
        self.status_tabs.addTab(status_widget, "📊 Status")
        
        # Batch progress tab
        batch_widget = QWidget()
        batch_layout = QVBoxLayout(batch_widget)
        
        self.batch_progress_label = QLabel("Batch processing not started")
        self.batch_progress_label.setStyleSheet("padding: 8px; border: 1px solid #ddd; background: #f9f9f9;")
        batch_layout.addWidget(self.batch_progress_label)
        
        # Scroll area for file progress
        scroll_area = QScrollArea()
        scroll_widget = QWidget()
        self.batch_results_layout = QVBoxLayout(scroll_widget)
        scroll_area.setWidget(scroll_widget)
        scroll_area.setWidgetResizable(True)
        scroll_area.setMaximumHeight(200)
        batch_layout.addWidget(scroll_area)
        
        self.status_tabs.addTab(batch_widget, "📁 Batch Progress")
        
        layout.addWidget(self.status_tabs)
        
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
    
    def browse_single_file(self):
        """Open file dialog to select a single file."""
        file_path, _ = QFileDialog.getOpenFileName(
            self, 
            "Select Excel File - Habib University",
            "",
            "Excel files (*.xlsx *.xls);;All files (*.*)"
        )
        
        if file_path:
            self.processing_mode = "single"
            self.current_file_paths = [file_path]
            self.load_file(file_path)
    
    def browse_folder(self):
        """Open folder dialog to select a folder for batch processing."""
        folder_path = QFileDialog.getExistingDirectory(
            self,
            "Select Folder with Excel Files - Habib University",
            ""
        )
        
        if folder_path:
            self.processing_mode = "batch"
            self.load_folder(folder_path)
    
    def load_file(self, file_path: str):
        """Load and process a single selected file."""
        self.current_file_path = file_path
        file_name = Path(file_path).name
        
        # Update UI
        self.file_label.setText(f"📄 Single File Mode\nSelected: {file_name}")
        self._update_status("Processing file...", "processing")
        
        # Show progress bar
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        
        # Disable buttons
        self._set_buttons_enabled(False)
        
        # Start processing
        self.file_processor = FileProcessor(file_path)
        self.file_processor.finished.connect(self.on_file_processed)
        self.file_processor.progress.connect(self.progress_bar.setValue)
        self.file_processor.start()
    
    def load_folder(self, folder_path: str):
        """Load and validate all Excel files in the selected folder."""
        folder_name = Path(folder_path).name
        
        # Update UI
        self.file_label.setText(f"📁 Batch Mode\nScanning folder: {folder_name}")
        self._update_status("Scanning folder for Excel files...", "processing")
        
        # Show progress bar
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        
        # Switch to batch progress tab
        self.status_tabs.setCurrentIndex(1)
        self.batch_progress_label.setText("Scanning folder for Excel files...")
        
        # Clear previous batch results
        self._clear_batch_results()
        
        # Disable buttons
        self._set_buttons_enabled(False)
        
        # Start batch file processing
        self.batch_file_processor = BatchFileProcessor(folder_path)
        self.batch_file_processor.finished.connect(self.on_batch_files_processed)
        self.batch_file_processor.progress.connect(self.progress_bar.setValue)
        self.batch_file_processor.file_progress.connect(self.batch_progress_label.setText)
        self.batch_file_processor.start()
    
    def on_file_processed(self, success: bool, message: str):
        """Handle single file processing completion."""
        # Hide progress bar
        self.progress_bar.setVisible(False)
        
        # Re-enable browse buttons
        self._set_buttons_enabled(True, process_enabled=success)
        
        # Update status
        if success:
            self._update_status(message, "success")
        else:
            self._update_status(f"Error: {message}", "error")
        
        # Clean up
        if self.file_processor:
            self.file_processor.deleteLater()
            self.file_processor = None
    
    def on_batch_files_processed(self, success: bool, message: str, file_paths: List[str]):
        """Handle batch file validation completion."""
        # Hide progress bar
        self.progress_bar.setVisible(False)
        
        # Re-enable browse buttons
        self._set_buttons_enabled(True, process_enabled=success and len(file_paths) > 0)
        
        if success and file_paths:
            self.current_file_paths = file_paths
            
            # Update file display
            file_list = "\n".join([f"• {Path(f).name}" for f in file_paths[:10]])
            if len(file_paths) > 10:
                file_list += f"\n... and {len(file_paths) - 10} more files"
            
            self.file_label.setText(f"📁 Batch Mode\nFound {len(file_paths)} valid Excel files:\n{file_list}")
            self._update_status(f"Ready to process {len(file_paths)} Excel files", "success")
            self.batch_progress_label.setText(f"Ready to process {len(file_paths)} files")
            
        else:
            self._update_status(f"Error: {message}", "error")
            self.batch_progress_label.setText(f"Error: {message}")
        
        # Clean up
        if self.batch_file_processor:
            self.batch_file_processor.deleteLater()
            self.batch_file_processor = None
    
    def process_files(self):
        """Process the loaded file(s)."""
        if self.processing_mode == "single":
            self._process_single_file()
        elif self.processing_mode == "batch":
            self._process_batch_files()
    
    def _process_single_file(self):
        """Process a single file."""
        if not self.current_file_path:
            self._update_status("No file selected for processing", "error")
            return
        
        # Update UI
        self._update_status("Running data.py processing...", "processing")
        self._set_buttons_enabled(False)
        
        # Start data processing
        self.data_processor = DataProcessor(self.current_file_path)
        self.data_processor.finished.connect(self.on_data_processed)
        self.data_processor.progress.connect(self._update_status)
        self.data_processor.start()
    
    def _process_batch_files(self):
        """Process multiple files in batch."""
        if not self.current_file_paths:
            self._update_status("No files selected for processing", "error")
            return
        
        # Update UI
        self._update_status(f"Starting batch processing of {len(self.current_file_paths)} files...", "processing")
        self._set_buttons_enabled(False)
        
        # Show progress bar and switch to batch tab
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.status_tabs.setCurrentIndex(1)
        
        # Clear previous results
        self._clear_batch_results()
        
        # Start batch processing
        self.batch_data_processor = BatchDataProcessor(self.current_file_paths)
        self.batch_data_processor.finished.connect(self.on_batch_data_processed)
        self.batch_data_processor.progress.connect(self.progress_bar.setValue)
        self.batch_data_processor.file_progress.connect(self.batch_progress_label.setText)
        self.batch_data_processor.file_completed.connect(self._add_batch_result)
        self.batch_data_processor.start()
    
    def on_data_processed(self, success: bool, message: str):
        """Handle single file data processing completion with enhanced error handling."""
        # Re-enable buttons
        self._set_buttons_enabled(True)
        
        if not success:
            # Check if it's a student ID validation error
            if "Invalid student ID format" in message:
                self._show_student_id_error_dialog(message)
                self._update_status("Student ID validation failed - check format", "error")
            else:
                self._update_status(f"Processing failed: {message}", "error")
            self._cleanup_single_processor()
            return
        
        # Print full output to console
        print("\n=== Data Processing Output ===")
        print(message)
        print("==============================\n")
        
        # Process the results
        try:
            self._process_calculation_results(message)
        except Exception as e:
            if "Invalid student ID format" in str(e):
                self._show_student_id_error_dialog(str(e))
                self._update_status("Student ID validation failed", "error")
            else:
                self._update_status(f"Calculation failed: {str(e)}", "error")
        
        self._cleanup_single_processor()
    
    def on_batch_data_processed(self, success: bool, message: str, results_summary: dict):
        """Handle batch data processing completion."""
        # Hide progress bar
        self.progress_bar.setVisible(False)
        
        # Re-enable buttons
        self._set_buttons_enabled(True)
        
        if success:
            self._update_status("Batch processing completed! Check terminal and Batch Progress tab for details.", "success")
            self.batch_progress_label.setText("✅ Batch processing completed!")
            
            # Show summary dialog
            self._show_batch_success_dialog(message, results_summary)
        else:
            self._update_status(f"Batch processing failed: {message}", "error")
            self.batch_progress_label.setText(f"❌ Batch processing failed: {message}")
        
        print("\n" + "="*60)
        print("BATCH PROCESSING SUMMARY")
        print("="*60)
        print(message)
        
        self._cleanup_batch_processor()
    
    def _process_calculation_results(self, message: str):
        """Process CLO/PLO calculations and append to original Excel file."""
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

        # Append results to the original Excel file
        try:
            updated_file_path = export_clo_plo_results(clo_scores, plo_scores, grades, data_dict, self.current_file_path)
            self._update_status(f"CLO/PLO calculation complete. Results appended to: {Path(updated_file_path).name}", "success")
            self._show_success_dialog(updated_file_path)
        except Exception as excel_error:
            print(f"\n❌ Excel append failed: {excel_error}")
            self._update_status("CLO/PLO calculation complete. See terminal output. (Excel append failed)", "warning")

    def _print_scores_to_console(self, clo_scores, plo_scores, grades, clo_weights):
        """Print CLO, PLO, Grades, and CLO weights to terminal."""
        print("\nCLO Scores:")
        for student, scores in clo_scores.items():
            print(f"{student}: {scores}")

        print("\nPLO Scores:")
        for student, scores in plo_scores.items():
            print(f"{student}: {scores}")

        print("\nFinal Grades:")
        for student, percent in grades.items():
            letter = get_letter_grade(percent)
            print(f"{student}: {percent:.2f}% ({letter})")

        print("\nTotal CLO Weights:")
        for clo, weight in clo_weights.items():
            print(f"{clo}: {weight} %")
    
    def _add_batch_result(self, file_name: str, success: bool, message: str):
        """Add a file result to the batch progress display."""
        result_label = QLabel(f"{file_name}: {message}")
        if success:
            result_label.setStyleSheet("color: #28a745; padding: 4px;")
        else:
            # Highlight student ID errors differently
            if "Invalid student ID format" in message:
                result_label.setStyleSheet("color: #e74c3c; padding: 4px; font-weight: bold;")
            else:
                result_label.setStyleSheet("color: #dc3545; padding: 4px;")
        self.batch_results_layout.addWidget(result_label)
    
    def _clear_batch_results(self):
        """Clear previous batch results from the display."""
        while self.batch_results_layout.count():
            child = self.batch_results_layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()
    
    def _set_buttons_enabled(self, enabled: bool, process_enabled: Optional[bool] = None):
        """Enable/disable buttons with optional separate control for process button."""
        self.browse_file_btn.setEnabled(enabled)
        self.browse_folder_btn.setEnabled(enabled)
        
        if process_enabled is not None:
            self.process_btn.setEnabled(process_enabled)
        else:
            # Check if we have valid files to process
            has_files = bool(self.current_file_path) or bool(self.current_file_paths)
            self.process_btn.setEnabled(enabled and has_files)
    
    def _show_success_dialog(self, output_file: str):
        """Show success dialog with file path for single file processing."""
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Icon.Information)
        msg.setWindowTitle("Processing Complete")
        msg.setText(f"CLO/PLO results have been successfully appended to your file:\n\n{output_file}\n\nNew sheet added:\n• CLO PLO Results\n\n📧 All student IDs formatted to Habib University email format")
        msg.setStandardButtons(QMessageBox.StandardButton.Ok)
        msg.exec()
    
    def _show_batch_success_dialog(self, summary_message: str, results_summary: dict):
        """Show success dialog for batch processing."""
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Icon.Information)
        msg.setWindowTitle("Batch Processing Complete")
        
        successful_count = len([f for f, data in results_summary.items() if data.get('excel_updated', False)])
        total_count = len(results_summary)
        
        dialog_text = f"Batch processing completed!\n\n"
        dialog_text += f"✅ Successfully processed: {successful_count}/{total_count} files\n\n"
        dialog_text += "Each successfully processed file now contains:\n"
        dialog_text += "• CLO PLO Results sheet with color-coded performance data\n"
        dialog_text += "• Student IDs formatted to Habib University email format\n\n"
        dialog_text += "Check the Batch Progress tab and terminal output for detailed results."
        
        msg.setText(dialog_text)
        msg.setDetailedText(summary_message)
        msg.setStandardButtons(QMessageBox.StandardButton.Ok)
        msg.exec()
    
    def _show_student_id_error_dialog(self, error_message: str):
        """Show detailed error dialog for student ID validation issues."""
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Icon.Critical)
        msg.setWindowTitle("Student ID Format Error")
        
        # Create a user-friendly message
        dialog_text = """❌ Student ID Format Error Detected

The Excel file contains student IDs that don't match the required Habib University format.

📧 Required Format:
• hm08298@st.habib.edu.pk (full email format)
• hm08298 (short format - will be auto-converted)

✅ Valid Examples:
• ab12345 (any 2 letters + 5 digits)
• xy67890@st.habib.edu.pk
• hm08298

❌ Invalid Examples:
• 12345 (missing initials)
• abc123 (wrong number of digits)
• hm123456 (too many digits)

Please fix the student IDs in your Excel file and try again."""
        
        msg.setText(dialog_text)
        msg.setDetailedText(error_message)
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
    
    def _cleanup_single_processor(self):
        """Clean up single file data processor thread."""
        if self.data_processor:
            self.data_processor.deleteLater()
            self.data_processor = None
    
    def _cleanup_batch_processor(self):
        """Clean up batch data processor thread."""
        if self.batch_data_processor:
            self.batch_data_processor.deleteLater()
            self.batch_data_processor = None
    
    def closeEvent(self, event):
        """Handle application close."""
        # Terminate all running threads
        threads_to_terminate = [
            self.file_processor,
            self.batch_file_processor,
            self.data_processor,
            self.batch_data_processor
        ]
        
        for thread in threads_to_terminate:
            if thread and thread.isRunning():
                thread.terminate()
                thread.wait()
        
        event.accept()


def main():
    """Main application entry point."""
    app = QApplication(sys.argv)
    app.setApplicationName("Habib University CLO/PLO Mapping Tool - Enhanced with ID Validation")
    app.setStyle("Fusion")
    
    window = HabibUniversityApp()
    window.show()
    
    sys.exit(app.exec())


if __name__ == "__main__":
    main()