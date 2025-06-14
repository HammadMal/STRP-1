#!/usr/bin/env python3
"""
Habib University File Manager Application
A PyQt6-based desktop application for file management and processing.
"""

import sys
import os
from pathlib import Path
from typing import Optional
import pandas as pd

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QTextEdit, QFrame, QMessageBox,
    QProgressBar, QGroupBox, QGridLayout, QSizePolicy
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QTimer, QUrl
from PyQt6.QtGui import QFont, QPixmap, QPalette, QColor, QIcon
from PyQt6.QtNetwork import QNetworkAccessManager, QNetworkRequest


class FileProcessor(QThread):
    """Background thread for processing files without blocking the UI."""
    
    finished = pyqtSignal(bool, str)  # success, message
    progress = pyqtSignal(int)  # progress percentage
    
    def __init__(self, file_path: str):
        super().__init__()
        self.file_path = file_path
        
    def run(self):
        """Process the uploaded file."""
        try:
            self.progress.emit(25)
            
            # Simulate file processing
            self.msleep(500)  # Small delay for demo
            
            file_ext = Path(self.file_path).suffix.lower()
            
            self.progress.emit(50)
            
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


class ModernButton(QPushButton):
    """Custom styled button with modern appearance."""
    
    def __init__(self, text: str, primary: bool = False):
        super().__init__(text)
        self.setMinimumHeight(40)
        self.setFont(QFont("Segoe UI", 10, QFont.Weight.Medium))
        
        if primary:
            self.setStyleSheet("""
                QPushButton {
                    background-color: #6B2C91;
                    color: white;
                    border: none;
                    border-radius: 8px;
                    padding: 8px 16px;
                    font-weight: bold;
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
        else:
            self.setStyleSheet("""
                QPushButton {
                    background-color: white;
                    color: #333333;
                    border: 2px solid #DDDDDD;
                    border-radius: 8px;
                    padding: 8px 16px;
                }
                QPushButton:hover {
                    border-color: #6B2C91;
                    background-color: #F8F9FA;
                }
                QPushButton:pressed {
                    background-color: #E9ECEF;
                }
            """)


class StatusLabel(QLabel):
    """Custom styled status label."""
    
    def __init__(self, text: str = ""):
        super().__init__(text)
        self.setFont(QFont("Segoe UI", 9))
        self.setWordWrap(True)
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setMinimumHeight(60)
        self.set_neutral()
    
    def set_success(self, message: str):
        self.setText(message)
        self.setStyleSheet("""
            QLabel {
                background-color: #D4EDDA;
                color: #155724;
                border: 1px solid #C3E6CB;
                border-radius: 6px;
                padding: 12px;
            }
        """)
    
    def set_error(self, message: str):
        self.setText(message)
        self.setStyleSheet("""
            QLabel {
                background-color: #F8D7DA;
                color: #721C24;
                border: 1px solid #F5C6CB;
                border-radius: 6px;
                padding: 12px;
            }
        """)
    
    def set_neutral(self):
        self.setText("No file selected")
        self.setStyleSheet("""
            QLabel {
                background-color: #F8F9FA;
                color: #6C757D;
                border: 1px solid #DEE2E6;
                border-radius: 6px;
                padding: 12px;
            }
        """)


class HabibUniversityApp(QMainWindow):
    """Main application window for Habib University File Manager."""
    
    def __init__(self):
        super().__init__()
        self.current_file_path: Optional[str] = None
        self.file_processor: Optional[FileProcessor] = None
        self.logo_pixmap: Optional[QPixmap] = None
        
        self.load_logo()
        self.init_ui()
        self.setup_connections()
    
    def init_ui(self):
        """Initialize the user interface."""
        self.setWindowTitle("Habib University - File Manager")
        self.setMinimumSize(600, 500)
        self.resize(800, 600)
        
        # Set application icon (you can add an icon file later)
        # self.setWindowIcon(QIcon("habib_logo.png"))
        
        # Create central widget and main layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(20)
        main_layout.setContentsMargins(30, 30, 30, 30)
        
        # Header section
        self.create_header(main_layout)
        
        # File upload section
        self.create_file_section(main_layout)
        
        # Status section
        self.create_status_section(main_layout)
        
        # Progress section
        self.create_progress_section(main_layout)
        
        # Add stretch to push everything to top
        main_layout.addStretch()
        
        # Apply modern styling
        self.apply_styling()
    
    def load_logo(self):
        """Load the Habib University logo."""
        try:
            # Try to load logo from assets folder
            logo_path = os.path.join("assets", "habib_logo.png")
            if os.path.exists(logo_path):
                self.logo_pixmap = QPixmap(logo_path)
            else:
                # Try alternative common logo file names
                possible_names = [
                    "habib_university_logo.png",
                    "habib_logo.jpg", 
                    "habib_logo.jpeg",
                    "logo.png",
                    "logo.jpg"
                ]
                
                for name in possible_names:
                    alt_path = os.path.join("assets", name)
                    if os.path.exists(alt_path):
                        self.logo_pixmap = QPixmap(alt_path)
                        break
                        
            # If logo loaded successfully, verify it's not null
            if self.logo_pixmap and self.logo_pixmap.isNull():
                print("Warning: Logo file found but could not be loaded properly")
                self.logo_pixmap = None
                
        except Exception as e:
            print(f"Error loading logo: {e}")
            self.logo_pixmap = None
    
    def create_header(self, parent_layout):
        """Create the header section with university branding."""
        header_frame = QFrame()
        header_frame.setStyleSheet("""
            QFrame {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1, 
                    stop:0 #6B2C91, stop:1 #8E44AD);
                border-radius: 12px;
                padding: 20px;
            }
        """)
        
        header_layout = QVBoxLayout(header_frame)
        
        # Logo and title container
        title_container = QHBoxLayout()
        
        # Logo (placeholder for now - you can add the actual logo file later)
        logo_label = QLabel()
        if self.logo_pixmap and not self.logo_pixmap.isNull():
            scaled_logo = self.logo_pixmap.scaled(80, 80, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
            logo_label.setPixmap(scaled_logo)
        else:
            # Placeholder logo text
            logo_label.setText("🦁")
            logo_label.setFont(QFont("Segoe UI", 48))
            logo_label.setStyleSheet("color: #F4E4BC;")
        
        logo_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        logo_label.setMaximumSize(100, 100)
        
        # Title section
        title_section = QVBoxLayout()
        
        # University name
        title_label = QLabel("Habib University")
        title_label.setFont(QFont("Segoe UI", 24, QFont.Weight.Bold))
        title_label.setStyleSheet("color: white; margin: 0px;")
        
        # Subtitle
        subtitle_label = QLabel("File Management System")
        subtitle_label.setFont(QFont("Segoe UI", 12))
        subtitle_label.setStyleSheet("color: #E8D5F2; margin: 0px;")
        
        title_section.addWidget(title_label)
        title_section.addWidget(subtitle_label)
        title_section.addStretch()
        
        # Add to container
        title_container.addWidget(logo_label)
        title_container.addSpacing(20)
        title_container.addLayout(title_section, 1)
        
        header_layout.addLayout(title_container)
        
        parent_layout.addWidget(header_frame)
    
    def create_file_section(self, parent_layout):
        """Create the file upload section."""
        file_group = QGroupBox("File Upload")
        file_group.setFont(QFont("Segoe UI", 11, QFont.Weight.Medium))
        file_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 2px solid #D1C4E9;
                border-radius: 10px;
                margin-top: 10px;
                padding-top: 15px;
                background-color: #FAFAFA;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 15px;
                padding: 0 8px 0 8px;
                background-color: white;
                color: #6B2C91;
            }
        """)
        
        file_layout = QVBoxLayout(file_group)
        file_layout.setSpacing(15)
        
        # File selection area
        file_selection_layout = QHBoxLayout()
        
        self.file_path_label = QLabel("No file selected")
        self.file_path_label.setFont(QFont("Segoe UI", 10))
        self.file_path_label.setStyleSheet("""
            QLabel {
                background-color: #F8F9FA;
                border: 1px solid #DEE2E6;
                border-radius: 6px;
                padding: 10px;
                color: #6C757D;
            }
        """)
        
        self.browse_button = ModernButton("Browse Files", primary=True)
        self.browse_button.setMaximumWidth(150)
        
        file_selection_layout.addWidget(self.file_path_label, 1)
        file_selection_layout.addWidget(self.browse_button)
        
        # Supported formats info
        info_label = QLabel("Supported formats: Excel (.xlsx, .xls) and CSV (.csv)")
        info_label.setFont(QFont("Segoe UI", 9))
        info_label.setStyleSheet("color: #6C757D; margin-top: 5px;")
        
        file_layout.addLayout(file_selection_layout)
        file_layout.addWidget(info_label)
        
        parent_layout.addWidget(file_group)
    
    def create_status_section(self, parent_layout):
        """Create the status display section."""
        status_group = QGroupBox("Status")
        status_group.setFont(QFont("Segoe UI", 11, QFont.Weight.Medium))
        status_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 2px solid #D1C4E9;
                border-radius: 10px;
                margin-top: 10px;
                padding-top: 15px;
                background-color: #FAFAFA;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 15px;
                padding: 0 8px 0 8px;
                background-color: white;
                color: #6B2C91;
            }
        """)
        
        status_layout = QVBoxLayout(status_group)
        
        self.status_label = StatusLabel()
        status_layout.addWidget(self.status_label)
        
        parent_layout.addWidget(status_group)
    
    def create_progress_section(self, parent_layout):
        """Create the progress bar section."""
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid #DDDDDD;
                border-radius: 6px;
                text-align: center;
                font-weight: bold;
                height: 25px;
            }
            QProgressBar::chunk {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, 
                    stop:0 #6B2C91, stop:1 #8E44AD);
                border-radius: 5px;
            }
        """)
        
        parent_layout.addWidget(self.progress_bar)
    
    def apply_styling(self):
        """Apply overall application styling."""
        self.setStyleSheet("""
            QMainWindow {
                background-color: #FFFFFF;
            }
            QWidget {
                font-family: 'Segoe UI', Arial, sans-serif;
            }
        """)
    
    def setup_connections(self):
        """Setup signal-slot connections."""
        self.browse_button.clicked.connect(self.browse_file)
    
    def browse_file(self):
        """Open file dialog to select a file."""
        file_dialog = QFileDialog(self)
        file_dialog.setNameFilter("Spreadsheet files (*.xlsx *.xls *.csv);;All files (*.*)")
        file_dialog.setFileMode(QFileDialog.FileMode.ExistingFile)
        file_dialog.setWindowTitle("Select File - Habib University")
        
        if file_dialog.exec():
            selected_files = file_dialog.selectedFiles()
            if selected_files:
                self.load_file(selected_files[0])
    
    def load_file(self, file_path: str):
        """Load and process the selected file."""
        self.current_file_path = file_path
        file_name = Path(file_path).name
        
        # Update UI
        self.file_path_label.setText(f"Selected: {file_name}")
        self.file_path_label.setStyleSheet("""
            QLabel {
                background-color: #F3E5F5;
                border: 1px solid #CE93D8;
                border-radius: 6px;
                padding: 10px;
                color: #6B2C91;
            }
        """)
        
        self.status_label.setText("Processing file...")
        self.status_label.setStyleSheet("""
            QLabel {
                background-color: #FFF3CD;
                color: #856404;
                border: 1px solid #FFEAA7;
                border-radius: 6px;
                padding: 12px;
            }
        """)
        
        # Show progress bar
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        
        # Disable browse button during processing
        self.browse_button.setEnabled(False)
        
        # Start file processing in background thread
        self.file_processor = FileProcessor(file_path)
        self.file_processor.finished.connect(self.on_file_processed)
        self.file_processor.progress.connect(self.progress_bar.setValue)
        self.file_processor.start()
    
    def on_file_processed(self, success: bool, message: str):
        """Handle file processing completion."""
        # Hide progress bar
        self.progress_bar.setVisible(False)
        
        # Re-enable browse button
        self.browse_button.setEnabled(True)
        
        # Update status
        if success:
            self.status_label.set_success(message)
        else:
            self.status_label.set_error(message)
        
        # Clean up thread
        if self.file_processor:
            self.file_processor.deleteLater()
            self.file_processor = None
    
    def closeEvent(self, event):
        """Handle application close event."""
        if self.file_processor and self.file_processor.isRunning():
            self.file_processor.terminate()
            self.file_processor.wait()
        event.accept()


def main():
    """Main application entry point."""
    app = QApplication(sys.argv)
    
    # Set application properties
    app.setApplicationName("Habib University File Manager")
    app.setApplicationVersion("1.0.0")
    app.setOrganizationName("Habib University")
    
    # Create and show main window
    window = HabibUniversityApp()
    window.show()
    
    # Start event loop
    sys.exit(app.exec())


if __name__ == "__main__":
    main()