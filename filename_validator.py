"""
Filename Validation Module for Habib University CLO/PLO Mapping Tool
Validates Excel filenames against the required format: 2515-EE-437-L1.xlsx

Expected format breakdown:
- Component 1: 4 digits (semester/year code) - e.g., "2515"
- Component 2: Department code (2-4 letters) - e.g., "EE", "CS", "BIO", "MGMT"
- Component 3: 3 digits (course number) - e.g., "437", "101"
- Component 4: Section identifier - L or T followed by digit(s) - e.g., "L1", "T2", "L10"
- Extension: Must be .xlsx or .xls
"""

import re
import os
from pathlib import Path
from typing import Tuple, List, Dict


class FilenameValidationError(Exception):
    """Custom exception for filename validation errors."""
    pass


def validate_filename_format(filename: str) -> Tuple[bool, str, Dict[str, str]]:
    """
    Validate if filename matches the required format: XXXX-XX-XXX-XX.xlsx
    
    Args:
        filename (str): The filename to validate
        
    Returns:
        Tuple[bool, str, Dict[str, str]]: 
            - success: True if valid, False otherwise
            - message: Success/error message
            - components: Dictionary with parsed components if valid
            
    Example:
        >>> validate_filename_format("2515-EE-437-L1.xlsx")
        (True, "Valid filename format", {
            "semester": "2515", 
            "department": "EE", 
            "course": "437", 
            "section": "L1",
            "extension": ".xlsx"
        })
    """
    
    # Initialize components dictionary
    components = {
        "semester": "",
        "department": "",
        "course": "",
        "section": "",
        "extension": ""
    }
    
    try:
        # Get the base name without directory path
        base_filename = os.path.basename(filename)
        
        # Split filename and extension
        name_part, extension = os.path.splitext(base_filename)
        components["extension"] = extension.lower()
        
        # Check if extension is valid
        if extension.lower() not in ['.xlsx', '.xls']:
            return False, f"Invalid file extension '{extension}'. Must be .xlsx or .xls", components
        
        # Split the name part by dashes
        parts = name_part.split('-')
        
        # Check if we have exactly 4 parts (3 dashes create 4 components)
        if len(parts) != 4:
            return False, f"Invalid filename structure. Expected format: XXXX-XX-XXX-XX.xlsx (found {len(parts)} components, expected 4)", components
        
        semester_code, dept_code, course_code, section_code = parts
        
        # Validate Component 1: Semester/Year Code (4 digits)
        if not re.match(r'^\d{4}$', semester_code):
            return False, f"Invalid semester code '{semester_code}'. Must be exactly 4 digits (e.g., '2515', '2412')", components
        components["semester"] = semester_code
        
        # Validate Component 2: Department Code (2-4 letters only)
        if not re.match(r'^[A-Za-z]{2,4}$', dept_code):
            return False, f"Invalid department code '{dept_code}'. Must be 2-4 letters only (e.g., 'EE', 'CS', 'BIO', 'MGMT')", components
        components["department"] = dept_code.upper()
        
        # Validate Component 3: Course Number (exactly 3 digits)
        if not re.match(r'^\d{3}$', course_code):
            return False, f"Invalid course code '{course_code}'. Must be exactly 3 digits (e.g., '437', '101', '205')", components
        components["course"] = course_code
        
        # Validate Component 4: Section Code (L or T followed by digit(s))
        if not re.match(r'^[LT]\d+$', section_code.upper()):
            return False, f"Invalid section code '{section_code}'. Must be L or T followed by digit(s) (e.g., 'L1', 'T2', 'L10')", components
        components["section"] = section_code.upper()
        
        # If we reach here, all validations passed
        success_message = f"✅ Valid filename format: {semester_code}-{components['department']}-{course_code}-{components['section']}{extension}"
        return True, success_message, components
        
    except Exception as e:
        return False, f"Error parsing filename: {str(e)}", components


def validate_batch_filenames(file_paths: List[str]) -> Tuple[List[str], List[str], Dict[str, str]]:
    """
    Validate multiple filenames for batch processing.
    
    Args:
        file_paths (List[str]): List of file paths to validate
        
    Returns:
        Tuple[List[str], List[str], Dict[str, str]]:
            - valid_files: List of valid file paths
            - invalid_files: List of invalid file paths
            - validation_results: Dictionary with filename -> error message for invalid files
    """
    
    valid_files = []
    invalid_files = []
    validation_results = {}
    
    for file_path in file_paths:
        filename = os.path.basename(file_path)
        is_valid, message, components = validate_filename_format(filename)
        
        if is_valid:
            valid_files.append(file_path)
            validation_results[filename] = message
        else:
            invalid_files.append(file_path)
            validation_results[filename] = f"❌ {message}"
    
    return valid_files, invalid_files, validation_results


def get_filename_format_help() -> str:
    """
    Get help text explaining the required filename format.
    
    Returns:
        str: Formatted help text
    """
    help_text = """
📋 Required Filename Format: XXXX-XX-XXX-XX.xlsx

🔍 Component Breakdown:
1️⃣ Semester Code: 4 digits (e.g., 2515, 2412, 2511)
2️⃣ Department Code: 2-4 letters (e.g., EE, CS, BIO, MGMT)
3️⃣ Course Number: 3 digits (e.g., 437, 101, 205)
4️⃣ Section Code: L or T + digit(s) (e.g., L1, T2, L10)

✅ Valid Examples:
• 2515-EE-437-L1.xlsx
• 2412-CS-101-T2.xlsx
• 2511-BIO-205-L3.xlsx
• 2515-MGMT-301-T1.xlsx

❌ Invalid Examples:
• 251-EE-437-L1.xlsx (semester code too short)
• 2515-E1-437-L1.xlsx (department has numbers)
• 2515-EE-43-L1.xlsx (course code too short)
• 2515-EE-437-A1.xlsx (section must start with L or T)
• 2515-EE-437-L1.csv (wrong file extension)

🚨 Common Issues to Fix:
• Use exactly 3 dashes to separate components
• Department must be 2-4 letters only (no numbers)
• Section code must start with L (Lab) or T (Tutorial/Theory)
• Course codes must be exactly 3 digits
• Use .xlsx or .xls extension only
"""
    return help_text


def format_validation_summary(valid_count: int, invalid_count: int, validation_results: Dict[str, str]) -> str:
    """
    Format a summary of validation results for display.
    
    Args:
        valid_count (int): Number of valid files
        invalid_count (int): Number of invalid files
        validation_results (Dict[str, str]): Validation results per file
        
    Returns:
        str: Formatted summary text
    """
    
    summary = f"📊 Filename Validation Summary:\n"
    summary += f"✅ Valid files: {valid_count}\n"
    summary += f"❌ Invalid files: {invalid_count}\n\n"
    
    if invalid_count > 0:
        summary += "🔍 Files with issues:\n"
        for filename, result in validation_results.items():
            if result.startswith("❌"):
                summary += f"• {filename}: {result[2:]}\n"  # Remove ❌ prefix
        
        summary += f"\n{get_filename_format_help()}"
    
    return summary


def extract_course_info(filename: str) -> Dict[str, str]:
    """
    Extract course information from a valid filename.
    
    Args:
        filename (str): Valid filename
        
    Returns:
        Dict[str, str]: Dictionary with course information
    """
    
    is_valid, message, components = validate_filename_format(filename)
    
    if not is_valid:
        raise FilenameValidationError(f"Cannot extract course info from invalid filename: {message}")
    
    return {
        "semester": components["semester"],
        "department": components["department"],
        "course_number": components["course"],
        "section": components["section"],
        "full_course_code": f"{components['department']}-{components['course']}",
        "semester_section": f"{components['semester']}-{components['section']}",
        "display_name": f"{components['department']} {components['course']} - Section {components['section']} (Semester {components['semester']})"
    }