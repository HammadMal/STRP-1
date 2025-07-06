"""
Excel Export Module for CLO/PLO Mapping Tool
Handles creation of formatted Excel reports with student performance data.
Modified to accept grades from terminal instead of calculating them internally.
"""

import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from datetime import datetime


def create_excel_output(clo_scores, plo_scores, grades, data_dict, output_file=None):
    """
    Create an Excel file with CLO, PLO scores, and final grades in a formatted table.
    
    Args:
        clo_scores (dict): Dictionary of CLO scores for each student
        plo_scores (dict): Dictionary of PLO scores for each student
        grades (dict): Dictionary of final grades for each student (from terminal calculation)
        data_dict (dict): Original data dictionary from data.py
        output_file (str, optional): Output file path. If None, auto-generates name.
    
    Returns:
        str: Path to the created Excel file
    """
    
    # Generate output filename if not provided
    if output_file is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"CLO_PLO_Results_{timestamp}.xlsx"
    
    # Prepare data for DataFrame
    data_rows = []
    
    # Get all unique CLOs and PLOs, ensuring we include all CLOs even if they're 0
    all_clos = set()
    all_plos = set()
    
    # First collect from actual scores
    for student_clos in clo_scores.values():
        all_clos.update(student_clos.keys())
    
    for student_plos in plo_scores.values():
        all_plos.update(student_plos.keys())
    
    # Ensure we always include CLO 1-5 (even if missing from scores)
    expected_clos = ['CLO 1', 'CLO 2', 'CLO 3', 'CLO 4', 'CLO 5']
    all_clos.update(expected_clos)
    
    # Sort CLOs and PLOs
    sorted_clos = sorted(all_clos, key=lambda x: int(x.split()[-1]) if x.split()[-1].isdigit() else 999)
    sorted_plos = sorted(all_plos, key=lambda x: int(x.split()[-1]) if x.split()[-1].isdigit() else 999)
    
    # Create data rows
    for student_id in clo_scores.keys():
        row_data = {'ID': student_id}
        
        # Add CLO scores (include all CLOs, showing 0 for missing ones)
        for clo in sorted_clos:
            row_data[clo] = clo_scores[student_id].get(clo, 0)
        
        # Add PLO scores
        for plo in sorted_plos:
            row_data[plo] = plo_scores[student_id].get(plo, 0)
        
        # Add overall grade (from terminal calculation)
        grade_percentage = grades.get(student_id, 0)
        letter_grade = _calculate_letter_grade(grade_percentage)
        row_data['Overall Grade'] = f"{grade_percentage:.2f}% ({letter_grade})"
        
        data_rows.append(row_data)
    
    # Create DataFrame
    df = pd.DataFrame(data_rows)
    
    # Write to Excel with formatting
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Write main data
        df.to_excel(writer, sheet_name='CLO PLO Results', index=False)
        
        # Get the workbook and worksheet
        workbook = writer.book
        worksheet = writer.sheets['CLO PLO Results']
        
        # Apply formatting
        _format_main_sheet(worksheet, df, sorted_clos, sorted_plos)
        
        # Create and format summary sheet
        _create_summary_sheet(writer, clo_scores, plo_scores, grades, sorted_clos, sorted_plos)
    
    return output_file


def _calculate_letter_grade(score):
    """Calculate letter grade based on numerical score using Habib University scale."""
    if score >= 95:
        return "A+"
    elif score >= 90:
        return "A"
    elif score >= 85:
        return "A-"
    elif score >= 80:
        return "B+"
    elif score >= 75:
        return "B"
    elif score >= 70:
        return "B-"
    elif score >= 67:
        return "C+"
    elif score >= 63:
        return "C"
    elif score >= 60:
        return "C-"
    else:
        return "F"


def _get_score_color(score):
    """Get color fill based on score value."""
    if score >= 70:
        return PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")  # Light green
    elif score >= 60:
        return PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")  # Yellow
    else:
        return PatternFill(start_color="FFB6C1", end_color="FFB6C1", fill_type="solid")  # Red


def _format_main_sheet(worksheet, df, sorted_clos, sorted_plos):
    """Apply formatting to the main results sheet."""
    
    # Format headers
    header_fill = PatternFill(start_color="6B2C91", end_color="6B2C91", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    
    for column in range(1, len(df.columns) + 1):
        cell = worksheet.cell(row=1, column=column)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    
    # Format data cells with color coding
    for row in range(2, len(df) + 2):
        for column in range(2, len(df.columns)):  # Skip Overall Grade column for score coloring
            cell = worksheet.cell(row=row, column=column)
            value = cell.value
            
            if isinstance(value, (int, float)):
                cell.fill = _get_score_color(value)
                cell.alignment = Alignment(horizontal="center")
                
                # Special formatting for zero scores - keep red
                if value == 0:
                    cell.fill = PatternFill(start_color="FFB6C1", end_color="FFB6C1", fill_type="solid")  # Red background
        
        # Format Overall Grade column differently
        grade_cell = worksheet.cell(row=row, column=len(df.columns))
        grade_cell.alignment = Alignment(horizontal="center")
        # You can add special formatting for grades if needed
    
    # Auto-adjust column widths
    for column in worksheet.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 20)
        worksheet.column_dimensions[column_letter].width = adjusted_width


def _create_summary_sheet(writer, clo_scores, plo_scores, grades, sorted_clos, sorted_plos):
    """Create and format the summary sheet with performance analytics."""
    
    summary_data = []
    
    # CLO summary (include all CLOs, even those with 0 scores)
    summary_data.append(['CLO Performance Summary'])
    summary_data.append([])
    summary_data.append(['CLO', 'Average Score', 'Students Above 80%', 'Students Below 60%'])
    
    for clo in sorted_clos:
        scores = [clo_scores[student].get(clo, 0) for student in clo_scores]
        avg_score = sum(scores) / len(scores) if scores else 0
        above_80 = sum(1 for score in scores if score >= 80)
        below_60 = sum(1 for score in scores if score < 60)
        summary_data.append([clo, f"{avg_score:.1f}", above_80, below_60])
    
    summary_data.append([])
    summary_data.append(['PLO Performance Summary'])
    summary_data.append([])
    summary_data.append(['PLO', 'Average Score', 'Students Above 80%', 'Students Below 60%'])
    
    for plo in sorted_plos:
        scores = [plo_scores[student][plo] for student in plo_scores if plo in plo_scores[student]]
        if scores:
            avg_score = sum(scores) / len(scores)
            above_80 = sum(1 for score in scores if score >= 80)
            below_60 = sum(1 for score in scores if score < 60)
            summary_data.append([plo, f"{avg_score:.1f}", above_80, below_60])
    
    # Add Overall Grade Summary
    summary_data.append([])
    summary_data.append(['Overall Grade Summary'])
    summary_data.append([])
    summary_data.append(['Grade Range', 'Number of Students'])
    
    # Count students in each grade range
    grade_ranges = {'A+ (95-100%)': 0, 'A (90-94%)': 0, 'A- (85-89%)': 0, 
                   'B+ (80-84%)': 0, 'B (75-79%)': 0, 'B- (70-74%)': 0,
                   'C+ (67-69%)': 0, 'C (63-66%)': 0, 'C- (60-62%)': 0, 'F (<60%)': 0}
    
    for student_grade in grades.values():
        letter_grade = _calculate_letter_grade(student_grade)
        if letter_grade == 'A+':
            grade_ranges['A+ (95-100%)'] += 1
        elif letter_grade == 'A':
            grade_ranges['A (90-94%)'] += 1
        elif letter_grade == 'A-':
            grade_ranges['A- (85-89%)'] += 1
        elif letter_grade == 'B+':
            grade_ranges['B+ (80-84%)'] += 1
        elif letter_grade == 'B':
            grade_ranges['B (75-79%)'] += 1
        elif letter_grade == 'B-':
            grade_ranges['B- (70-74%)'] += 1
        elif letter_grade == 'C+':
            grade_ranges['C+ (67-69%)'] += 1
        elif letter_grade == 'C':
            grade_ranges['C (63-66%)'] += 1
        elif letter_grade == 'C-':
            grade_ranges['C- (60-62%)'] += 1
        else:
            grade_ranges['F (<60%)'] += 1
    
    for grade_range, count in grade_ranges.items():
        summary_data.append([grade_range, count])
    
    # Write summary data
    summary_df = pd.DataFrame(summary_data)
    summary_df.to_excel(writer, sheet_name='Summary', index=False, header=False)
    
    # Format summary sheet
    summary_ws = writer.sheets['Summary']
    
    # Bold the summary headers
    header_rows = [1, len(sorted_clos) + 6, len(sorted_clos) + len(sorted_plos) + 12]  # CLO, PLO, and Grade summary headers
    for row in header_rows:
        try:
            cell = summary_ws.cell(row=row, column=1)
            cell.font = Font(bold=True, size=14)
        except:
            pass
    
    # Bold the column headers
    column_header_rows = [3, len(sorted_clos) + 8, len(sorted_clos) + len(sorted_plos) + 14]  # CLO, PLO, and Grade column headers
    for row in column_header_rows:
        try:
            for col in range(1, 5):
                cell = summary_ws.cell(row=row, column=col)
                cell.font = Font(bold=True)
        except:
            pass
    
    # Auto-adjust column widths
    for column in summary_ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 25)
        summary_ws.column_dimensions[column_letter].width = adjusted_width


def export_clo_plo_results(clo_scores, plo_scores, grades, data_dict, output_dir=None):
    """
    Main export function - creates Excel file with CLO/PLO results.
    
    Args:
        clo_scores (dict): CLO scores for all students
        plo_scores (dict): PLO scores for all students
        grades (dict): Final grades for all students (from terminal calculation)
        data_dict (dict): Original data from data.py
        output_dir (str, optional): Directory to save file in
        
    Returns:
        str: Path to created Excel file
        
    Raises:
        Exception: If Excel creation fails
    """
    try:
        output_file = None
        if output_dir:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = f"{output_dir}/CLO_PLO_Results_{timestamp}.xlsx"
        
        file_path = create_excel_output(clo_scores, plo_scores, grades, data_dict, output_file)
        print(f"✅ Excel export successful: {file_path}")
        return file_path
        
    except Exception as e:
        print(f"❌ Excel export failed: {str(e)}")
        raise e


if __name__ == "__main__":
    # Test function - can be used for standalone testing
    print("Excel Exporter Module")
    print("This module should be imported, not run directly.")