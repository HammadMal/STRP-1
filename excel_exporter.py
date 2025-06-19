"""
Excel Export Module for CLO/PLO Mapping Tool
Handles creation of formatted Excel reports with student performance data.
"""

import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from datetime import datetime


def create_excel_output(clo_scores, plo_scores, data_dict, output_file=None):
    """
    Create an Excel file with CLO and PLO scores in a formatted table.
    
    Note: Overall grade calculation temporarily removed - to be added after 
    confirming calculation method with instructor.
    
    Args:
        clo_scores (dict): Dictionary of CLO scores for each student
        plo_scores (dict): Dictionary of PLO scores for each student
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
        
        # TODO: Add overall grade calculation after confirming method with sir
        # row_data['Overall Grade'] = "TBD"
        
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
        _create_summary_sheet(writer, clo_scores, plo_scores, sorted_clos, sorted_plos)
    
    return output_file


def _calculate_overall_grade_like_sir(clo_scores, student_raw_scores):
    """
    Calculate overall grade using sir's weighted assessment method.
    
    Args:
        clo_scores (dict): CLO scores for the student
        student_raw_scores (dict): Raw assessment scores for bonus calculation
    
    Returns:
        float: Overall grade matching sir's calculation
    """
    # Assessment weights from the course structure
    assessment_weights = {
        'CLO 1': 0.15,  # Q1 weight (15%)
        'CLO 4': 0.15,  # Quiz 3 weight (15%)
        'CLO 5': 0.35   # Quiz 2 weight (35%)
    }
    
    # Calculate weighted sum for active CLOs
    weighted_sum = 0
    total_weight = 0
    
    for clo, weight in assessment_weights.items():
        if clo in clo_scores:
            weighted_sum += clo_scores[clo] * weight
            total_weight += weight
    
    # Calculate base overall score (scale to 100%)
    base_score = (weighted_sum / total_weight) if total_weight > 0 else 0
    
    # Add bonus points if available
    bonus = 0
    if student_raw_scores and 'Bonus' in student_raw_scores:
        try:
            bonus = float(student_raw_scores['Bonus'])
        except (ValueError, TypeError):
            bonus = 0
    
    overall_score = base_score + bonus
    
    return round(overall_score, 1)


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
    if score >= 80:
        return PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")  # Light green
    elif score >= 60:
        return PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")  # Light yellow
    elif score >= 40:
        return PatternFill(start_color="FFB347", end_color="FFB347", fill_type="solid")  # Light orange
    else:
        return PatternFill(start_color="FFB6C1", end_color="FFB6C1", fill_type="solid")  # Light red


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
        for column in range(2, len(df.columns) + 1):  # Include all columns now (no Overall Grade to skip)
            cell = worksheet.cell(row=row, column=column)
            value = cell.value
            
            if isinstance(value, (int, float)):
                cell.fill = _get_score_color(value)
                cell.alignment = Alignment(horizontal="center")
                
                # Special formatting for zero scores (like CLO 2, CLO 3)
                if value == 0:
                    cell.fill = PatternFill(start_color="FF6B6B", end_color="FF6B6B", fill_type="solid")  # Red background
    
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


def _create_summary_sheet(writer, clo_scores, plo_scores, sorted_clos, sorted_plos):
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
    
    # Write summary data
    summary_df = pd.DataFrame(summary_data)
    summary_df.to_excel(writer, sheet_name='Summary', index=False, header=False)
    
    # Format summary sheet
    summary_ws = writer.sheets['Summary']
    
    # Bold the summary headers
    header_rows = [1, len(sorted_clos) + 6]  # CLO and PLO summary headers
    for row in header_rows:
        try:
            cell = summary_ws.cell(row=row, column=1)
            cell.font = Font(bold=True, size=14)
        except:
            pass
    
    # Bold the column headers
    column_header_rows = [3, len(sorted_clos) + 8]  # CLO and PLO column headers
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

    # Add calculation explanation sheet
    _create_calculation_explanation_sheet(writer)


def _create_calculation_explanation_sheet(writer):
    """Create a sheet explaining the overall grade calculation method."""
    
    explanation_data = [
        ['Overall Grade Calculation Method'],
        [''],
        ['This Excel file uses the same calculation method as the original course spreadsheet.'],
        [''],
        ['Formula: Overall Grade = (CLO1×15% + CLO4×15% + CLO5×35%) ÷ 65% + Bonus'],
        [''],
        ['Explanation:'],
        ['• CLO 1 weight: 15% (from Q1 assessment)'],
        ['• CLO 4 weight: 15% (from Quiz 3 assessment)'],
        ['• CLO 5 weight: 35% (from Quiz 2 assessment)'],
        ['• Total active weight: 65% (CLO 2 and CLO 3 had no assessments)'],
        ['• Division by 65% scales the score to represent full course performance'],
        ['• Bonus points are added after the weighted calculation'],
        [''],
        ['This method ensures:'],
        ['• Fair weighting based on actual assessment importance'],
        ['• No penalty for unassessed CLOs (CLO 2, CLO 3)'],
        ['• Consistency with instructor\'s grading system'],
        [''],
        ['Generated by Habib University CLO/PLO Mapping Tool']
    ]
    
    # Write explanation data
    explanation_df = pd.DataFrame(explanation_data)
    explanation_df.to_excel(writer, sheet_name='Calculation Method', index=False, header=False)
    
    # Format explanation sheet
    explanation_ws = writer.sheets['Calculation Method']
    
    # Bold the title
    title_cell = explanation_ws.cell(row=1, column=1)
    title_cell.font = Font(bold=True, size=16, color="6B2C91")
    
    # Bold section headers
    for row in [7, 15]:  # "Explanation:" and "This method ensures:"
        try:
            cell = explanation_ws.cell(row=row, column=1)
            cell.font = Font(bold=True, size=12)
        except:
            pass
    
    # Adjust column width
    explanation_ws.column_dimensions['A'].width = 80


def export_clo_plo_results(clo_scores, plo_scores, data_dict, output_dir=None):
    """
    Main export function - creates Excel file with CLO/PLO results.
    
    Args:
        clo_scores (dict): CLO scores for all students
        plo_scores (dict): PLO scores for all students  
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
        
        file_path = create_excel_output(clo_scores, plo_scores, data_dict, output_file)
        print(f"✅ Excel export successful: {file_path}")
        return file_path
        
    except Exception as e:
        print(f"❌ Excel export failed: {str(e)}")
        raise e


if __name__ == "__main__":
    # Test function - can be used for standalone testing
    print("Excel Exporter Module")
    print("This module should be imported, not run directly.")