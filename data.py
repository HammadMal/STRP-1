import pandas as pd
import re
import sys
import json

# === Student ID Validation and Formatting ===
def validate_and_format_student_id(student_id_raw):
    """
    Validate and format student ID to Habib University email format.
    
    Expected format: hm08298@st.habib.edu.pk
    Where:
    - First two characters are initials (letters)
    - Next 5 characters are numbers
    - Followed by @st.habib.edu.pk
    
    Args:
        student_id_raw: Raw student ID from Excel
        
    Returns:
        str: Formatted email address
        
    Raises:
        ValueError: If ID format is invalid
    """
    if pd.isnull(student_id_raw):
        raise ValueError("Student ID cannot be empty or null")
    
    student_id = str(student_id_raw).strip().lower()
    
    # If already in email format, validate it
    if '@st.habib.edu.pk' in student_id:
        # Extract the part before @
        local_part = student_id.split('@')[0]
        
        # Validate the local part format
        if not re.match(r'^[a-z]{2}\d{5}$', local_part):
            raise ValueError(f"Invalid student email format: {student_id}. Expected format: xx12345@st.habib.edu.pk (where xx are initials and 12345 is student number)")
        
        # Ensure it ends with the correct domain
        if not student_id.endswith('@st.habib.edu.pk'):
            raise ValueError(f"Invalid domain in student email: {student_id}. Must end with @st.habib.edu.pk")
        
        return student_id
    
    # If not in email format, try to format it
    else:
        # Check if it matches the pattern: 2 letters + 5 digits
        if re.match(r'^[a-z]{2}\d{5}$', student_id):
            return f"{student_id}@st.habib.edu.pk"
        
        # Handle common variations
        # Remove any spaces or special characters except letters and numbers
        cleaned_id = re.sub(r'[^a-z0-9]', '', student_id)
        
        # Check if cleaned version matches the pattern
        if re.match(r'^[a-z]{2}\d{5}$', cleaned_id):
            return f"{cleaned_id}@st.habib.edu.pk"
        
        # Try to extract initials and numbers from mixed format
        letters = re.findall(r'[a-z]', student_id)
        numbers = re.findall(r'\d', student_id)
        
        if len(letters) >= 2 and len(numbers) >= 5:
            # Take first 2 letters and first 5 numbers
            formatted_id = ''.join(letters[:2]) + ''.join(numbers[:5])
            if re.match(r'^[a-z]{2}\d{5}$', formatted_id):
                return f"{formatted_id}@st.habib.edu.pk"
        
        # If all else fails, raise an error with helpful message
        raise ValueError(f"Invalid student ID format: '{student_id_raw}'. Expected format examples: 'hm08298', 'ab12345', or 'hm08298@st.habib.edu.pk'")


def validate_all_student_ids(student_scores):
    """
    Validate and format all student IDs in the dataset.
    
    Args:
        student_scores (dict): Dictionary with student IDs as keys
        
    Returns:
        dict: Dictionary with formatted student IDs as keys
        
    Raises:
        ValueError: If any student ID is invalid
    """
    formatted_scores = {}
    invalid_ids = []
    
    for student_id_raw, scores in student_scores.items():
        try:
            formatted_id = validate_and_format_student_id(student_id_raw)
            formatted_scores[formatted_id] = scores
        except ValueError as e:
            invalid_ids.append(f"Row with ID '{student_id_raw}': {str(e)}")
    
    if invalid_ids:
        error_message = "[!] Invalid student ID format(s) found:\n\n" + "\n".join(invalid_ids)
        error_message += "\n\nValid formats:"
        error_message += "\n- hm08298 (initials + 5-digit number)"
        error_message += "\n- ab12345 (any 2 letters + 5 digits)"
        error_message += "\n- hm08298@st.habib.edu.pk (full email format)"
        error_message += "\n\nPlease fix the student IDs in your Excel file and try again."
        raise ValueError(error_message)
    
    return formatted_scores


# === Load Excel ===
def load_excel(file_path: str):
    try:
        xlsx = pd.ExcelFile(file_path, engine='openpyxl')
        sheet_name = 'Data' if 'Data' in xlsx.sheet_names else xlsx.sheet_names[0]
        df = pd.read_excel(xlsx, sheet_name=sheet_name, header=None)

        print("[OK] Loaded Excel file")
        return df
    except Exception as e:
        print("[!] Failed to load Excel:", e)
        return None

# === Clean individual cell ===
def clean_cell(value):
    if pd.isnull(value):
        return None
    val = str(value)
    val = val.replace('\u200b', '').replace('\xa0', ' ')
    val = re.sub(r'[\r\n\t]', ' ', val)
    val = re.sub(r'[^\x00-\x7F]+', '', val)
    return val.strip()

# === Clean entire DataFrame ===
def clean_dataframe(df):
    df = df.applymap(lambda x: clean_cell(x))
    df.dropna(axis=0, how='all', inplace=True)
    df.dropna(axis=1, how='all', inplace=True)
    return df

# === Drop nearly empty rows ===
def drop_short_rows(df, char_limit=2):
    def row_char_count(row):
        all_text = ''.join([str(cell).strip() for cell in row if pd.notnull(cell)])
        return len(all_text)
    return df[df.apply(row_char_count, axis=1) > char_limit]

# === Automatically find module and student row indices ===
def find_data_rows(df):
    module_row = None
    for i in range(len(df)):
        cell_val = str(df.iloc[i, 0]).strip().lower()
        if isinstance(cell_val, str) and re.search(r'\bmodule(s)?\b', cell_val, re.IGNORECASE):
            module_row = i
            break
    if module_row is None:
        print("[!] 'Modules' row not found, trying default fallback row 10")
        module_row = 10

    return module_row, module_row + 1, module_row + 2, module_row + 3

# === Extract CLOs, PLOs, Modules, Scores ===
def extract_clo_plo_data(df):
    clos = {}
    clo_to_plo = {}
    student_scores = {}

    all_defined_clos = {}
    
    for i in range(len(df)):
        clo_cell = df.iloc[i, 0]
        if pd.notnull(clo_cell) and str(clo_cell).strip().startswith('CLO'):
            clo_id = str(clo_cell).strip()
            description = df.iloc[i, 1] if len(df.columns) > 1 else ""
            ldl = df.iloc[i, 2] if len(df.columns) > 2 else ""
            plo_map = df.iloc[i, 3] if len(df.columns) > 3 else ""
            
            if pd.notnull(description) and str(description).strip() and len(str(description).strip()) > 10:
                all_defined_clos[clo_id] = {"description": description, "LDL": ldl}
                
                if isinstance(plo_map, str) and ";" in plo_map:
                    try:
                        plo_id, weight = plo_map.split(";")
                        weight = float(weight)  # ✅ float instead of int
                        clo_to_plo[clo_id] = {"PLO": f"PLO {plo_id.strip()}", "weight": weight}
                    except Exception as e:
                        print(f"[!] Failed to parse PLO mapping for {clo_id}: {plo_map} ({e})")

    clos = all_defined_clos

    module_row, clo_map_row, max_score_row, student_start_row = find_data_rows(df)

    module_names = df.iloc[module_row, 1:].tolist()
    clo_mapping = df.iloc[clo_map_row, 1:].tolist()
    max_scores = df.iloc[max_score_row, 1:].tolist()

    clo_assessments = {}
    for i, (module, mapping, max_score) in enumerate(zip(module_names, clo_mapping, max_scores)):
        if isinstance(mapping, str) and ";" in mapping:
            try:
                clo_index, weight = mapping.split(";")
                clo_id = f"CLO {clo_index.strip()}"
                if clo_id not in clo_assessments:
                    clo_assessments[clo_id] = []
                clo_assessments[clo_id].append({
                    "module": module,
                    "max_score": float(max_score),  # ✅ now handles 15.0 or 12.5
                    "weight": float(weight)         # ✅ supports weights like 10.5
                })
            except Exception as e:
                print(f"[!] Failed to parse CLO assessment mapping: {mapping} ({e})")

    # Extract student scores with ID validation
    raw_student_scores = {}
    for i in range(student_start_row, df.shape[0]):
        student_id_raw = df.iloc[i, 0]
        if pd.isnull(student_id_raw):
            continue
        
        scores = df.iloc[i, 1:].tolist()
        raw_student_scores[student_id_raw] = {
            module_names[j]: scores[j] for j in range(len(module_names))
            if pd.notnull(scores[j])
        }
    
    # Validate and format all student IDs
    try:
        student_scores = validate_all_student_ids(raw_student_scores)
        print(f"[OK] Successfully validated and formatted {len(student_scores)} student IDs")
        
        # Print formatted IDs for verification
        print("\n[INFO] Formatted Student IDs:")
        for formatted_id in sorted(student_scores.keys()):
            print(f"  - {formatted_id}")
        
    except ValueError as e:
        print(f"\n{str(e)}")
        raise e

    return clos, clo_to_plo, clo_assessments, student_scores

# === Run Full Preprocessing ===
def preprocess_excel_and_extract(file_path):
    df = load_excel(file_path)
    if df is None:
        return

    df = clean_dataframe(df)
    df = drop_short_rows(df, char_limit=2)
    df = df.loc[:, df.isnull().mean() < 0.7]

    try:
        clos, clo_to_plo, clo_assessments, student_scores = extract_clo_plo_data(df)
    except Exception as e:
        print(f"[!] Data extraction failed: {e}")
        return

    # Output as structured JSON
    print(json.dumps({
        "clos": clos,
        "clo_to_plo": clo_to_plo,
        "clo_assessments": clo_assessments,
        "student_scores": student_scores
    }, indent=2))

# === Run if executed directly ===
if __name__ == "__main__":
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
    else:
        file_path = "example_file.xlsx"  # fallback
    preprocess_excel_and_extract(file_path)