import pandas as pd
import re
import sys
import json

# === Load Excel ===
def load_excel(file_path: str):
    try:
        df = pd.read_excel(file_path, sheet_name='Data', header=None, engine='openpyxl')
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
        if "module" in cell_val or "modules" in cell_val:
            module_row = i
            break
    if module_row is None:
        raise ValueError("Could not find 'Modules' row in sheet.")
    
    return module_row, module_row + 1, module_row + 2, module_row + 3

# === Extract CLOs, PLOs, Modules, Scores ===
def extract_clo_plo_data(df):
    clos = {}
    clo_to_plo = {}
    student_scores = {}

    # Extract ALL CLO definitions from the Excel structure dynamically
    # This will capture CLOs that are defined in the course but may not have assessments
    all_defined_clos = {}
    
    # Dynamically find CLO definitions instead of hardcoding rows 2-6
    for i in range(len(df)):
        if i < len(df):
            clo_cell = df.iloc[i, 0]
            if pd.notnull(clo_cell) and str(clo_cell).strip().startswith('CLO'):
                clo_id = str(clo_cell).strip()
                description = df.iloc[i, 1] if i < len(df) and len(df.columns) > 1 else ""
                ldl = df.iloc[i, 2] if i < len(df) and len(df.columns) > 2 else ""
                plo_map = df.iloc[i, 3] if i < len(df) and len(df.columns) > 3 else ""
                
                # Validate that this is actually a CLO row (has proper description)
                if pd.notnull(description) and str(description).strip() and len(str(description).strip()) > 10:
                    # Store CLO definition
                    all_defined_clos[clo_id] = {"description": description, "LDL": ldl}
                    
                    # Store PLO mapping if it exists
                    if isinstance(plo_map, str) and ";" in plo_map:
                        plo_id, weight = plo_map.split(";")
                        clo_to_plo[clo_id] = {"PLO": f"PLO {plo_id.strip()}", "weight": int(weight)}

    # Use the comprehensive CLO definitions
    clos = all_defined_clos

    # Dynamically find module/start rows
    module_row, clo_map_row, max_score_row, student_start_row = find_data_rows(df)

    module_names = df.iloc[module_row, 1:].tolist()
    clo_mapping = df.iloc[clo_map_row, 1:].tolist()
    max_scores = df.iloc[max_score_row, 1:].tolist()

    # Map CLO to assessments (this might only include CLOs with actual assessments)
    clo_assessments = {}
    for i, (module, mapping, max_score) in enumerate(zip(module_names, clo_mapping, max_scores)):
        if isinstance(mapping, str) and ";" in mapping:
            clo_index, weight = mapping.split(";")
            clo_id = f"CLO {clo_index.strip()}"
            if clo_id not in clo_assessments:
                clo_assessments[clo_id] = []
            clo_assessments[clo_id].append({
                "module": module,
                "max_score": float(max_score),
                "weight": float(weight)
            })

    # Extract student scores
    for i in range(student_start_row, df.shape[0]):
        student_id_raw = df.iloc[i, 0]
        if pd.isnull(student_id_raw):
            continue
        student_id = str(student_id_raw).strip() 
        scores = df.iloc[i, 1:].tolist()
        student_scores[student_id] = {
        module_names[j]: scores[j] for j in range(len(module_names))
        if pd.notnull(scores[j])
    }

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


# returns marks/clo/plo in dictionary form