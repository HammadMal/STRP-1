import pandas as pd
import re


# === Load Excel File ===
def load_excel(file_path: str):
    try:
        df = pd.read_excel(file_path, sheet_name='Data', header=None, engine='openpyxl')

        print("[✓] Loaded Excel file")
        return df
    except Exception as e:
        print("[!] Failed to load Excel:", e)
        return None

# === Clean Full DataFrame ===
def clean_dataframe(df):
    # 1. Remove hidden or undefined characters
    df = df.applymap(lambda x: clean_cell(x))

    # 2. Drop completely empty rows and columns
    df.dropna(axis=0, how='all', inplace=True)  # rows
    df.dropna(axis=1, how='all', inplace=True)  # columns

    print("[✓] Cleaned hidden characters and empty rows/cols")
    return df

# === Clean a single cell ===
def clean_cell(value):
    if pd.isnull(value):
        return None
    val = str(value)
    # Remove non-breaking spaces, zero-width spaces, and control chars
    val = val.replace('\u200b', '').replace('\xa0', ' ')
    val = re.sub(r'[\r\n\t]', ' ', val)        # newlines, tabs
    val = re.sub(r'[^\x00-\x7F]+', '', val)    # remove non-ASCII
    val = val.strip()
    return val

def drop_short_rows(df, char_limit=2):
    def row_char_count(row):
        # Combine all cells into one string, strip spaces, and count characters
        all_text = ''.join([str(cell).strip() for cell in row if pd.notnull(cell)])
        return len(all_text)

    # Keep only rows with more than `char_limit` characters
    df = df[df.apply(row_char_count, axis=1) > char_limit]
    return df

# === Main Driver ===
def preprocess_excel(file_path):
    df = load_excel(file_path)
    if df is None:
        return None

    df = clean_dataframe(df)
    df= drop_short_rows(df, char_limit=2)
    return df.loc[:, df.isnull().mean() < 0.7]

df_clean = preprocess_excel("C:/Users/AWAIS GILL/Desktop/OBE STRP/OBE Program/Python files/1911-EE-000-T1.xlsx")


print(df_clean)
