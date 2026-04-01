import pandas as pd

def _cell_to_text(value):
    if pd.isna(value):
        return ""
    if hasattr(value, "strftime"):
        return value.strftime("%d/%m/%Y")
    return str(value).strip()

def read_excel(file_path):
    df = pd.read_excel(file_path, keep_default_na=False)

    df.columns = df.columns.str.strip().str.lower()

    data = []
    for _, row in df.iterrows():
        data.append({
            "company_name": _cell_to_text(row.get("company_name", "")),
            "company_address": _cell_to_text(row.get("company_address", "")),
            "employer_identification_number": _cell_to_text(row.get("employer_identification_number", "")),
            "sector": _cell_to_text(row.get("sector", "")),
            "automation_tool": _cell_to_text(row.get("automation_tool", "")),
            "annual_automation_saving": _cell_to_text(row.get("annual_automation_saving", "")),
            "date_of_first_project": _cell_to_text(row.get("date_of_first_project", ""))
        })

    return data
