import os
import re
import pandas as pd
DATA_DIR = r".\data"
FILES = [
    "Ship-to_VBU-GBM_DueDate_2025-06-02.xlsx",
    "展示会_人とくるま展横浜.xlsx",
    "電源セミナー申込者.xlsx",
]
TARGET_SHEET_NAME ={
    "Ship-to_VBU-GBM_DueDate_2025-06-02.xlsx":["JP Assignment List"],
    "展示会_人とくるま展横浜.xlsx":["ヒアリングシート"],
    "電源セミナー申込者.xlsx":["power-seminar-2025"]
}
LIKELY_HEADER_TOKENS = [
    "ECC Ship-to",
    "Cust",
    "Cust_Name",    
    "Address_1",
    "City",
    "Country",
    "Sales_Coverage",
    "End_Mkt_Segment",
    "メールアドレス",
    "会社名",
    "姓",
    "姓（かな）",
    "名",
    "名（かな）",
    "郵便番号",
    "都道府県",
    "電話番号",
    "Organization Name",
    "EmailAddress",
    "Given Name",
    "Family Name",
    "State Province",
    "Postal Code",
    "Telephone"
]
def clean_cols(cols: pd.Index) -> pd.Index:
    s = cols.astype(str)
    s = s.str.replace("\r|\n", "", regex=True).str.replace("　", "", regex=True).str.strip()
    return s

def find_header_row(xlsx_path: str, sheet_name) -> int:
  
    tmp = pd.read_excel(xlsx_path, sheet_name=sheet_name, header=None, nrows=8)
    for i in range(min(8, len(tmp))):
        row_vals = [str(v) for v in tmp.iloc[i].tolist()]
        row_join = " ".join(row_vals)
        if any(tok in row_join for tok in LIKELY_HEADER_TOKENS):
            return i
    return 0

def read_sheet_with_detected_header(xlsx_path: str, sheet_name):
    header_row = find_header_row(xlsx_path, sheet_name)
    df = pd.read_excel(xlsx_path, sheet_name=sheet_name, header=header_row)
    df.columns = clean_cols(df.columns)
    return header_row, df

def probe_file(xlsx_path: str, sheets_to_read=None) -> str:
 
    report_lines = []
    if not os.path.exists(xlsx_path):
        return f"[WARN] Missing file: {xlsx_path}\n"

    xl = pd.ExcelFile(xlsx_path)
    all_sheets = xl.sheet_names

    target_sheets = sheets_to_read if sheets_to_read else all_sheets

    report_lines.append(f"=== File: {xlsx_path} | target sheets: {target_sheets}\n")
    for sh in target_sheets:
        if sh not in all_sheets:
            report_lines.append(f"[WARN] Sheet '{sh}' not found in {xlsx_path}\n")
            continue
        header_row, df = read_sheet_with_detected_header(xlsx_path, sh)
        cols = df.columns.tolist()
        preview = df.head(10)
        report_lines.append(f"[Sheet] {sh} (header row = {header_row+1})\n")
        report_lines.append(f"Columns: {cols}\n")
        report_lines.append("Preview:\n")
        report_lines.append(preview.to_string(index=False))
        report_lines.append("\n" + "-"*80 + "\n")
    return "".join(report_lines)
def main():
    os.makedirs(DATA_DIR, exist_ok=True)
    report_path = os.path.join(DATA_DIR, "_probe_report.txt")
    all_text = []
    for fname in FILES:
        path = os.path.join(DATA_DIR, fname)
        sheets_to_read = TARGET_SHEET_NAME.get(fname, None)
        all_text.append(probe_file(path, sheets_to_read=sheets_to_read))
    with open(report_path, "w", encoding="utf-8") as f:
        f.write("\n".join(all_text))
    print(f"[OK] Probe finished. Report saved to: {report_path}")
    print("请把该文件的内容贴给我：data/_probe_report.txt（至少贴出每个文件的列名部分）")

if __name__ == "__main__":
    main()

