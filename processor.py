import re
import pandas as pd
import jaconv
from pathlib import Path

# ===== Day1 と同じパラメータ =====
DATA_DIR = r".\data"
FILES = [
    "Ship-to_VBU-GBM_DueDate_2025-06-02.xlsx",
    "展示会_人とくるま展横浜.xlsx",
    "電源セミナー申込者.xlsx",
]
TARGET_SHEET_NAME = {
    "Ship-to_VBU-GBM_DueDate_2025-06-02.xlsx": ["JP Assignment List"],
    "展示会_人とくるま展横浜.xlsx": ["ヒアリングシート"],
    "電源セミナー申込者.xlsx": ["power-seminar-2025"]
}
LIKELY_HEADER_TOKENS = [
    "ECC Ship-to", "Cust", "Cust_Name", "Address_1", "City", "Country",
    "Sales_Coverage", "End_Mkt_Segment", "メールアドレス", "会社名", "姓", "姓（かな）",
    "名", "名（かな）", "郵便番号", "都道府県", "電話番号", "Organization Name",
    "EmailAddress", "Given Name", "Family Name", "State Province", "Postal Code",
    "Telephone"
]

# ===== ユーティリティ（列名正規化） =====
def clean_cols(cols: pd.Index) -> pd.Index:
    s = pd.Index(cols).astype(str)
    s = s.str.replace(r"[\r\n]", "", regex=True)  # 改行除去
    s = s.str.replace("　", "", regex=True)      # 全角スペース除去
    s = s.str.strip()
    return s

def norm_key(s: str) -> str:
    """列名をトークン照合用に正規化（全角→半角、小文字化、空白/記号の除去）"""
    if not isinstance(s, str):
        return ""
    x = jaconv.z2h(s, ascii=True, digit=True).lower().strip()
    x = re.sub(r"[\s_/\\\-]+", "", x)
    return x

TOKEN_KEYS = {norm_key(t) for t in LIKELY_HEADER_TOKENS}

# ===== 前処理（会社名/都道府県/メール） =====
PREF_MAP = {
    "北海道":"Hokkaido","青森県":"Aomori","岩手県":"Iwate","宮城県":"Miyagi","秋田県":"Akita",
    "山形県":"Yamagata","福島県":"Fukushima","茨城県":"Ibaraki","栃木県":"Tochigi","群馬県":"Gunma",
    "埼玉県":"Saitama","千葉県":"Chiba","東京都":"Tokyo","神奈川県":"Kanagawa","新潟県":"Niigata",
    "富山県":"Toyama","石川県":"Ishikawa","福井県":"Fukui","山梨県":"Yamanashi","長野県":"Nagano",
    "岐阜県":"Gifu","静岡県":"Shizuoka","愛知県":"Aichi","三重県":"Mie","滋賀県":"Shiga",
    "京都府":"Kyoto","大阪府":"Osaka","兵庫県":"Hyogo","奈良県":"Nara","和歌山県":"Wakayama",
    "鳥取県":"Tottori","島根県":"Shimane","岡山県":"Okayama","広島県":"Hiroshima","山口県":"Yamaguchi",
    "徳島県":"Tokushima","香川県":"Kagawa","愛媛県":"Ehime","高知県":"Kochi","福岡県":"Fukuoka",
    "佐賀県":"Saga","長崎県":"Nagasaki","熊本県":"Kumamoto","大分県":"Oita","宮崎県":"Miyazaki",
    "鹿児島県":"Kagoshima","沖縄県":"Okinawa",
    "東京":"Tokyo","神奈川":"Kanagawa","大阪":"Osaka","愛知":"Aichi","埼玉":"Saitama","千葉":"Chiba","福岡":"Fukuoka"
}

def normalize_company(name: str) -> str:
    """会社名の正規化（全半角・法人格・記号・空白の統一）"""
    if not isinstance(name, str):
        return ""
    s = jaconv.z2h(name, ascii=True, digit=True)
    s = re.sub(r"\s+", " ", s).strip()
    suffix_patterns = [
        r"\(株\)", r"\(有\)", r"株式会社", r"有限会社", r"合同会社",
        r"Inc\.?", r"Co\.?,?\s*Ltd\.?", r"Ltd\.?", r"LLC", r"Limited", r"Company", r"（株）"
    ]
    for p in suffix_patterns:
        s = re.sub(p, "", s, flags=re.IGNORECASE)
    s = re.sub(r"[・\.\,\-\/\(\)（）]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s.lower()

def normalize_prefecture(pref: str) -> str:
    """都道府県の英語化（マップにあれば置換）"""
    if not isinstance(pref, str):
        return ""
    p = pref.strip()
    return PREF_MAP.get(p, p)

def extract_domain(email: str) -> str:
    """メールアドレスからドメインを抽出"""
    if not isinstance(email, str):
        return ""
    m = re.search(r"@([A-Za-z0-9\.\-]+)$", email.strip())
    return m.group(1).lower() if m else ""

def ensure_row_id(df: pd.DataFrame) -> pd.DataFrame:
    """Row_ID が無い場合は連番を付与"""
    if "Row_ID" not in df.columns:
        df = df.copy()
        df["Row_ID"] = range(1, len(df)+1)
    return df

# ===== ヘッダー検出（トークンに一致する列名が出る行を表頭とみなす）=====
def detect_header_row(path: Path, sheet: str, max_check_rows: int = 8) -> int:
    for i in range(max_check_rows):
        try:
            df_tmp = pd.read_excel(path, sheet_name=sheet, header=i, nrows=1)
        except Exception:
            continue
        cols = clean_cols(df_tmp.columns)
        col_keys = {norm_key(c) for c in cols}
        if TOKEN_KEYS & col_keys:
            return i
    return 0

# ===== トークン列 + 対応する正規化列だけを抽出 =====
NORMALIZED_MAP = {
    # 参照側
    "Cust_Name": ["cust_norm"],
    "Cust": ["alias_norm"],
    "State/Prefecture": ["state_norm"],
    "PostalCode": ["postal_norm"],
    # 入力側
    "会社名": ["company_norm"],
    "Organization Name": ["company_norm"],
    "都道府県": ["pref_en"],
    "State Province": ["pref_en"],
    "都道府県（単体）": ["pref_en"],
    "メールアドレス": ["email_domain"],
    "Email": ["email_domain"],
    "EmailAddress": ["email_domain"],
}

def filter_columns_with_norm(df: pd.DataFrame) -> pd.DataFrame:
    """tokens に一致する元列 + その列に対応する正規化列をまとめて抽出"""
    base_cols = [c for c in df.columns if norm_key(c) in TOKEN_KEYS]
    extra_cols = []
    for base in base_cols:
        for k, norms in NORMALIZED_MAP.items():
            if norm_key(base) == norm_key(k):
                for ncol in norms:
                    if ncol in df.columns:
                        extra_cols.append(ncol)
    keep_cols = base_cols + [c for c in extra_cols if c not in base_cols]
    return df[keep_cols] if keep_cols else df

# ===== 1ファイル処理：読み込み→正規化列作成→抽出→保存 =====
def preprocess_file(fname: str, sheet: str, out_dir: Path):
    path = Path(DATA_DIR) / fname
    header_row = detect_header_row(path, sheet)
    df = pd.read_excel(path, sheet_name=sheet, header=header_row)
    df.columns = clean_cols(df.columns)

    # 参照 or 入力で正規化列を用意
    if "Cust_Name" in df.columns or "ECC Ship-to" in df.columns:
        df_full = df.copy()
        df_full["cust_norm"]  = df_full.get("Cust_Name","").apply(normalize_company)
        df_full["alias_norm"] = df_full.get("Cust","").apply(normalize_company)
        if "State/Prefecture" in df_full.columns:
            df_full["state_norm"] = df_full["State/Prefecture"].astype(str).str.strip()
        if "PostalCode" in df_full.columns:
            df_full["postal_norm"] = df_full["PostalCode"].astype(str).str.replace("-","").str.strip()
    else:
        df_full = ensure_row_id(df.copy())
        ccol = next((c for c in ["会社名","Organization Name"] if c in df_full.columns), None)
        pcol = next((c for c in ["都道府県","State Province","都道府県（単体）"] if c in df_full.columns), None)
        ecol = next((c for c in ["メールアドレス","Email","EmailAddress"] if c in df_full.columns), None)
        df_full["company_norm"] = df_full[ccol].apply(normalize_company) if ccol else ""
        df_full["pref_en"]      = df_full[pcol].apply(normalize_prefecture) if pcol else ""
        df_full["email_domain"] = df_full[ecol].apply(extract_domain) if ecol else ""

    # トークン列 + 正規化列のみを抽出して 1 ファイル出力
    df_for_match = filter_columns_with_norm(df_full)
    out_dir.mkdir(exist_ok=True)
    out_path = out_dir / f"{Path(fname).stem}_for_match.xlsx"
    df_for_match.to_excel(out_path, index=False)
    print(f"[OK] {fname} | {sheet} → {out_path}")

def main():
    out_dir = Path("./out")
    for fname in FILES:
        for sheet in TARGET_SHEET_NAME.get(fname, []):
            preprocess_file(fname, sheet, out_dir)

if __name__ == "__main__":
    main()
