import argparse
import re
from pathlib import Path
import pandas as pd

def _column_letter_to_index(s: str) -> int:
    s = s.strip().upper()
    idx = 0
    for ch in s:
        idx = idx * 26 + (ord(ch) - 64)
    return idx - 1

def sanitize_id(val) -> str | None:
    if pd.isna(val): return None
    s = str(val).strip()
    m = re.search(r"\d+", s)
    return m.group(0) if m else None

def load_tic_list(excel_path: str, sheet: str | int = 0, column: str | int = "A") -> set[str]:
    df = pd.read_excel(excel_path, sheet_name=sheet, dtype=str, header=None)
    col_idx = _column_letter_to_index(column) if isinstance(column, str) and not column.isdigit() else int(column)
    series = df.iloc[:, col_idx].map(sanitize_id).dropna()
    return set(series.tolist())

def load_txt_df(txt_path: str) -> pd.DataFrame:
    try:
        df = pd.read_csv(txt_path, sep=r"\s+", engine="python", dtype={"TIC_ID": str},
                         usecols=["TIC_ID", "RA_OBJ", "DEC_OBJ"])
    except Exception:
        rows = []
        with open(txt_path, "r", encoding="utf-8", errors="ignore") as f:
            header_seen = False
            for line in f:
                line = line.strip()
                if not line: continue
                if not header_seen and "TIC" in line and "RA" in line and "DEC" in line:
                    header_seen = True
                    continue
                parts = re.split(r"\s+", line)
                if len(parts) < 3: continue
                tic = sanitize_id(parts[0])
                try:
                    ra = float(parts[1])
                    dec = float(parts[2])
                except ValueError:
                    continue
                if tic:
                    rows.append({"TIC_ID": tic, "RA_OBJ": ra, "DEC_OBJ": dec})
        df = pd.DataFrame(rows, columns=["TIC_ID", "RA_OBJ", "DEC_OBJ"])
    df["TIC_ID"] = df["TIC_ID"].map(sanitize_id)
    df = df.dropna(subset=["TIC_ID"])
    return df[["TIC_ID", "RA_OBJ", "DEC_OBJ"]]

def main():
    p = argparse.ArgumentParser()
    p.add_argument("-e", "--excel", required=True)
    p.add_argument("-t", "--txt", required=True)
    p.add_argument("-o", "--output", default="tic_ra_dec_matched.xlsx")
    p.add_argument("--sheet", default=0)
    p.add_argument("-c", "--column", default="A")
    args = p.parse_args()

    tic_set = load_tic_list(args.excel, sheet=args.sheet, column=args.column)
    txt_df = load_txt_df(args.txt)
    matched = txt_df[txt_df["TIC_ID"].isin(tic_set)].copy()
    matched = matched.sort_values("TIC_ID", key=lambda s: s.astype(str))
    out_path = Path(args.output)

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        matched.to_excel(writer, index=False, header=False, sheet_name="Matched", float_format="%.15f")

    not_found = sorted(tic_set - set(matched["TIC_ID"].astype(str).tolist()))
    if not_found:
        nf_path = out_path.with_name(out_path.stem + "_NOT_FOUND.txt")
        with open(nf_path, "w", encoding="utf-8") as f:
            f.write("\n".join(not_found))
    print(f"Знайдено {len(matched)} із {len(tic_set)}. Результат: {out_path}")

if __name__ == "__main__":
    main()
