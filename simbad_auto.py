import time
import re
import unicodedata
import requests
import pandas as pd
from bs4 import BeautifulSoup
from datetime import datetime

INPUT_XLSX = "result.xlsx"
OUTPUT_XLSX = "result2.xlsx"

BASE_OBJ_URL = "https://simbad.u-strasbg.fr/simbad/sim-id"
OBJ_PARAMS = "NbIdent=1&Radius=2&Radius.unit=arcmin&submit=submit+id"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (compatible; TIC-Scraper/1.0)"
}

def build_obj_url(tic_num: str) -> str:
    return f"{BASE_OBJ_URL}?Ident=TIC+{tic_num}&{OBJ_PARAMS}"

def build_bib_url(tic_num: str) -> str:
    year = datetime.utcnow().year
    return f"{BASE_OBJ_URL}?Ident=TIC+{tic_num}&bibdisplay=refsum&bibyear1=1850&bibyear2={year}#lab_bib"

def fetch_references_text(html: str) -> str | None:
    soup = BeautifulSoup(html, "html.parser")
    for a in soup.find_all("a", href=True):
        href = a["href"]
        if "sim-ref" in href and ("Ident=" in href or "object=" in href):
            txt = a.get_text(" ", strip=True)
            if txt and txt.lower().startswith("references"):
                return txt
    cand = soup.find("a", string=lambda s: isinstance(s, str) and s.strip().lower().startswith("references"))
    if cand:
        return cand.get_text(" ", strip=True)
    text = soup.get_text(" ", strip=True)
    m = re.search(r"References\s*\([^)]*\)\s*\(Total\s*\d+\)", text, flags=re.I)
    if m:
        return m.group(0)
    return None

def _normalize_text(s: str) -> str:
    s = unicodedata.normalize("NFKC", s)
    s = s.replace("\u2013", "-").replace("\u2014", "-")
    s = re.sub(r"\s+", " ", s)
    return s.lower()

def find_terms_free(text: str, term_patterns: dict[str, re.Pattern]) -> bool:
    for pat in term_patterns.values():
        if pat.search(text):
            return True
    return False

# ⬇️ ТУТ ЗМІНА — видалено re.IGNORECASE, щоб шукати тільки ВЕЛИКІ літери (V*, EB*, ...)
def find_codes(text: str, codes: list[str]) -> bool:
    for code in codes:
        esc = re.escape(code)
        pat = re.compile(rf"(?<!\w){esc}(?!\w)")  # БЕЗ flags=re.IGNORECASE
        if pat.search(text):
            return True
    return False

FREE_TERMS = {
    "Variability": r"\bvariabilit[y|ies]\b|\bvariable\b",
    "Variable star": r"\bvariable star[s]?\b",
    "Light curve": r"\blight[- ]curve[s]?\b",
    "Periodic variable stars": r"\bperiodic variable\b",
    "Period determination": r"\bperiod determination\b",
    "Solar-like oscillations": r"\bsolar[- ]like oscillation[s]?\b",
    "Oscillating red giants": r"\boscillating red giant[s]?\b",
    "Rotation period": r"\brotation period[s]?\b",
    "Evolved massive stars": r"\bevolved massive star[s]?\b",
    "Stellar flares": r"\bstellar flare[s]?\b",
    "Flare stars": r"\bflare star[s]?\b",
    "Optical flares": r"\boptical flare[s]?\b",
    "Flaring M dwarfs": r"\bflaring m[- ]dwarf[s]?\b",
    "Flare activity": r"\bflare activity\b",
    "Superflare": r"\bsuperflare[s]?\b",
    "Outbursts": r"\boutburst[s]?\b",
    "Superoutbursts": r"\bsuperoutburst[s]?\b",
}

SIMBAD_CODES = [
    "V*", "EB*", "SB*", "PulsV*", "RotV*", "RRLyr", "DCEP", "DCEPS",
    "RV", "Mira", "DSCT", "BCEP", "GDOR", "SX*", "BY*", "RS*", "ACV", "CV*", "LP*"
]

CATEGORY_TERMS = {
    "Wolf–Rayet variables": r"\bwolf[-–]rayet variable[s]?\b|\bwolf[-–]rayet\b",
    "Evolved stars": r"\bevolved star[s]?\b",
    "Variable / transiting exoplanets": r"\btransiting exoplanet[s]?\b|\bvariable exoplanet[s]?\b",
    "Cataclysmic variables": r"\bcataclysmic variable[s]?\b|\bcv[s]?\b(?!\*)",
    "T Tauri and flare stars": r"\bt tauri\b|\bflare star[s]?\b",
    "Rotation and pulsation": r"\brotation and pulsation\b",
    "OB-type pulsators": r"\bob[- ]type pulsator[s]?\b",
    "RR Lyrae": r"\brr lyrae\b",
    "Cepheids": r"\bcepheid[s]?\b",
    "RV Tauri": r"\brv tauri\b",
    "Mira variables": r"\bmira variable[s]?\b|\bmira\b",
    "Delta Scuti": r"\bdelta scuti\b|\bδ scuti\b",
    "SX Phoenicis": r"\bsx phoenicis\b|\bsx phe\b",
    "Beta Cephei": r"\bbeta cephei\b|\bβ cephei\b",
    "Gamma Doradus": r"\bgamma doradus\b|\bγ doradus\b|\bgdor\b",
    "Eclipsing binaries": r"\beclipsing binar(y|ies)\b|\beb\b",
}

FREE_PATTERNS_1 = {k: re.compile(v, flags=re.IGNORECASE) for k, v in FREE_TERMS.items()}
FREE_PATTERNS_3 = {k: re.compile(v, flags=re.IGNORECASE) for k, v in CATEGORY_TERMS.items()}

def extract_any_match(html: str) -> bool:
    soup = BeautifulSoup(html, "html.parser")
    raw_text = soup.get_text(" ", strip=True)
    norm_text = _normalize_text(raw_text)
    if find_terms_free(norm_text, FREE_PATTERNS_1):
        return True
    if find_codes(raw_text, SIMBAD_CODES):
        return True
    if find_terms_free(norm_text, FREE_PATTERNS_3):
        return True
    return False

def main():
    df_in = pd.read_excel(INPUT_XLSX, header=None, usecols=[0])
    df_in.columns = ["TIC_digit_only"]
    rows = []
    for _, row in df_in.iterrows():
        tic_raw = str(row["TIC_digit_only"]).strip()
        m = re.search(r"\d+", tic_raw)
        if not m:
            continue
        tic_num = m.group(0)
        url_obj = build_obj_url(tic_num)
        url_bib = build_bib_url(tic_num)
        ref_text = None
        flag = "Ні"
        try:
            r = requests.get(url_bib, headers=HEADERS, timeout=25)
            r.raise_for_status()
            html = r.text
            ref_text = fetch_references_text(html)
            if extract_any_match(html):
                flag = "Так"
        except Exception:
            try:
                r = requests.get(url_obj, headers=HEADERS, timeout=20)
                r.raise_for_status()
                html = r.text
                ref_text = fetch_references_text(html)
                if extract_any_match(html):
                    flag = "Так"
            except Exception:
                pass
        if ref_text and not re.search(r"Total\s*0\)", ref_text):
            ref_link = ref_text
        else:
            ref_link = ""
        rows.append({
            "A_TIC": f"TIC {tic_num}",
            "B_References_link_text": ref_link,
            "B_References_url": url_bib,
            "C_Так_чи_Ні": flag,
        })
        time.sleep(0.4)
    out = pd.DataFrame(rows)
    with pd.ExcelWriter(OUTPUT_XLSX, engine="xlsxwriter") as writer:
        out.to_excel(writer, index=False, sheet_name="Sheet1", startrow=0)
        ws = writer.sheets["Sheet1"]
        col_txt = out.columns.get_loc("B_References_link_text")
        col_url = out.columns.get_loc("B_References_url")
        for r in range(len(out)):
            link_text = out.iat[r, col_txt]
            link_url = out.iat[r, col_url]
            if link_text:
                ws.write_url(r + 1, col_txt, link_url, string=link_text)
        ws.set_column(col_url, col_url, None, None, {'hidden': True})

if __name__ == "__main__":
    main()
