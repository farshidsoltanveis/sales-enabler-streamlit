# ======================================
# Purolator (FR) Invoices → Excel (Jupyter one-cell, robust weights & pieces)
# ======================================

# --- 0) Dependencies: install if missing (works in Jupyter or plain Python) ---
import sys, subprocess
def _ensure(pkg):
    try:
        __import__(pkg)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg])

for _p in ("pymupdf", "pandas", "openpyxl"):
    _ensure(_p)

# --- 1) Imports & Globals ---
import re
from datetime import datetime
from pathlib import Path
from typing import List, Tuple, Dict, Any, Optional

import fitz  # PyMuPDF
import pandas as pd

# ====== Normalization & Regex ======
def _norm(s: str) -> str:
    # remove soft hyphen, normalize NBSP & narrow NBSP to space, trim
    return s.replace("\u00ad", "").replace("\u00a0", " ").replace("\u202f", " ").strip()

DATE_YMD_RE = re.compile(r"^\s*(\d{4})/(\d{2})/(\d{2})\s*$")          # e.g., 2025/09/20
TRACKING_RE = re.compile(r"^\s*(\d{9,})\s*$")                         # numeric Purolator N° d’envoi
MONEY_RE    = re.compile(r"(-?\d{1,3}(?:[ .]\d{3})*,\d{2})\s*\$")     # '1 229,67 $' or '7,38 $'
UNIT_LB     = r"(?:LB|LBS|lb|lbs)\b"                                  # accept singular/plural/case

# Labelled patterns (tolerate line breaks/hyphens by searching on merged text)
_NUM = r"(\d+(?:[.,]\d+)?)"
RE_PIECES_LABEL   = re.compile(r"(?:Nbre|Nb|Nombre)\s*de\s*pi[eè]ces\s*:?\s*(\d+)", re.IGNORECASE)
RE_POIDS_FACTURE  = re.compile(r"Poids\s*factur[eé]\s*:?\s*" + _NUM + r"\s*" + UNIT_LB, re.IGNORECASE)
RE_POIDS_DECLARE  = re.compile(r"Poids\s*d[eé]clar[eé]\s*:?\s*" + _NUM + r"\s*" + UNIT_LB, re.IGNORECASE)

# Compact & very-permissive fallback:
#   "1 23 LB Poids déclaré 21 LB" or "2 57 LB"
RE_COMPACT_ROW    = re.compile(
    r"\b(?P<pieces>\d+)\s+(?P<billed>\d+(?:[.,]\d+)?)\s*" + UNIT_LB +
    r"(?:.*?Poids\s*d[eé]clar[eé]\s*:?\s*(?P<declared>\d+(?:[.,]\d+)?)\s*" + UNIT_LB + r")?",
    re.IGNORECASE
)

def _to_float_fr(x: str) -> float:
    return float(x.replace(",", "."))

def _clean_amount(s: str) -> Optional[float]:
    m = MONEY_RE.search(s)
    if not m: return None
    txt = m.group(1).replace(" ", "").replace(".", "").replace(",", ".")
    try: return float(txt)
    except: return None

def _extract_all_amounts(s: str) -> List[float]:
    vals = []
    for m in MONEY_RE.finditer(s):
        txt = m.group(1).replace(" ", "").replace(".", "").replace(",", ".")
        try: vals.append(float(txt))
        except: pass
    return vals

def read_pdf_lines(pdf_path: str) -> List[str]:
    doc = fitz.open(pdf_path)
    text = "\n".join([p.get_text("text") for p in doc])
    doc.close()
    return [_norm(ln) for ln in text.splitlines()]

# ====== Header fields ======
def parse_invoice_number(lines: List[str]) -> Optional[str]:
    for ln in lines:
        if "Numéro de facture" in ln:
            m = re.search(r"Numéro de facture\s*:\s*([A-Za-z0-9\-]+)", ln)
            if m: return m.group(1).strip()
    return None

def parse_invoice_date(lines: List[str]) -> Optional[str]:
    for ln in lines:
        if "Date de facture" in ln:
            m = re.search(r"Date de facture\s*:\s*(\d{4}/\d{2}/\d{2})", ln)
            if m: return m.group(1)
    return None

# ====== Robust extractor for Pieces/Weights over a whole block ======
def _extract_pieces_weights_from_block(block_lines: List[str]) -> Tuple[Optional[int], Optional[float], Optional[float]]:
    """
    Given all lines of a shipment block (between date and next date/summary),
    return (pieces, billed_lb, declared_lb).
    Strategy:
      1) Try labelled forms on the merged text (handles line breaks).
      2) Try compact rows on each single line and on 2-line merges.
    """
    pieces = billed = declared = None
    merged = " ".join([_norm(x) for x in block_lines if _norm(x)])

    # 1) Labelled across the block
    m = RE_PIECES_LABEL.search(merged)
    if m: pieces = int(m.group(1))
    m = RE_POIDS_FACTURE.search(merged)
    if m: billed = _to_float_fr(m.group(1))
    m = RE_POIDS_DECLARE.search(merged)
    if m: declared = _to_float_fr(m.group(1))

    # 2) Compact on each line and two-line merges
    if pieces is None or billed is None:
        n = len(block_lines)
        for k in range(n):
            l1 = _norm(block_lines[k])
            if l1:
                mc = RE_COMPACT_ROW.search(l1)
                if mc:
                    if pieces is None: pieces = int(mc.group("pieces"))
                    if billed  is None: billed  = _to_float_fr(mc.group("billed"))
                    if declared is None and mc.group("declared"):
                        declared = _to_float_fr(mc.group("declared"))
            if (pieces is None or billed is None) and k+1 < n:
                l2 = (l1 + " " + _norm(block_lines[k+1])).strip()
                mc2 = RE_COMPACT_ROW.search(l2)
                if mc2:
                    if pieces is None: pieces = int(mc2.group("pieces"))
                    if billed  is None: billed  = _to_float_fr(mc2.group("billed"))
                    if declared is None and mc2.group("declared"):
                        declared = _to_float_fr(mc2.group("declared"))
            if pieces is not None and billed is not None:
                break

    return pieces, billed, declared

# ====== Shipments (Envois du compte …) ======
def parse_purolator_shipments(lines: List[str], invoice_file: str, invoice_number: Optional[str]) -> pd.DataFrame:
    records: List[Dict[str, Any]] = []
    in_ship_section = False
    i = 0
    while i < len(lines):
        ln = lines[i]
        if ("Envois du compte" in ln) and ("Date de facture" not in ln):
            in_ship_section = True
            i += 1
            continue
        if not in_ship_section:
            i += 1
            continue

        m_date = DATE_YMD_RE.match(ln)
        if m_date:
            date_iso = f"{m_date.group(1)}-{m_date.group(2)}-{m_date.group(3)}"
            # tracking on next line (numeric)
            tracking = None
            if i + 1 < len(lines):
                m_tr = TRACKING_RE.match(lines[i+1])
                if m_tr: tracking = m_tr.group(1)

            # capture block until next date or a partial total line
            j = i + 2
            block_lines: List[str] = []
            while j < len(lines) and not DATE_YMD_RE.match(lines[j]) and "Total partiel" not in lines[j]:
                block_lines.append(lines[j])
                j += 1

            # 1) robust pieces/weights
            pieces, billed_lb, declared_lb = _extract_pieces_weights_from_block(block_lines)

            # 2) service/amounts/taxes from merged row
            row = " ".join(block_lines)
            service_name = None
            if "Purolator Express" in row:
                service_name = "Purolator Express"
            elif "Purolator Routier" in row:
                service_name = "Purolator Routier"
            else:
                m_srv = re.search(r"Description du service\s*:\s*([A-Za-zÀ-ÿ' \-]+)", row)
                if m_srv: service_name = m_srv.group(1).strip()

            amts = _extract_all_amounts(row)
            service_charge = None
            line_total = None
            if amts:
                line_total = amts[-1]
                if len(amts) >= 2:
                    service_charge = amts[0]

            fuel = tps = tvq = tvh = 0.0
            m = re.search(r"Supplément de carburant\s+([0-9 .,\-]+\$)", row)
            if m:
                v = _clean_amount(m.group(1)); fuel = v if v is not None else 0.0
            m = re.search(r"TPS\s+([0-9 .,\-]+\$)", row)
            if m:
                v = _clean_amount(m.group(1)); tps = v if v is not None else 0.0
            m = re.search(r"TVQ\s+([0-9 .,\-]+\$)", row)
            if m:
                v = _clean_amount(m.group(1)); tvq = v if v is not None else 0.0
            m = re.search(r"TVH\s+[A-Z]{2}\s+([0-9 .,\-]+\$)", row) or re.search(r"TVH\s+([0-9 .,\-]+\$)", row)
            if m:
                v = _clean_amount(m.group(1)); tvh = v if v is not None else 0.0

            records.append({
                "Carrier": "Purolator",
                "Invoice File": invoice_file,
                "Invoice Number": invoice_number,
                "Date": date_iso,
                "Tracking Number": tracking,
                "Service": service_name,
                "Pieces": pieces,
                "Billed Weight (lb)": billed_lb,
                "Declared Weight (lb)": declared_lb,
                "Service Charge (CAD)": service_charge,
                "Fuel Surcharge (CAD)": fuel,
                "TPS (CAD)": tps,
                "TVQ (CAD)": tvq,
                "TVH (CAD)": tvh,
                "Line Total (CAD)": line_total
            })

            i = j
            continue

        if "Total partiel du compte de facturation" in ln or "TotalPoids total" in ln:
            in_ship_section = False
        i += 1

    return pd.DataFrame(records)

# ====== Autres services → Charge Corrections ======
def parse_purolator_other_services(lines: List[str], invoice_file: str, invoice_number: Optional[str]) -> pd.DataFrame:
    records: List[Dict[str, Any]] = []
    in_other = False
    i = 0
    while i < len(lines):
        ln = lines[i]
        if "Autres services" in ln and "Date d'expédition" in ln:
            in_other = True
            i += 1
            continue

        if in_other:
            if DATE_YMD_RE.match(ln):
                date_iso = ln.replace("/", "-")
                tracking = None
                desc = None
                billed = None
                tps = tvq = tvh = 0.0
                reason = None

                if i + 1 < len(lines):
                    row = " ".join(lines[i+1:i+6])
                    m_tr = re.search(r"(\d{9,})(?:AC)?", row)
                    if m_tr: tracking = m_tr.group(1)
                    m_desc = re.search(r"(Adresse corrig[eé]e|Adresse déclarée|.*?correction.*?|.*?service.*?)", row, re.IGNORECASE)
                    if m_desc: desc = m_desc.group(1)
                    amts = _extract_all_amounts(row)
                    if amts: billed = amts[0]
                    m = re.search(r"TPS\s+([0-9 .,\-]+\$)", row)
                    if m: tps = _clean_amount(m.group(1)) or 0.0
                    m = re.search(r"TVQ\s+([0-9 .,\-]+\$)", row)
                    if m: tvq = _clean_amount(m.group(1)) or 0.0
                    m = re.search(r"TVH\s+([0-9 .,\-]+\$)", row)
                    if m: tvh = _clean_amount(m.group(1)) or 0.0
                    m_r = re.search(r"Raison\s*:\s*([A-Za-z0-9 \-_/]+)", row, re.IGNORECASE)
                    if m_r: reason = m_r.group(1).strip()

                records.append({
                    "Carrier": "Purolator",
                    "Invoice File": invoice_file,
                    "Invoice Number": invoice_number,
                    "Date": date_iso,
                    "Tracking Number": tracking,
                    "Description": desc or "Autres services",
                    "Billed Amount (CAD)": billed,
                    "TPS (CAD)": tps,
                    "TVQ (CAD)": tvq,
                    "TVH (CAD)": tvh,
                    "Reason": reason
                })

            if "Total partiel du compte de facturation" in ln or "TotalPoids" in ln:
                in_other = False
        i += 1

    return pd.DataFrame(records)

# ====== Driver ======
def process_purolator_invoices(pdf_paths: List[str]) -> Tuple[pd.DataFrame, pd.DataFrame]:
    all_out, all_corr = [], []
    for pdf in pdf_paths:
        path = str(Path(pdf))
        lines = read_pdf_lines(path)
        inv_number = parse_invoice_number(lines)

        df_ship = parse_purolator_shipments(lines, invoice_file=Path(pdf).name, invoice_number=inv_number)
        df_other = parse_purolator_other_services(lines, invoice_file=Path(pdf).name, invoice_number=inv_number)

        all_out.append(df_ship)
        all_corr.append(df_other)
        print(f"✓ {Path(pdf).name} → Outbound:{len(df_ship)}  Other services:{len(df_other)}")

    df_out = pd.concat([d for d in all_out if d is not None], ignore_index=True) if any(len(d)>0 for d in all_out) else pd.DataFrame()
    df_cor = pd.concat([d for d in all_corr if d is not None], ignore_index=True) if any(len(d)>0 for d in all_corr) else pd.DataFrame()
    return df_out, df_cor


# --- Integrated workbook helpers (Purolator) ---
import os
from pathlib import Path
from datetime import datetime
import pandas as pd

def _safe_concat(existing: pd.DataFrame, new: pd.DataFrame) -> pd.DataFrame:
    """Column-union concat preserving the order of the first dataset."""
    if existing is None or existing.empty:
        return new.copy()
    if new is None or new.empty:
        return existing.copy()

    # Keep the existing column order; append any truly new columns at the end
    all_cols = list(existing.columns) + [c for c in new.columns if c not in existing.columns]
    return pd.concat(
        [existing.reindex(columns=all_cols), new.reindex(columns=all_cols)],
        ignore_index=True
    )


def _dedup(df: pd.DataFrame, keys: list[str]) -> pd.DataFrame:
    if df is None or df.empty: return df
    present_keys = [k for k in keys if k in df.columns]
    if not present_keys: return df
    return df.drop_duplicates(subset=present_keys, keep="last").reset_index(drop=True)

def _dup_key_cols_for(sheet_name: str) -> list[str]:
    # Adjust if you ever change schemas
    if sheet_name == "Outbound_Shipments":
        return ["Invoice Number", "Tracking Number", "Date"]
    if sheet_name == "Charge_Corrections":
        return ["Invoice Number", "Tracking Number", "Date"]
    return []

def _make_dup_key(df: pd.DataFrame, cols: list[str]) -> pd.Series:
    """
    Return a Series of composite keys like 'INV|TRACK|DATE' for each row.
    Handles missing columns by substituting empty strings.
    """
    if df is None or df.empty or not cols:
        return pd.Series([], dtype=object)

    parts = []
    for c in cols:
        if c in df.columns:
            s = df[c].fillna("").astype(str).str.strip()
        else:
            s = pd.Series([""] * len(df), index=df.index)
        parts.append(s)

    # Join elementwise: 'a'|'b'|'c'
    key = parts[0]
    for s in parts[1:]:
        key = key + "|" + s
    return key

def _normalize_key_columns(df: pd.DataFrame, key_cols: list[str]) -> pd.DataFrame:
    """Make key columns comparable across runs by normalizing types & formats."""
    if df is None or df.empty or not key_cols:
        return df
    df = df.copy()

    if "Tracking Number" in key_cols and "Tracking Number" in df.columns:
        s = df["Tracking Number"].astype(str).str.strip()
        # drop any trailing ".0" introduced by numeric coercion
        s = s.str.replace(r"\.0$", "", regex=True)
        # keep only digits (robust against spaces, commas, etc.)
        s = s.str.replace(r"\D", "", regex=True)
        df["Tracking Number"] = s

    if "Invoice Number" in key_cols and "Invoice Number" in df.columns:
        s = df["Invoice Number"].astype(str).str.strip()
        # choose your rule; uppercase helps stability
        df["Invoice Number"] = s.str.upper()

    if "Date" in key_cols and "Date" in df.columns:
        # force YYYY-MM-DD
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.strftime("%Y-%m-%d")

    return df


def update_purolator_integrated_workbook(
    df_out: pd.DataFrame,
    df_cor: pd.DataFrame,
    integrated_path: str = "data/Purolator_integrated.xlsx",
) -> str:
    """
    Append outbound & corrections into a persistent Excel, pre-filtering duplicates
    using a composite key (Invoice Number, Tracking Number, Date). Preserves
    first-seen column order and appends any new columns to the end.
    """
    integrated_path = str(Path(integrated_path))
    Path(integrated_path).parent.mkdir(parents=True, exist_ok=True)

    imported_at = datetime.now().isoformat(timespec="seconds")
    def _tag(df):
        if isinstance(df, pd.DataFrame) and not df.empty:
            df = df.copy()
            df["Imported At"] = imported_at
        return df

    df_out = _tag(df_out)
    df_cor = _tag(df_cor)

    # Read existing (if any)
    existing = {}
    if os.path.exists(integrated_path):
        try:
            existing = pd.read_excel(
                integrated_path, sheet_name=None, engine="openpyxl", dtype=str
            )
        except Exception:
            existing = {}


    # Build a per-sheet plan
    sheets_plan = [
        ("Outbound_Shipments", df_out),
        ("Charge_Corrections", df_cor),
    ]

    # Pre-filter: remove rows whose composite key already exists in the target sheet
    filtered = {}
    for sname, df_new in sheets_plan:
        if df_new is None or df_new.empty:
            filtered[sname] = df_new
            continue

        key_cols = _dup_key_cols_for(sname)

        # normalize incoming rows
        df_new_norm = _normalize_key_columns(df_new, key_cols)

        # normalize existing key columns (only the columns we need)
        if sname in existing and isinstance(existing[sname], pd.DataFrame) and not existing[sname].empty:
            exist_df = existing[sname]
            exist_df = _normalize_key_columns(exist_df, key_cols)
            exist_keys = set(_make_dup_key(exist_df[key_cols], key_cols))
        else:
            exist_keys = set()

        new_keys = _make_dup_key(df_new_norm[key_cols], key_cols)
        keep_mask = ~new_keys.isin(exist_keys)
        filtered[sname] = df_new_norm.loc[keep_mask].copy()


    # Now merge with order-preserving column union and write
    with pd.ExcelWriter(integrated_path, engine="openpyxl", mode="w") as writer:
        for sname, df_new in sheets_plan:
            df_exist = existing.get(sname, pd.DataFrame())
            df_new_kept = filtered[sname]

            # Nothing to write at all?
            if (df_exist is None or df_exist.empty) and (df_new_kept is None or df_new_kept.empty):
                continue

            merged = _safe_concat(df_exist, df_new_kept)  # preserves existing order; new cols at end

            # (Optional extra safety) de-dup again, cheap with set-sized subset
            key_cols = _dup_key_cols_for(sname)
            merged = _dedup(merged, key_cols)

            if not merged.empty:
                merged.to_excel(writer, sheet_name=sname, index=False)

    return integrated_path





def run_purolator_parser(input_pdfs, output_dir="output", output_basename="Purolator_invoices",
                         single_sheet=False,
                         update_integrated=False,
                         integrated_path="data/Purolator_integrated.xlsx"):

    """
    input_pdfs: list[str] of file paths (absolute or relative like 'input/Purolator-Sept.pdf')
    output_dir: folder to save Excel(s)
    output_basename: file name stem for outputs
    single_sheet: if True, also produce a combined single-sheet file
    Returns: str path to the main Excel file
    """
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    output_excel = output_dir / f"{output_basename}.xlsx"
    one_sheet_excel = output_dir / f"{output_basename}_single.xlsx"

    # Use your existing function here (change the name if different):
    df_out, df_cor = process_purolator_invoices(input_pdfs)

    with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
        if not df_out.empty:
            df_out.to_excel(writer, sheet_name="Outbound_Shipments", index=False)
        if not df_cor.empty:
            df_cor.to_excel(writer, sheet_name="Charge_Corrections", index=False)

    if single_sheet:
        frames = []
        if not df_out.empty:
            frames.append(df_out.assign(Section="Outbound"))
        if not df_cor.empty:
            frames.append(df_cor.assign(Section="Charge_Corrections"))
        if frames:
            pd.concat(frames, ignore_index=True).to_excel(one_sheet_excel, index=False)

       # (NEW) Update integrated workbook
    if update_integrated:
        integrated_written = update_purolator_integrated_workbook(
            df_out, df_cor, integrated_path=integrated_path
        )
        print(f"✅ Integrated workbook updated: {Path(integrated_written).resolve()}")


    return str(output_excel)

if __name__ == "__main__":
    INPUT_PDFS = [
        "input/Purolator-Sept-2025-Louve.pdf",
    ]
    run_purolator_parser(INPUT_PDFS, output_dir="output", output_basename="Purolator_invoices")


