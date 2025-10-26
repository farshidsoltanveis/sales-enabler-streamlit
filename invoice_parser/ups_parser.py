# ======================================
# UPS Invoices → Excel (Outbound filtered + surcharge Published/Incentive/Billed)
# Jupyter one-cell version
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
from typing import List, Dict, Any, Tuple, Optional

import fitz  # PyMuPDF
import pandas as pd
from datetime import datetime
import os

# Regex helpers
FLOAT_RE = re.compile(r"^-?\d+\.\d{2}$")
CAN_POSTAL_RE = re.compile(r"[A-Z]\d[A-Z]\s?\d[A-Z]\d")
DATE_MD_RE = re.compile(r"^\d{2}/\d{2}$")


# --- 2) Low-level helpers ---
def read_pdf_lines(pdf_path: str) -> List[str]:
    """Extract text lines from the PDF (page order), normalized for parsing."""
    doc = fitz.open(pdf_path)
    text = "\n".join([p.get_text("text") for p in doc])
    doc.close()
    return [ln.strip() for ln in text.splitlines()]


def normalize_postal(s: Optional[str]) -> Optional[str]:
    return s.replace(" ", "") if s else s


def parse_invoice_year(lines: List[str], fallback_year: int = None) -> int:
    """
    Parse 'Invoice Date August 16, 2025' → 2025.
    Fallback to current year when not found.
    """
    joined = "\n".join(lines)
    m = re.search(r"Invoice Date\s+([A-Za-z]+\s+\d{1,2},\s+(\d{4}))", joined)
    if m:
        try:
            dt = datetime.strptime(m.group(1), "%B %d, %Y")
            return dt.year
        except Exception:
            pass
    return fallback_year or datetime.now().year


def parse_invoice_number(lines: List[str]) -> Optional[str]:
    """Capture 'Invoice Number XXXXX...' from the header for traceability."""
    joined = "\n".join(lines)
    m = re.search(r"Invoice Number\s+([A-Za-z0-9]+)", joined)
    return m.group(1) if m else None


# --- 3) Parsers for each section ---
def parse_outbound(lines: List[str], invoice_year: int, invoice_file: str, invoice_number: Optional[str]) -> List[Dict[str, Any]]:
    """
    Parse 'Outbound Shipping API' shipments (one record per package).
    Robust to line breaks where Postal and Zone can be on the same line or split.
    Parses base triplet (Published/Incentive/Billed), surcharges triplets (Published/Incentive/Billed)
    and the per-shipment 'Total' billed. DAS and DAS-Extended are summed together.
    """
    records = []
    in_outbound = False
    current_date_iso = None
    i = 0

    while i < len(lines):
        ln = lines[i]

        # Enter/exit outbound blocks (can appear across pages)
        if "Outbound" in ln and "Shipping API" in " ".join(lines[i:i+3]):
            in_outbound = True
            i += 1
            continue
        if "Total for Internet-ID" in ln:
            in_outbound = False

        if not in_outbound:
            i += 1
            continue

        # Date lines (MM/DD) — UPS omits year in line items
        if DATE_MD_RE.match(ln):
            try:
                mm, dd = map(int, ln.split("/"))
                current_date_iso = datetime(invoice_year, mm, dd).date().isoformat()
            except Exception:
                current_date_iso = None
            i += 1
            continue

        # Shipment header starts with Tracking (1Z...)
        if ln.startswith("1Z") and len(ln) > 10:
            tracking = ln
            service = lines[i+1] if i+1 < len(lines) else ""
            postal_zone_line = lines[i+2] if i+2 < len(lines) else ""

            # Postal & Zone may share line or be split
            m_post = CAN_POSTAL_RE.search(postal_zone_line)
            if not m_post:
                i += 1
                continue

            recv_postal = normalize_postal(m_post.group(0))
            consumed = 1
            m_zone_same = re.search(r"\b(\d{3})\b", postal_zone_line[m_post.end():])
            if m_zone_same:
                zone_val = int(m_zone_same.group(1))
            else:
                # Zone likely in the next line
                try:
                    zone_val = int(lines[i+3])
                    consumed = 2
                except Exception:
                    i += 1
                    continue

            # Find first 'X lbs' (Weight), then next 3 floats are Published/Incentive/Billed
            j = i + 2 + consumed
            weight_val = None
            published = incentive = billed = None

            while j < len(lines):
                # Stop scanning when we hit next block headers
                if lines[j].startswith("1Z") or DATE_MD_RE.match(lines[j]) or \
                   "Total for Internet-ID" in lines[j] or \
                   ("Outbound" in lines[j] and "Shipping API" in " ".join(lines[j:j+3])):
                    break

                m_w = re.search(r"(\d+(?:\.\d+)?)\s*lbs", lines[j])
                if m_w and weight_val is None:
                    weight_val = float(m_w.group(1))
                    # Collect the next 3 floats
                    k = j + 1
                    floats = []
                    while k < len(lines) and len(floats) < 3:
                        if FLOAT_RE.match(lines[k]):
                            floats.append(float(lines[k]))
                        k += 1
                    if len(floats) == 3:
                        published, incentive, billed = floats
                        j = k
                        break
                j += 1

            if weight_val is None or published is None or incentive is None or billed is None:
                i += 1
                continue

            rec = {
                "Invoice File": invoice_file,
                "Invoice Number": invoice_number,
                "Date": current_date_iso,
                "Tracking Number": tracking,
                "Service": service,
                "Sender Postal Code": None,
                "Receiver Postal Code": recv_postal,
                "Sender Name": None,
                "Zone": zone_val,
                # Optional Customer Weight
                "Customer Weight (lb)": None,
                "Standard Weight (lb)": weight_val,

                # Base triplet (already captured above)
                "Published Charge (CAD)": published,
                "Incentive Credit (CAD)": incentive,
                "Billed Charge (CAD)": billed,

                # Surcharge triplets (initialize to 0.0 and sum if multiple lines appear)
                "Residential Surcharge Published (CAD)": 0.0,
                "Residential Surcharge Incentive (CAD)": 0.0,
                "Residential Surcharge Billed (CAD)": 0.0,

                "Delivery Area Surcharge Published (CAD)": 0.0,
                "Delivery Area Surcharge Incentive (CAD)": 0.0,
                "Delivery Area Surcharge Billed (CAD)": 0.0,

                "Fuel Surcharge Published (CAD)": 0.0,
                "Fuel Surcharge Incentive (CAD)": 0.0,
                "Fuel Surcharge Billed (CAD)": 0.0,

                # Per-shipment total billed (if present)
                "Total (CAD)": None,

                
            }

            # Walk details until new shipment/date/section boundary
            i = j
            while i < len(lines):
                l2 = lines[i]
                if l2.startswith("1Z") or DATE_MD_RE.match(l2) or \
                   "Total for Internet-ID" in l2 or \
                   ("Outbound" in l2 and "Shipping API" in " ".join(lines[i:i+3])):
                    break

                # Customer Weight can be two lines
                if l2.startswith("Customer Weight"):
                    if i + 1 < len(lines):
                        mw = re.search(r"(\d+(?:\.\d+)?)\s*lbs", lines[i+1])
                        if mw:
                            rec["Customer Weight (lb)"] = float(mw.group(1))
                    i += 2
                    continue

                # Helper to read a numeric triplet (Published, Incentive, Billed)
                def read_triplet(idx: int):
                    vals, k = [], idx
                    while k < len(lines) and len(vals) < 3:
                        if FLOAT_RE.match(lines[k]):
                            vals.append(float(lines[k]))
                        k += 1
                    return vals, k  # vals=[pub, inc, bill]

                # Residential Surcharge
                if l2.startswith("Residential Surcharge"):
                    vals, k = read_triplet(i + 1)
                    if len(vals) == 3:
                        rec["Residential Surcharge Published (CAD)"] += vals[0]
                        rec["Residential Surcharge Incentive (CAD)"] += vals[1]
                        rec["Residential Surcharge Billed (CAD)"] += vals[2]
                    i = k
                    continue

                # Delivery Area Surcharge (regular or Extended) – aggregate together
                if l2.startswith("Delivery Area Surcharge"):
                    vals, k = read_triplet(i + 1)
                    if len(vals) == 3:
                        rec["Delivery Area Surcharge Published (CAD)"] += vals[0]
                        rec["Delivery Area Surcharge Incentive (CAD)"] += vals[1]
                        rec["Delivery Area Surcharge Billed (CAD)"] += vals[2]
                    i = k
                    continue

                # Fuel Surcharge
                if l2.startswith("Fuel Surcharge"):
                    vals, k = read_triplet(i + 1)
                    if len(vals) == 3:
                        rec["Fuel Surcharge Published (CAD)"] += vals[0]
                        rec["Fuel Surcharge Incentive (CAD)"] += vals[1]
                        rec["Fuel Surcharge Billed (CAD)"] += vals[2]
                    i = k
                    continue

                # Per-shipment Total
                if l2.strip().startswith("Total"):
                    vals, k = read_triplet(i + 1)
                    if len(vals) == 3:
                        rec["Total (CAD)"] = vals[2]
                    i = k
                    continue

                # Sender block
                if l2.startswith("Sender"):
                    s_lines = []
                    i += 1
                    while i < len(lines) and not lines[i].startswith("Receiver") and not lines[i].startswith("UserID"):
                        s_lines.append(lines[i])
                        i += 1
                    sender_block = " ".join(s_lines)
                    msp = CAN_POSTAL_RE.search(sender_block)
                    if msp:
                        rec["Sender Postal Code"] = normalize_postal(msp.group(0))
                    # Sender name until we encounter address number
                    name_parts = []
                    for s in s_lines:
                        if re.match(r"^\d", s):
                            break
                        if s.strip():
                            name_parts.append(s.strip())
                    if name_parts:
                        np = [p.replace("Sender  :", "").replace("Sender :", "").strip(" :") for p in name_parts]
                        rec["Sender Name"] = " ".join(np)
                    continue

                # Receiver block → ensure Receiver postal
                if l2.startswith("Receiver"):
                    r_lines = []
                    i += 1
                    while i < len(lines) and not lines[i].startswith("Message Codes") and \
                          not lines[i].startswith("1Z") and not DATE_MD_RE.match(lines[i]):
                        r_lines.append(lines[i])
                        i += 1
                    r_block = " ".join(r_lines)
                    mrp = CAN_POSTAL_RE.search(r_block)
                    if mrp:
                        rec["Receiver Postal Code"] = normalize_postal(mrp.group(0))
                    continue

                i += 1

            records.append(rec)
            continue

        i += 1

    # Fill missing sender fields with the mode (optional, completeness)
    if records:
        df_tmp = pd.DataFrame(records)
        for col in ["Sender Postal Code", "Sender Name"]:
            mode = df_tmp[col].dropna().mode()
            if not mode.empty:
                fill_val = mode.iloc[0]
                for r in records:
                    if not r.get(col):
                        r[col] = fill_val
    return records


def parse_returns(lines: List[str], invoice_year: int, invoice_file: str, invoice_number: Optional[str]) -> List[Dict[str, Any]]:
    """Parse the 'Inbound UPS Returns Transportation' block."""
    records = []
    in_returns = False
    i = 0

    while i < len(lines):
        ln = lines[i]
        if "Inbound UPS Returns Transportation" in " ".join(lines[i:i+4]):
            in_returns = True
            i += 1
            continue
        if in_returns and ln.strip().startswith("Total UPS Returns Transportation"):
            in_returns = False

        if not in_returns:
            i += 1
            continue

        if DATE_MD_RE.match(ln):
            try:
                date_iso = datetime(invoice_year, int(ln[:2]), int(ln[3:5])).date().isoformat()
            except Exception:
                date_iso = None
            tracking = lines[i+1] if i+1 < len(lines) else ""
            service = lines[i+2] if i+2 < len(lines) else ""
            zone_line = lines[i+3] if i+3 < len(lines) else ""
            try:
                zone = int(zone_line)
            except Exception:
                i += 1
                continue
            weight_line = lines[i+4] if i+4 < len(lines) else ""
            m_w = re.search(r"(\d+(?:\.\d+)?)\s*lbs", weight_line)
            if not m_w:
                i += 1
                continue

            j = i + 5
            pub = inc = bill = None
            while j < len(lines):
                if lines[j].strip().lower().startswith("total"):
                    if FLOAT_RE.match(lines[j+1]) and FLOAT_RE.match(lines[j+2]) and FLOAT_RE.match(lines[j+3]):
                        pub = float(lines[j+1]); inc = float(lines[j+2]); bill = float(lines[j+3])
                    break
                j += 1

            records.append({
                "Invoice File": invoice_file,
                "Invoice Number": invoice_number,
                "Returned Date": date_iso,
                "Tracking Number": tracking,
                "Service": service,
                "Zone": zone,
                "Weight (lb)": float(m_w.group(1)),
                "Published Charge (CAD)": pub,
                "Incentive Credit (CAD)": inc,
                "Billed Charge (CAD)": bill,
            })
            i = j + 4
            continue

        i += 1

    return records


def parse_residential_adjustments(lines: List[str], invoice_year: int, invoice_file: str, invoice_number: Optional[str]) -> List[Dict[str, Any]]:
    """Parse 'Adjustments & Other Charges' → 'Residential Adjustments' (Shipping API)."""
    records = []
    in_adj = False
    date_iso = None
    i = 0

    while i < len(lines):
        ln = lines[i]
        if "Adjustments & Other Charges" in ln and "Residential Adjustments" in " ".join(lines[i:i+6]):
            in_adj = True
            i += 1
            continue
        if in_adj and ln.startswith("Total Residential Adjustments"):
            in_adj = False

        if not in_adj:
            i += 1
            continue

        if DATE_MD_RE.match(ln):
            try:
                date_iso = datetime(invoice_year, int(ln[:2]), int(ln[3:5])).date().isoformat()
            except Exception:
                date_iso = None
            i += 1
            continue

        if ln.startswith("1Z") and len(ln) > 10:
            tracking = ln
            j = i + 1
            entry = {
                "Invoice File": invoice_file,
                "Invoice Number": invoice_number,
                "Date": date_iso,
                "Tracking Number": tracking,
                "Residential Surcharge Billed (CAD)": 0.0,
                "Fuel Surcharge Billed (CAD)": 0.0
            }
            while j < len(lines) and not lines[j].startswith("1Z") and not DATE_MD_RE.match(lines[j]) \
                  and not lines[j].startswith("Total"):
                if lines[j].startswith("Residential Surcharge"):
                    if FLOAT_RE.match(lines[j+1]) and FLOAT_RE.match(lines[j+2]) and FLOAT_RE.match(lines[j+3]):
                        entry["Residential Surcharge Billed (CAD)"] += float(lines[j+3])
                        j += 4
                        continue
                if lines[j].startswith("Fuel Surcharge"):
                    if FLOAT_RE.match(lines[j+1]) and FLOAT_RE.match(lines[j+2]) and FLOAT_RE.match(lines[j+3]):
                        entry["Fuel Surcharge Billed (CAD)"] += float(lines[j+3])
                        j += 4
                        continue
                j += 1
            records.append(entry)
            i = j
            continue

        i += 1

    return records


def parse_charge_corrections(lines: List[str], invoice_year: int, invoice_file: str, invoice_number: Optional[str]) -> List[Dict[str, Any]]:
    """Parse 'Shipping Charge Corrections' table; capture the trailing 'Adjustment Amount'."""
    records = []
    in_corr = False
    date_iso = None
    i = 0

    while i < len(lines):
        ln = lines[i]
        if "Shipping Charge Corrections" in ln and "avoid" in " ".join(lines[i:i+8]).lower():
            in_corr = True
            i += 1
            continue
        if in_corr and ln.startswith("Total Shipping Charge Corrections"):
            in_corr = False

        if not in_corr:
            i += 1
            continue

        if DATE_MD_RE.match(ln):
            try:
                date_iso = datetime(invoice_year, int(ln[:2]), int(ln[3:5])).date().isoformat()
            except Exception:
                date_iso = None
            i += 1
            continue

        if ln.startswith("1Z") and len(ln) > 10:
            tracking = ln
            j = i + 1
            adj_amount = None
            while j < len(lines) and not lines[j].startswith("Sender"):
                if FLOAT_RE.match(lines[j]):
                    adj_amount = float(lines[j])
                j += 1
            while j < len(lines) and not DATE_MD_RE.match(lines[j]) and not lines[j].startswith("1Z") \
                  and "Total Shipping Charge Corrections" not in lines[j]:
                j += 1

            records.append({
                "Invoice File": invoice_file,
                "Invoice Number": invoice_number,
                "Date": date_iso,
                "Tracking Number": tracking,
                "Adjustment Amount (CAD)": adj_amount
            })
            i = j
            continue

        i += 1

    return records


# --- 4) Driver that aggregates multiple PDFs ---
def process_invoices(pdf_paths: List[str]):
    all_outbound, all_returns, all_adj, all_corr = [], [], [], []

    for pdf in pdf_paths:
        pdf_path = str(Path(pdf))
        lines = read_pdf_lines(pdf_path)
        year = parse_invoice_year(lines)
        inv_number = parse_invoice_number(lines)

        outbound = parse_outbound(lines, year, invoice_file=Path(pdf).name, invoice_number=inv_number)
        returns = parse_returns(lines, year, invoice_file=Path(pdf).name, invoice_number=inv_number)
        adj = parse_residential_adjustments(lines, year, invoice_file=Path(pdf).name, invoice_number=inv_number)
        corr = parse_charge_corrections(lines, year, invoice_file=Path(pdf).name, invoice_number=inv_number)

        all_outbound.extend(outbound)
        all_returns.extend(returns)
        all_adj.extend(adj)
        all_corr.extend(corr)

        print(f"✓ {Path(pdf).name}  →  Outbound:{len(outbound)}  Returns:{len(returns)}  "
              f"ResAdj:{len(adj)}  Corrections:{len(corr)}")

    # Build DataFrames (aggregate across invoices)
    df_out = pd.DataFrame(all_outbound)
    df_ret = pd.DataFrame(all_returns)
    df_adj = pd.DataFrame(all_adj)
    df_cor = pd.DataFrame(all_corr)
    return df_out, df_ret, df_adj, df_cor




def _safe_concat(existing: pd.DataFrame, new: pd.DataFrame) -> pd.DataFrame:
    """Column-union concat to avoid key errors when columns differ slightly."""
    if existing is None or existing.empty:
        return new.copy()
    if new is None or new.empty:
        return existing.copy()
    # Align columns (union)
    all_cols = sorted(set(existing.columns) | set(new.columns))
    return pd.concat([existing.reindex(columns=all_cols),
                      new.reindex(columns=all_cols)], ignore_index=True)

def _dedup(df: pd.DataFrame, keys: list[str]) -> pd.DataFrame:
    """De-duplicate with a stable strategy: keep the last (newest) row."""
    if df is None or df.empty:
        return df
    present_keys = [k for k in keys if k in df.columns]
    if not present_keys:
        return df
    return df.drop_duplicates(subset=present_keys, keep="last").reset_index(drop=True)

def update_integrated_workbook(
    df_out: pd.DataFrame,
    df_ret: pd.DataFrame,
    df_adj: pd.DataFrame,
    df_cor: pd.DataFrame,
    integrated_path: str = "data/UPS_integrated.xlsx",
) -> str:
    """
    Merge the four section DataFrames into a persistent, de-duplicated Excel file.
    Creates the file if missing; otherwise appends + dedups per sheet.
    Returns the integrated_path.
    """
    integrated_path = str(Path(integrated_path))
    Path(integrated_path).parent.mkdir(parents=True, exist_ok=True)

    # Add an import timestamp (helps auditing/overrides)
    imported_at = datetime.now().isoformat(timespec="seconds")
    def _tag(df):
        if isinstance(df, pd.DataFrame) and not df.empty:
            df = df.copy()
            df["Imported At"] = imported_at
            return df
        return df

    df_out = _tag(df_out)
    df_ret = _tag(df_ret)
    df_adj = _tag(df_adj)
    df_cor = _tag(df_cor)

    # Read existing if present
    existing = {}
    if os.path.exists(integrated_path):
        try:
            existing = pd.read_excel(integrated_path, sheet_name=None, engine="openpyxl")
        except Exception:
            existing = {}

    # Merge per sheet
    sheets = {
        "Outbound_Shipments": df_out,                 # keys chosen to best avoid dupes
        "UPS_Returns": df_ret,
        "Residential_Adjustments": df_adj,
        "Charge_Corrections": df_cor,
    }
    # De-dup keys per sheet
    dedup_keys = {
        "Outbound_Shipments": ["Invoice Number", "Tracking Number", "Date"],  # your outbound has these cols :contentReference[oaicite:2]{index=2}
        "UPS_Returns": ["Invoice Number", "Tracking Number", "Returned Date"],# returns schema :contentReference[oaicite:3]{index=3}
        "Residential_Adjustments": ["Invoice Number", "Tracking Number", "Date"],  # adjustments schema :contentReference[oaicite:4]{index=4}
        "Charge_Corrections": ["Invoice Number", "Tracking Number", "Date"],  # corrections schema :contentReference[oaicite:5]{index=5}
    }

    with pd.ExcelWriter(integrated_path, engine="openpyxl", mode="w") as writer:
        for name, df_new in sheets.items():
            df_exist = existing.get(name, pd.DataFrame())
            merged = _safe_concat(df_exist, df_new)
            merged = _dedup(merged, dedup_keys.get(name, []))
            if not merged.empty:
                merged.to_excel(writer, sheet_name=name, index=False)

    return integrated_path



def run_ups_parser(
    input_pdfs,
    output_dir="output",
    output_basename="UPS_invoices",
    single_sheet=False,
    # NEW:
    update_integrated=False,
    integrated_path="data/UPS_integrated.xlsx",
):
    """
    Runs the UPS invoice parser using your existing logic (process_invoices) and
    writes Excel outputs to output_dir. Returns the main Excel path.

    Parameters
    ----------
    input_pdfs : list[str]
        Paths like ["input/UPS-aug-16.pdf", ...]
    output_dir : str
        Directory to write output Excel(s).
    output_basename : str
        Base filename (no extension).
    single_sheet : bool
        Also write a combined single-sheet Excel if True.
    update_integrated : bool
        If True, append results into a persistent, de-duplicated workbook.
    integrated_path : str
        Path to the integrated workbook (created if missing).
    """
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    output_excel = output_dir / f"{output_basename}.xlsx"
    single_excel = output_dir / f"{output_basename}_single.xlsx"

    # 1) Parse PDFs → 4 dataframes
    df_out, df_ret, df_adj, df_cor = process_invoices(input_pdfs)

    # 2) Build exclusion map → drop these trackings from outbound
    exclusion_map = {}  # tracking -> set(reasons)

    def _add_reason(df, tracking_col, reason):
        if isinstance(df, pd.DataFrame) and not df.empty and tracking_col in df.columns:
            for val in df[tracking_col].astype(str).str.strip():
                if not val:
                    continue
                exclusion_map.setdefault(val, set()).add(reason)

    _add_reason(df_adj, "Tracking Number", "Residential Adjustment")
    _add_reason(df_cor, "Tracking Number", "Charge Correction")
    exclude_trackings = set(exclusion_map.keys())

    # 3) Excluded rows (for reporting)
    key_cols = [
        "Invoice File", "Invoice Number", "Date", "Tracking Number",
        "Service", "Sender Postal Code", "Receiver Postal Code", "Sender Name",
        "Zone", "Standard Weight (lb)", "Customer Weight (lb)"
    ]
    if isinstance(df_out, pd.DataFrame) and not df_out.empty:
        present_cols = [c for c in key_cols if c in df_out.columns]
        excluded_rows = []
        for _, r in df_out.iterrows():
            t = str(r.get("Tracking Number", "")).strip()
            if t and t in exclude_trackings:
                row_dict = {k: r.get(k) for k in present_cols}
                row_dict["Reason"] = ", ".join(sorted(exclusion_map.get(t, [])))
                excluded_rows.append(row_dict)
        df_excluded = pd.DataFrame(excluded_rows)
    else:
        df_excluded = pd.DataFrame(columns=key_cols + ["Reason"])

    # 4) Filter outbound
    if isinstance(df_out, pd.DataFrame) and not df_out.empty:
        before = len(df_out)
        df_out_filtered = df_out[~df_out["Tracking Number"].astype(str).isin(exclude_trackings)].copy()
        after = len(df_out_filtered)
    else:
        df_out_filtered = df_out.copy() if isinstance(df_out, pd.DataFrame) else pd.DataFrame()
        before, after = 0, 0

    print(f"\n[UPS] Filtering outbound shipments:")
    print(f"  Outbound rows before: {before}")
    print(f"  Excluded trackings (unique): {len(exclude_trackings)}")
    print(f"  Outbound rows after:  {after}")

    # 5) Write multi-sheet Excel
    with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
        if not df_out_filtered.empty:
            df_out_filtered.to_excel(writer, sheet_name="Outbound_Shipments", index=False)
        if isinstance(df_ret, pd.DataFrame) and not df_ret.empty:
            df_ret.to_excel(writer, sheet_name="UPS_Returns", index=False)
        if isinstance(df_adj, pd.DataFrame) and not df_adj.empty:
            df_adj.to_excel(writer, sheet_name="Residential_Adjustments", index=False)
        if isinstance(df_cor, pd.DataFrame) and not df_cor.empty:
            df_cor.to_excel(writer, sheet_name="Charge_Corrections", index=False)
        if not df_excluded.empty:
            df_excluded.to_excel(writer, sheet_name="Excluded_Shipments", index=False)

    # 6) Optional combined single-sheet
    if single_sheet:
        frames = []
        if not df_out_filtered.empty:
            frames.append(df_out_filtered.assign(Section="Outbound (Filtered)"))
        if isinstance(df_ret, pd.DataFrame) and not df_ret.empty:
            frames.append(df_ret.assign(Section="UPS_Returns"))
        if isinstance(df_adj, pd.DataFrame) and not df_adj.empty:
            frames.append(df_adj.assign(Section="Residential_Adjustments"))
        if isinstance(df_cor, pd.DataFrame) and not df_cor.empty:
            frames.append(df_cor.assign(Section="Charge_Corrections"))
        if frames:
            pd.concat(frames, ignore_index=True).to_excel(single_excel, index=False)

    # 7) (NEW) Update the integrated workbook (append + dedup)
    if update_integrated:
        integrated_written = update_integrated_workbook(
            df_out_filtered, df_ret, df_adj, df_cor, integrated_path=integrated_path
        )
        print(f"✅ Integrated workbook updated: {Path(integrated_written).resolve()}")

    print(f"✅ Wrote: {output_excel.resolve()}")
    if single_sheet:
        print(f"✅ Wrote: {single_excel.resolve()}")

    return str(output_excel)



# Keep standalone behavior for quick tests (will NOT run on import)
if __name__ == "__main__":
    sample_files = [
        "input/UPS-aug-16.pdf",
        # "input/UPS-aug-09.pdf",
        # "input/UPS-aug-02.pdf",
    ]
    run_ups_parser(sample_files, output_dir="output", output_basename="UPS_invoices", single_sheet=False)
# -------------------------------------------------------------------------------