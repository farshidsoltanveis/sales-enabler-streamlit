# Sales Enabler/integrations/integrate_competitors_rate.py
from __future__ import annotations

import argparse
import os
import re
import sys
from typing import Optional, Tuple

import numpy as np
import pandas as pd


# ---------------------------------------------------------------------
# Import the analytics engine without requiring a package install.
# (Works even if "Sales Enabler" has a space in the folder name.)
# ---------------------------------------------------------------------
def _add_analytics_to_path():
    here = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.abspath(os.path.join(here, ".."))  # Sales Enabler/
    analytics_dir = os.path.join(project_root, "analytics_engine")
    if analytics_dir not in sys.path:
        sys.path.insert(0, analytics_dir)

_add_analytics_to_path()
try:
    from insight_generator import generate_insights
except Exception as e:
    raise RuntimeError("Failed to import insight_generator from analytics_engine") from e


# ---------------------------------------------------------------------
# Carrier detection & constants
# ---------------------------------------------------------------------
CARRIER_UPS = "UPS"
CARRIER_PURO = "Purolator"

def detect_uploaded_carrier(df: pd.DataFrame, file_path: str) -> Optional[str]:
    """
    Heuristic:
      1) Look at file name (UPS|Purolator).
      2) If unavailable, peek at 'Invoice File' column if present.
      3) Otherwise return None and require --carrier from CLI.
    """
    name = (file_path or "").lower()
    if "ups" in name:
        return CARRIER_UPS
    if "purolator" in name or "puro" in name:
        return CARRIER_PURO

    if "Invoice File" in df.columns:
        inv = str(df["Invoice File"].iloc[0]).lower()
        if "ups" in inv:
            return CARRIER_UPS
        if "purolator" in inv or "puro" in inv:
            return CARRIER_PURO

    return None


# ---------------------------------------------------------------------
# Placeholder “API” quotes (synthetic multipliers)
# Replace these later with real API calls.
# ---------------------------------------------------------------------
def quote_cpc(base_total: pd.Series, rng: np.random.Generator) -> pd.Series:
    """
    CPC: ~ +10% vs base (uniform 1.05–1.15)
    """
    mult = rng.uniform(1.05, 1.15, size=base_total.size)
    return (base_total * mult).round(2)

def quote_fedex(base_total: pd.Series, rng: np.random.Generator) -> pd.Series:
    """
    FedEx: ~ on par (uniform 0.98–1.02)
    """
    mult = rng.uniform(0.98, 1.02, size=base_total.size)
    return (base_total * mult).round(2)

def quote_other_competitor(base_total: pd.Series, rng: np.random.Generator) -> pd.Series:
    """
    Other competitor: ~ -5% vs base (uniform 0.90–1.00)
    When invoice is UPS, this yields Purolator totals; when invoice is Purolator,
    this yields UPS totals (placeholder behavior).
    """
    mult = rng.uniform(0.90, 1.00, size=base_total.size)
    return (base_total * mult).round(2)


# ---------------------------------------------------------------------
# Integration pipeline
# ---------------------------------------------------------------------
def integrate_rates(
    excel_path: str,
    out_excel: Optional[str] = None,
    carrier_hint: Optional[str] = None,
    seed: Optional[int] = 42,
) -> Tuple[str, pd.DataFrame]:
    """
    Reads uploaded invoice Excel, detects carrier (or uses carrier_hint),
    adds CPC/FedEx/Other competitor columns with synthetic quotes,
    writes an integrated Excel, and returns (output_path, integrated_df).
    """
    if not os.path.exists(excel_path):
        raise FileNotFoundError(excel_path)

    df = pd.read_excel(excel_path)
    if "Total (CAD)" not in df.columns:
        raise ValueError("Input file must contain 'Total (CAD)' (all-in total of the uploaded carrier).")

    # Alias Purolator columns so downstream stays happy
    if "Total (CAD)" not in df.columns and "Line Total (CAD)" in df.columns:
        df = df.rename(columns={"Line Total (CAD)": "Total (CAD)"})
    if "Standard Weight (lb)" not in df.columns and "Billed Weight (lb)" in df.columns:
        df = df.rename(columns={"Billed Weight (lb)": "Standard Weight (lb)"})

    # Detect carrier
    carrier = carrier_hint or detect_uploaded_carrier(df, excel_path)
    if carrier not in {CARRIER_UPS, CARRIER_PURO}:
        raise ValueError(
            "Could not detect carrier. Please pass --carrier UPS or --carrier Purolator."
        )

    # Synthetic “API” quotes based on the uploaded carrier's total
    rng = np.random.default_rng(seed)
    base = df["Total (CAD)"].astype(float)

    # Always add CPC and FedEx placeholders
    df["CPC Total"] = quote_cpc(base, rng)
    df["FedEx Total"] = quote_fedex(base, rng)

    # Add the *other* competitor (if invoice is UPS → add Purolator; if Purolator → add UPS)
    if carrier == CARRIER_UPS:
        df["Purolator Total"] = quote_other_competitor(base, rng)
    else:
        df["UPS Total"] = quote_other_competitor(base, rng)

    # Decide output path
    if out_excel is None:
        root, ext = os.path.splitext(excel_path)
        out_excel = f"{root}_integrated.xlsx"

    # Persist integrated file
    df.to_excel(out_excel, index=False)
    return out_excel, df


def run_end_to_end(
    excel_path: str,
    out_excel: Optional[str],
    out_json: Optional[str],
    carrier: Optional[str],
    seed: Optional[int],
) -> str:
    """
    1) Integrate competitor rates (synthetic for now)
    2) Generate analytics insights JSON
    3) Return output Excel path
    """
    integrated_path, _ = integrate_rates(
        excel_path=excel_path,
        out_excel=out_excel,
        carrier_hint=carrier,
        seed=seed,
    )

    insights = generate_insights(integrated_path)

    if out_json:
        with open(out_json, "w", encoding="utf-8") as f:
            import json
            json.dump(insights, f, ensure_ascii=False, indent=2)
        print(f"Wrote insights JSON → {out_json}")
    else:
        # Brief console preview
        print(f"Rows processed: {insights.get('num_rows_processed')}")
        port = insights.get("portfolio", {})
        print(f"Avg CPC vs UPS gap (simple mean): {port.get('avg_gap_pct', 0):.2%}")
        print(f"Revenue at Risk (CPC ≥ +15%): {port.get('revenue_at_risk', 0):,.2f}")
        print(f"Value Uplift (CPC ≤ −15% → move to −5%): {insights.get('portfolio', {}).get('value_uplift', 0):,.2f}")

    return integrated_path


# ---------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------
def _cli():
    p = argparse.ArgumentParser(description="Integrate competitor rates (synthetic) and run analytics.")
    p.add_argument("input", help="Path to uploaded invoice Excel (UPS or Purolator). Must contain 'Total (CAD)'.")
    p.add_argument("--carrier", choices=[CARRIER_UPS, CARRIER_PURO],
                   help="Override auto-detection (UPS or Purolator).")
    p.add_argument("--out-excel", default=None,
                   help="Path to write integrated file (default: <input>_integrated.xlsx)")
    p.add_argument("--out-json", default=None,
                   help="Optional path to write analytics insights JSON (e.g., insights.json)")
    p.add_argument("--seed", type=int, default=42, help="Random seed for synthetic quotes (default: 42)")
    args = p.parse_args()

    integrated_path = run_end_to_end(
        excel_path=args.input,
        out_excel=args.out_excel,
        out_json=args.out_json,
        carrier=args.carrier,
        seed=args.seed,
    )
    print(f"Wrote integrated Excel → {integrated_path}")

if __name__ == "__main__":
    _cli()
