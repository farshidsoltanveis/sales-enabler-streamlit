from __future__ import annotations
import math
from typing import Dict, List, Any, Optional
import pandas as pd
import numpy as np

# ---- helpers ----
def _to_weight_band(w: float) -> str:
    if pd.isna(w): return "unknown"
    if w < 1:   return "0–1 lb"
    if w < 5:   return "1–5 lb"
    if w < 10:  return "5–10 lb"
    if w < 20:  return "10–20 lb"
    return "20+ lb"

def _classify_band(band: str) -> str:
    s = (band or "").lower()
    if "10–20" in s or "20+" in s or "20" in s: return "heavy"
    if "0–1" in s or "1–5" in s: return "light"
    if "5–10" in s: return "mid"
    return "unknown"

def _safe_mean(series: Optional[pd.Series]) -> Optional[float]:
    if series is None: return None
    try:
        x = pd.to_numeric(series, errors="coerce").dropna()
        return float(x.mean()) if len(x) else None
    except Exception:
        return None

def _pick_ups_total(df: pd.DataFrame) -> pd.Series:
    for c in ["UPS Total (CAD)", "Total (CAD)", "Billed Charge (CAD)"]:
        if c in df.columns:
            return pd.to_numeric(df[c], errors="coerce")
    # last resort: empty series aligned to df
    return pd.Series(np.nan, index=df.index)

def _summary_range(xs: List[float]) -> Dict[str, float]:
    if not xs: return {"min": 0.0, "max": 0.0, "mean": 0.0, "count": 0}
    arr = np.array(xs, dtype=float)
    return {"min": float(arr.min()), "max": float(arr.max()), "mean": float(arr.mean()), "count": len(arr)}

# ---- core ----
def generate_insights(integrated_excel_path: str) -> Dict[str, Any]:
    df = pd.read_excel(integrated_excel_path)

    # Ensure columns exist
    for c in ["Service","Sender Postal Code","Receiver Postal Code","Standard Weight (lb)","CPC Total"]:
        if c not in df.columns: df[c] = np.nan
    for c in ["Purolator Total","FedEx Total"]:
        if c not in df.columns: df[c] = np.nan

    # Numerics
    for c in ["CPC Total","Purolator Total","FedEx Total","Standard Weight (lb)"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    # UPS anchor + backfill if empty
    ups_series = _pick_ups_total(df)
    if ups_series.isna().all() and "CPC Total" in df.columns:
        # Backfill from CPC using our synthetic factor (CPC ≈ UPS * 1.10)
        ups_series = pd.to_numeric(df["CPC Total"], errors="coerce") / 1.10
    df["UPS_total_anchor"] = ups_series

    # Bands & FSAs
    if "weight_band" not in df.columns:
        df["weight_band"] = df["Standard Weight (lb)"].apply(_to_weight_band)
    df["origin_fsa"] = df["Sender Postal Code"].astype(str).str[:3]
    df["destination_fsa"] = df["Receiver Postal Code"].astype(str).str[:3]

    # Gap vs UPS
    cpc = pd.to_numeric(df.get("CPC Total"), errors="coerce")
    ups = pd.to_numeric(df.get("UPS_total_anchor"), errors="coerce")
    df["gap_pct"] = (cpc - ups) / ups

    # Portfolio stats
    gap_series = df["gap_pct"].replace([np.inf, -np.inf], np.nan).dropna()
    avg_gap = float(gap_series.mean()) if len(gap_series) else 0.0
    median_gap = float(gap_series.median()) if len(gap_series) else avg_gap

    # Competitor means
    competitor_means = {
        "cpc_mean": _safe_mean(df.get("CPC Total")),
        "ups_mean": _safe_mean(df.get("UPS_total_anchor")),
        "purolator_mean": _safe_mean(df.get("Purolator Total")),
        "fedex_mean": _safe_mean(df.get("FedEx Total")),
    }

    # Lanes aggregation
    group_cols = ["origin_fsa","destination_fsa","Service","weight_band"]
    agg = df.groupby(group_cols, dropna=False).agg(
        shipment_count=("CPC Total","size"),
        avg_cpc_total=("CPC Total","mean"),
        avg_ups_total=("UPS_total_anchor","mean"),
        avg_gap_pct=("gap_pct","mean"),
    ).reset_index()

    # Thresholds
    offense_thr   = -0.10  # CPC ≤ −10% cheaper
    defense_thr   =  0.10  # CPC ≥ +10% pricier
    unhealthy_thr =  0.15  # |gap| ≥ 15%

    top_offense = (agg[agg["avg_gap_pct"] <= offense_thr]
                   .sort_values("avg_gap_pct").head(20)).to_dict("records")
    top_defense = (agg[agg["avg_gap_pct"] >= defense_thr]
                   .sort_values("avg_gap_pct", ascending=False).head(20)).to_dict("records")
    unhealthy = (agg[agg["avg_gap_pct"].abs() >= unhealthy_thr]
                 .sort_values("avg_gap_pct", ascending=False).to_dict("records"))

    # Significant gaps (top 6)
    significant_gaps = []
    for r in unhealthy[:6]:
        significant_gaps.append({
            "origin_fsa": r.get("origin_fsa"),
            "destination_fsa": r.get("destination_fsa"),
            "Service": r.get("Service"),
            "weight_band": r.get("weight_band"),
            "weight_class": _classify_band(r.get("weight_band")),
            "avg_gap_pct": float(r.get("avg_gap_pct", 0.0)),
            "shipment_count": int(r.get("shipment_count", 0) or 0),
        })

    # Synthetic win rates (demo)
    offense_share = float((agg["avg_gap_pct"] <= offense_thr).mean()) if len(agg) else 0.0
    defense_share = float((agg["avg_gap_pct"] >= defense_thr).mean()) if len(agg) else 0.0
    neutral_share = max(0.0, 1.0 - offense_share - defense_share)
    win_rates = {
        "overall": round(0.50 + 0.05*neutral_share + 0.12*offense_share - 0.12*defense_share, 2),
        "offense": 0.66,
        "defense": 0.38,
    }

    # Targets & elasticity (demo)
    elasticity = -1.2  # default assumption for price elasticity of demand
    target_defense = 0.05   # aim to +5% premium
    target_offense = -0.05  # aim to −5% discount

    # Lane-level recommended % change toward target (not applied here; summarized)
    defense_deltas = [max(0.0, float(r["avg_gap_pct"]) - target_defense)
                      for r in agg[agg["avg_gap_pct"] >= defense_thr].to_dict("records")]
    offense_deltas = [max(0.0, target_offense - float(r["avg_gap_pct"]))
                      for r in agg[agg["avg_gap_pct"] <= offense_thr].to_dict("records")]

    recommended_moves = {
        "elasticity_assumption": elasticity,
        "defense": {
            "definition": "CPC ≥ +10% vs UPS",
            "target_premium_pct": target_defense,
            "summary": _summary_range(defense_deltas),
        },
        "offense": {
            "definition": "CPC ≤ −10% vs UPS",
            "target_discount_pct": abs(target_offense),
            "summary": _summary_range(offense_deltas),
        },
    }

    insights: Dict[str, Any] = {
        "num_rows_processed": int(len(df)),
        "portfolio": {
            "avg_gap_pct": avg_gap,
            "median_gap_pct": median_gap,
        },
        "competitor_summary": competitor_means,
        "top_offense_lanes": top_offense,
        "top_defense_lanes": top_defense,
        "unhealthy_lanes": unhealthy,
        "significant_gaps": significant_gaps,
        "win_rates": win_rates,
        "recommended_moves": recommended_moves,
    }
    return insights
