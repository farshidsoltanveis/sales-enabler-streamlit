from __future__ import annotations
from typing import Dict

def _pct(x) -> str:
    try: return f"{float(x)*100:.1f}%"
    except: return "n/a"

def _money(x) -> str:
    try: return f"${float(x):,.2f}"
    except: return "n/a"

def _lane_line_short(r: Dict) -> str:
    lane = f"{r.get('origin_fsa','?')}→{r.get('destination_fsa','?')}"
    svc  = r.get("Service","?")
    wt   = r.get("weight_band","?")
    wc   = r.get("weight_class","?")
    gap  = _pct(r.get("avg_gap_pct", 0))
    return f"- {lane} | {svc} | {wt} ({wc}) | gap={gap}"

def insights_to_prompt(insights: Dict, company_name: str = "Canada Post", *, max_sig_rows: int = 3) -> str:
    port = insights.get("portfolio", {}) or {}
    comp = insights.get("competitor_summary", {}) or {}
    sig  = (insights.get("significant_gaps", []) or [])[:max_sig_rows]
    wins = insights.get("win_rates", {}) or {}
    recs = insights.get("recommended_moves", {}) or {}

    avg_gap = float(port.get("avg_gap_pct", 0.0))
    median_gap = float(port.get("median_gap_pct", avg_gap))
    direction = "cheaper" if avg_gap < 0 else "more expensive"

    # Means
    cpc_m = comp.get("cpc_mean")
    ups_m = comp.get("ups_mean")
    pur_m = comp.get("purolator_mean")
    fed_m = comp.get("fedex_mean")

    # Recommendation summaries & elasticity
    def_sum   = (recs.get("defense") or {}).get("summary") or {}
    off_sum   = (recs.get("offense") or {}).get("summary") or {}
    tgt_prem  = (recs.get("defense") or {}).get("target_premium_pct")
    tgt_disc  = (recs.get("offense") or {}).get("target_discount_pct")
    elasticity = recs.get("elasticity_assumption", -1.2)

    out = []
    out.append(f"You are a pricing strategist for {company_name}. Use ONLY the data provided. No assumptions.")
    out.append("")
    out.append("### Executive Summary")
    # 1) CPC vs UPS directionality
    out.append(f"- CPC is {direction} than UPS on average by { _pct(avg_gap) } (mean) and { _pct(median_gap) } (median).")
    # 2) Average prices of all players
    parts = []
    if cpc_m is not None: parts.append(f"CPC {_money(cpc_m)}")
    if ups_m is not None: parts.append(f"UPS {_money(ups_m)}")
    if pur_m is not None: parts.append(f"Purolator {_money(pur_m)}")
    if fed_m is not None: parts.append(f"FedEx {_money(fed_m)}")
    out.append(f"- Average prices (this volume profile): " + (", ".join(parts) if parts else "n/a") + ".")
    # 3) Meaningful gaps
    if sig:
        out.append("- Meaningful gaps (≥15%) concentrate in:")
        for r in sig:
            out.append(_lane_line_short(r))
    else:
        out.append("- No lanes with gaps ≥15% were detected.")
    # 4) Definitions + historical win rates (now explicit)
    out.append("- Definitions: Offense = CPC ≤ −10% vs UPS; Defense = CPC ≥ +10% vs UPS.")
    if wins:
        try:
            ovr = f"{int(wins.get('overall', 0)*100)}%"
            off = f"{int(wins.get('offense', 0)*100)}%"
            dfn = f"{int(wins.get('defense', 0)*100)}%"
        except Exception:
            ovr = off = dfn = "n/a"
        out.append(f"- Historical win rates (synthetic demo): overall {ovr}, offense {off}, defense {dfn}.")
    else:
        out.append("- Historical win rates: n/a.")
    # 5) Elasticity-based recommendation
    rec_bits = []
    if def_sum and def_sum.get("count",0)>0:
        rec_bits.append(f"reduce CPC by ~{ _pct(def_sum.get('mean',0)) } on defense lanes (target premium { _pct(tgt_prem) })")
    if off_sum and off_sum.get("count",0)>0:
        rec_bits.append(f"increase CPC by ~{ _pct(off_sum.get('mean',0)) } on offense lanes (target discount { _pct(-float(tgt_disc)) })")
    if rec_bits:
        out.append(f"- Considering price elasticity of {elasticity}, optimal price change would be to " + "; ".join(rec_bits) + ".")
    else:
        out.append(f"- Considering price elasticity of {elasticity}, optimal price change: hold and monitor.")
    return "\n".join(out)
