# streamlit_app.py
import os, uuid, tempfile
from pathlib import Path
from datetime import datetime

import streamlit as st
import pandas as pd
import fitz  # PyMuPDF for carrier detection

st.set_page_config(page_title="Sales Enabler — Invoice Parser", layout="wide")
st.title("📄 Sales Enabler — Invoice Parser (per-run output)")

# ---- Import your parsers with an error guard so we don't get a blank page
try:
    from invoice_parser.ups_parser import run_ups_parser
    from invoice_parser.purolator_parser import run_purolator_parser
except Exception as e:
    st.error("Import failed. Check requirements and package structure.")
    st.exception(e)
    st.stop()

# ---- Helpers
def write_bytes(path: Path, data: bytes):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(data)

def detect_carrier(pdf_bytes: bytes) -> str:
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        text = "\n".join(p.get_text("text") for p in doc).upper()
        if "PUROLATOR" in text:
            return "Purolator"
        if "UNITED PARCEL" in text or "UPS" in text:
            return "UPS"
    except Exception:
        pass
    return "Unknown"

# ---- UI
st.caption("Upload UPS and/or Purolator invoices (PDF). We auto-detect the carrier and return per-run Excel downloads. No persistent storage.")
files = st.file_uploader("Upload one or more PDFs", type=["pdf"], accept_multiple_files=True)
go = st.button("Process invoices", type="primary", use_container_width=True)

if go:
    if not files:
        st.warning("Please upload at least one PDF.")
        st.stop()

    # Per-run temp workspace
    run_id = datetime.now().strftime("%Y%m%d-%H%M%S") + "-" + uuid.uuid4().hex[:6]
    root = Path(tempfile.mkdtemp(prefix=f"sales-enabler-{run_id}-"))
    in_dir  = root / "input"
    out_dir = root / "output"
    in_dir.mkdir(parents=True, exist_ok=True)
    out_dir.mkdir(parents=True, exist_ok=True)

    # Save and bucket PDFs by carrier
    ups_paths, puro_paths, unknown = [], [], []
    with st.spinner("Saving uploads and detecting carriers…"):
        for f in files:
            data = f.read()
            carrier = detect_carrier(data)
            p = in_dir / Path(f.name).name.replace("/", "_").replace("\\", "_")
            write_bytes(p, data)
            if carrier == "UPS":
                ups_paths.append(str(p))
            elif carrier == "Purolator":
                puro_paths.append(str(p))
            else:
                unknown.append(str(p))

    c1, c2 = st.columns(2)
    with c1: st.info(f"UPS files: **{len(ups_paths)}**")
    with c2: st.info(f"Purolator files: **{len(puro_paths)}**")
    if unknown:
        st.warning(f"Unknown carrier for {len(unknown)} file(s). They were skipped.")

    # Wrappers so minor signature differences don't break deploy
    def call_ups(paths):
        base = f"ups_invoices_{run_id}"
        try:
            return run_ups_parser(
                paths,
                output_dir=str(out_dir),
                output_basename=base,
                single_sheet=False,
                update_integrated=False,           # no persistent storage
            )
        except TypeError:
            # Older signature without update_integrated
            return run_ups_parser(
                paths,
                output_dir=str(out_dir),
                output_basename=base,
                single_sheet=False,
            )

    def call_puro(paths):
        base = f"purolator_invoices_{run_id}"
        try:
            return run_purolator_parser(
                paths,
                output_dir=str(out_dir),
                output_basename=base,
                single_sheet=False,
                update_integrated=False,           # no persistent storage
            )
        except TypeError:
            return run_purolator_parser(
                paths,
                output_dir=str(out_dir),
                output_basename=base,
                single_sheet=False,
            )

    outputs, errors = [], []

    if ups_paths:
        with st.spinner("Parsing UPS…"):
            try:
                outputs.append(("UPS", call_ups(ups_paths)))
            except Exception as e:
                errors.append(("UPS", e))

    if puro_paths:
        with st.spinner("Parsing Purolator…"):
            try:
                outputs.append(("Purolator", call_puro(puro_paths)))
            except Exception as e:
                errors.append(("Purolator", e))

    # Show parser errors (if any) right on the page
    for label, err in errors:
        st.error(f"{label} parsing failed:")
        st.exception(err)

    # Download buttons for outputs
    if not outputs:
        st.warning("No outputs produced.")
        st.stop()

    st.success("✅ Done. Download your Excel file(s):")
    for label, outpath in outputs:
        p = Path(outpath)
        if p.exists():
            st.download_button(
                f"⬇️ Download {label} Excel: {p.name}",
                p.read_bytes(),
                file_name=p.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        else:
            st.warning(f"{label} output not found at: {outpath}")

    with st.expander("Notes"):
        st.markdown(
            "- This build does **not** append to integrated workbooks.\n"
            "- Files are processed in a temporary workspace each run.\n"
            "- If you need persistent integrated Excel later, we’ll add object storage."
        )
