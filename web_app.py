import sys, traceback, uuid, os
from pathlib import Path
from datetime import datetime
import streamlit as st

# --- 0) Bootstrap & Constants ---
ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "data"
UPS_INTEGRATED = DATA_DIR / "UPS_integrated.xlsx"
PURO_INTEGRATED = DATA_DIR / "Purolator_integrated.xlsx"
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

st.set_page_config(page_title="Invoice Parser — Sales Enabler", layout="wide")

# --- 1) Header/UI scaffold ---
st.title("📄 Invoice Parser — Sales Enabler")
st.caption("Upload PDFs for one carrier per run. A per-run multi-sheet Excel is created and appended into the carrier's integrated workbook.")

# Import parsers with inline error surface so UI never stays blank
try:
    from invoice_parser.ups_parser import run_ups_parser
    from invoice_parser.purolator_parser import run_purolator_parser
    PARSERS_OK = True
except Exception:
    PARSERS_OK = False
    st.error("❌ Parser import failed. Traceback:")
    st.code("".join(traceback.format_exc()))

carrier = st.radio("Choose carrier:", ["UPS", "Purolator"], horizontal=True)

# File uploader only (no toggles)
uploaded_files = st.file_uploader(
    "Upload one or more invoice PDFs", type=["pdf"], accept_multiple_files=True
)

# --- 2) Action button ---
go = st.button("Process invoices", type="primary", use_container_width=True)

# --- 3) Helpers ---
def _persist(files, base: Path):
    base.mkdir(parents=True, exist_ok=True)
    paths = []
    for f in files:
        # Normalize filename to avoid traversal and OS issues
        safe_name = Path(f.name).name.replace("/", "_").replace("\\", "_")
        p = base / safe_name
        with open(p, "wb") as out:
            out.write(f.read())
        paths.append(p)
    return paths

# --- 4) Main action ---
if go:
    if not PARSERS_OK:
        st.stop()

    if not uploaded_files:
        st.warning("Please upload at least one PDF.")
        st.stop()

    run_id = datetime.now().strftime("%Y%m%d-%H%M%S") + "-" + uuid.uuid4().hex[:6]
    uploads = ROOT / "input" / "web_runs" / run_id
    outdir = ROOT / "output" / "web_runs" / run_id
    outdir.mkdir(parents=True, exist_ok=True)

    with st.spinner("Saving uploads…"):
        pdfs = _persist(uploaded_files, uploads)

    st.info(f"Parsing **{carrier}** invoices…")
    basename = f"{carrier.lower()}_invoices_{run_id}"

    try:
        if carrier == "UPS":
            outpath = run_ups_parser(
                [str(p) for p in pdfs],
                output_dir=str(outdir),
                output_basename=basename,
                single_sheet=False,                    # single-sheet disabled
                update_integrated=True,                # always append to integrated
                integrated_path=str(UPS_INTEGRATED),
            )
        else:  # Purolator
            outpath = run_purolator_parser(
                [str(p) for p in pdfs],
                output_dir=str(outdir),
                output_basename=basename,
                single_sheet=False,                    # single-sheet disabled
                update_integrated=True,                # always append to integrated
                integrated_path=str(PURO_INTEGRATED),
            )
    except Exception:
        st.error("❌ Parsing failed. Traceback:")
        st.code("".join(traceback.format_exc()))
        st.stop()

    # --- 5) Download (multi-sheet only) ---
    main_xlsx = Path(outpath)

    st.success("✅ Done! Your multi-sheet Excel is ready below. The integrated workbook has been updated as well.")

    if main_xlsx.exists():
        with open(main_xlsx, "rb") as f:
            st.download_button(
                f"⬇️ Download Excel (multi-sheet): {main_xlsx.name}",
                f.read(),
                file_name=main_xlsx.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    # --- 6) Integrated workbook preview (last 10 rows per sheet) ---
    try:
        import pandas as pd
        integ_path = UPS_INTEGRATED if carrier == "UPS" else PURO_INTEGRATED
        if os.path.exists(integ_path):
            sheets = pd.read_excel(integ_path, sheet_name=None, engine="openpyxl")
            st.info(f"Integrated workbook updated at: `{integ_path}`")
            with st.expander("Preview integrated workbook (last 10 rows per sheet)", expanded=False):
                for sname, df in sheets.items():
                    st.markdown(f"**{sname}** — rows: {len(df):,}")
                    if isinstance(df, pd.DataFrame) and not df.empty:
                        st.dataframe(df.tail(10), use_container_width=True)
                    else:
                        st.write("(empty)")
        else:
            st.warning("Integrated workbook path does not exist yet (no qualifying rows were written).")
    except Exception:
        st.warning("Integrated workbook exists but could not be previewed.")
        st.code("".join(traceback.format_exc()))

    # --- 7) Run details ---
    with st.expander("Run details"):
        st.write("**Carrier:**", carrier)
        st.write("**Uploaded files:**", [p.name for p in pdfs])
        st.write("**Output folder:**", str(outdir))
        st.write("**Integrated path:**", str(UPS_INTEGRATED if carrier == "UPS" else PURO_INTEGRATED))
