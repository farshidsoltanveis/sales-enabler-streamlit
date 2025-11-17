"""
Sales Enabler – Flask Web Application
Replaces Streamlit with a proper web app using the existing HTML/CSS design.
"""

import os
import sys
import json
import uuid
import tempfile
import importlib.util
from pathlib import Path
from datetime import datetime
from typing import Optional, List, Dict, Any

from flask import Flask, render_template, request, jsonify, send_file, send_from_directory
from werkzeug.utils import secure_filename
from werkzeug.exceptions import BadRequest
import pandas as pd

# Add project root to path
PROJECT_ROOT = Path(__file__).parent
sys.path.insert(0, str(PROJECT_ROOT))

# Load environment variables
try:
    from dotenv import load_dotenv, find_dotenv
    env_path = find_dotenv(usecwd=True)
    if env_path:
        load_dotenv(env_path, override=True)
except Exception:
    pass

app = Flask(__name__, template_folder='frontend', static_folder='frontend')
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB max file size
app.config['UPLOAD_FOLDER'] = PROJECT_ROOT / 'input' / 'web_runs'
app.config['OUTPUT_FOLDER'] = PROJECT_ROOT / 'output'

# Ensure directories exist
app.config['UPLOAD_FOLDER'].mkdir(parents=True, exist_ok=True)
app.config['OUTPUT_FOLDER'].mkdir(parents=True, exist_ok=True)

# Allowed file extensions
ALLOWED_EXTENSIONS = {'pdf', 'xlsx', 'xls'}


def allowed_file(filename: str) -> bool:
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def load_module_by_path(module_name: str, file_path: str):
    """Load a Python module from a file path."""
    spec = importlib.util.spec_from_file_location(module_name, file_path)
    if spec is None or spec.loader is None:
        raise ImportError(f"Could not load spec for {module_name} from {file_path}")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def detect_carrier_from_name(path_or_name: str) -> Optional[str]:
    """Detect carrier from filename."""
    name = os.path.basename(path_or_name).lower()
    if "ups" in name:
        return "UPS"
    if "purolator" in name or "puro" in name:
        return "Purolator"
    return None


def detect_carrier_from_pdf_bytes(pdf_bytes: bytes) -> str:
    """Detect carrier from PDF content."""
    try:
        import fitz  # PyMuPDF
        with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
            text_chunks = []
            for i, page in enumerate(doc):
                if i >= 2:
                    break
                text_chunks.append(page.get_text("text") or "")
        text = " ".join(text_chunks).lower()
    except Exception:
        return "Unknown"

    if "united parcel service" in text or "ups canada" in text or "ups" in text:
        return "UPS"
    if "purolator" in text or "puro" in text:
        return "Purolator"
    return "Unknown"


def normalize_parsed_excel(excel_path: str, carrier: str) -> str:
    """Normalize parsed Excel to ensure required columns exist."""
    try:
        df = pd.read_excel(excel_path)
    except Exception:
        dfs = pd.read_excel(excel_path, sheet_name=None)
        df = next(iter(dfs.values()))

    df = df.copy()

    if carrier == "Purolator":
        if "Total (CAD)" not in df.columns and "Line Total (CAD)" in df.columns:
            df.rename(columns={"Line Total (CAD)": "Total (CAD)"}, inplace=True)
        if "Standard Weight (lb)" not in df.columns and "Billed Weight (lb)" in df.columns:
            df.rename(columns={"Billed Weight (lb)": "Standard Weight (lb)"}, inplace=True)

    required_cols = [
        "Total (CAD)",
        "Standard Weight (lb)",
        "Service",
        "Sender Postal Code",
        "Receiver Postal Code",
    ]
    for c in required_cols:
        if c not in df.columns:
            df[c] = None

    norm_path = os.path.splitext(excel_path)[0] + "_normalized.xlsx"
    df.to_excel(norm_path, index=False)
    return norm_path


# Load modules (lazy loading)
_paths = {
    "integrations": PROJECT_ROOT / "integrations" / "integrate_competitors_rate.py",
    "analytics": PROJECT_ROOT / "analytics_engine" / "insight_generator.py",
    "summarizer": PROJECT_ROOT / "analytics_engine" / "sales_summarizer.py",
    "ups_parser": PROJECT_ROOT / "invoice_parser" / "ups_parser.py",
    "puro_parser": PROJECT_ROOT / "invoice_parser" / "purolator_parser.py",
}

# Ensure analytics_engine is on path
analytics_dir = PROJECT_ROOT / "analytics_engine"
if str(analytics_dir) not in sys.path:
    sys.path.insert(0, str(analytics_dir))


def get_parser_modules():
    """Lazy load parser modules."""
    try:
        integrate_mod = load_module_by_path("integrate_competitors_rate", str(_paths["integrations"]))
        analytics_mod = load_module_by_path("insight_generator", str(_paths["analytics"]))
        summarizer_mod = load_module_by_path("sales_summarizer", str(_paths["summarizer"]))
        ups_mod = load_module_by_path("ups_parser", str(_paths["ups_parser"]))
        puro_mod = load_module_by_path("purolator_parser", str(_paths["puro_parser"]))

        return {
            "integrate_rates": getattr(integrate_mod, "integrate_rates"),
            "generate_insights": getattr(analytics_mod, "generate_insights"),
            "summarize_with_openai": getattr(summarizer_mod, "summarize_with_openai"),
            "summarize_offline": getattr(summarizer_mod, "summarize_offline"),
            "run_ups_parser": getattr(ups_mod, "run_ups_parser"),
            "run_purolator_parser": getattr(puro_mod, "run_purolator_parser"),
        }
    except Exception as e:
        raise RuntimeError(f"Failed to load required modules: {str(e)}. Make sure all dependencies are installed.") from e


@app.route('/')
def index():
    """Serve the main HTML page."""
    return render_template('index.html')


@app.route('/api/upload', methods=['POST'])
def upload_files():
    """Handle file uploads and start processing."""
    if 'files' not in request.files:
        return jsonify({"error": "No files provided"}), 400

    files = request.files.getlist('files')
    if not files or files[0].filename == '':
        return jsonify({"error": "No files selected"}), 400

    # Create a unique run directory
    run_id = datetime.now().strftime("%Y%m%d-%H%M%S") + "-" + uuid.uuid4().hex[:6]
    run_dir = app.config['UPLOAD_FOLDER'] / run_id
    run_dir.mkdir(parents=True, exist_ok=True)

    uploaded_files = []
    carrier_counts = {"UPS": 0, "Purolator": 0, "Unknown": 0}

    try:
        for file in files:
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                file_path = run_dir / filename
                file.save(str(file_path))

                # Detect carrier
                if filename.lower().endswith('.pdf'):
                    with open(file_path, 'rb') as f:
                        pdf_bytes = f.read()
                    carrier = detect_carrier_from_pdf_bytes(pdf_bytes)
                else:
                    carrier = detect_carrier_from_name(filename) or "Unknown"

                carrier_counts[carrier] = carrier_counts.get(carrier, 0) + 1
                uploaded_files.append({
                    "filename": filename,
                    "path": str(file_path),
                    "carrier": carrier
                })

        return jsonify({
            "run_id": run_id,
            "files": uploaded_files,
            "carrier_counts": carrier_counts,
            "message": f"Uploaded {len(uploaded_files)} file(s)"
        }), 200

    except Exception as e:
        return jsonify({"error": f"Upload failed: {str(e)}"}), 500


@app.route('/api/convert', methods=['POST'])
def convert_invoices():
    """Convert PDF invoices to Excel files (Component 1: Invoice to Excel)."""
    data = request.get_json()
    run_id = data.get('run_id')
    if not run_id:
        return jsonify({"error": "run_id required"}), 400

    run_dir = app.config['UPLOAD_FOLDER'] / run_id
    if not run_dir.exists():
        return jsonify({"error": "Run directory not found"}), 404

    # Setup output directory
    parsed_dir = app.config['OUTPUT_FOLDER'] / "parsed"
    parsed_dir.mkdir(parents=True, exist_ok=True)

    try:
        modules = get_parser_modules()
        results = []
        errors = []

        # Get all files in run directory
        files = list(run_dir.glob('*'))

        for file_path in files:
            if file_path.is_dir():
                continue

            filename = file_path.name
            base_name = file_path.stem

            # Only process PDF files
            if not filename.lower().endswith('.pdf'):
                continue

            carrier = detect_carrier_from_name(filename)
            if not carrier:
                errors.append(f"{filename}: Unknown carrier")
                continue

            try:
                # Parse PDF to Excel
                if carrier == "UPS":
                    parsed_xlsx = modules["run_ups_parser"](
                        [str(file_path)],
                        output_dir=str(parsed_dir),
                        output_basename=base_name
                    )
                elif carrier == "Purolator":
                    parsed_xlsx = modules["run_purolator_parser"](
                        [str(file_path)],
                        output_dir=str(parsed_dir),
                        output_basename=base_name
                    )
                else:
                    errors.append(f"{filename}: Unsupported carrier")
                    continue

                # Normalize the parsed Excel
                normalized_excel = normalize_parsed_excel(parsed_xlsx, carrier)

                results.append({
                    "filename": filename,
                    "carrier": carrier,
                    "parsed_excel": str(parsed_dir / f"{base_name}.xlsx"),
                    "normalized_excel": normalized_excel,
                })

            except Exception as e:
                errors.append(f"{filename}: {str(e)}")

        return jsonify({
            "success": True,
            "results": results,
            "errors": errors,
            "message": f"Converted {len(results)} file(s)"
        }), 200

    except Exception as e:
        return jsonify({"error": f"Conversion failed: {str(e)}"}), 500


@app.route('/api/generate-insights', methods=['POST'])
def generate_insights():
    """Generate insights from Excel files (Component 2: Insight Generator)."""
    data = request.get_json()
    run_id = data.get('run_id')
    if not run_id:
        return jsonify({"error": "run_id required"}), 400

    run_dir = app.config['UPLOAD_FOLDER'] / run_id
    if not run_dir.exists():
        return jsonify({"error": "Run directory not found"}), 404

    # Setup output directories
    integrated_dir = app.config['OUTPUT_FOLDER'] / "integrated"
    insights_dir = app.config['OUTPUT_FOLDER'] / "insights"
    for d in [integrated_dir, insights_dir]:
        d.mkdir(parents=True, exist_ok=True)

    try:
        modules = get_parser_modules()
        results = []
        errors = []

        # Get all files in run directory
        files = list(run_dir.glob('*'))

        for file_path in files:
            if file_path.is_dir():
                continue

            filename = file_path.name
            base_name = file_path.stem

            # Only process Excel files
            if not filename.lower().endswith(('.xlsx', '.xls')):
                continue

            carrier = detect_carrier_from_name(filename)
            if not carrier:
                carrier = "UPS"  # Default for Excel files

            try:
                # Step 1: Integrate rates
                integrated_path = integrated_dir / f"{base_name}_integrated.xlsx"
                integrated_path, _ = modules["integrate_rates"](
                    excel_path=str(file_path),
                    out_excel=str(integrated_path),
                    carrier_hint=carrier,
                    seed=42,
                )

                # Step 2: Generate insights
                insights = modules["generate_insights"](str(integrated_path))
                insights_path = insights_dir / f"{base_name}_insights.json"
                with open(insights_path, 'w', encoding='utf-8') as f:
                    json.dump(insights, f, ensure_ascii=False, indent=2)

                # Step 3: Generate summary (try OpenAI, fallback to offline)
                brief_md = None
                brief_suffix = "_brief.md"
                try:
                    brief_md = modules["summarize_with_openai"](
                        insights,
                        company_name="Canada Post",
                        model=None,
                        temperature=0.25,
                        max_tokens=400,
                        max_sig_rows=6,
                    )
                    print(f"✓ Successfully generated OpenAI summary for {filename}")
                except ValueError as e:
                    # API key missing - use offline
                    print(f"⚠️  OpenAI API key not configured: {e}")
                    print(f"   Using offline summary for {filename}")
                    brief_md = modules["summarize_offline"](insights, company_name="Canada Post")
                    brief_suffix = "_brief_offline.md"
                except Exception as e:
                    # Other OpenAI errors - log and fallback
                    print(f"⚠️  OpenAI summarization failed for {filename}: {type(e).__name__}: {e}")
                    print(f"   Falling back to offline summary")
                    brief_md = modules["summarize_offline"](insights, company_name="Canada Post")
                    brief_suffix = "_brief_offline.md"

                brief_path = insights_dir / f"{base_name}{brief_suffix}"
                with open(brief_path, 'w', encoding='utf-8') as f:
                    f.write(brief_md or "")

                results.append({
                    "filename": filename,
                    "carrier": carrier,
                    "integrated_excel": str(integrated_path),
                    "insights_json": str(insights_path),
                    "brief_md": str(brief_path),
                    "insights": insights,
                })

            except Exception as e:
                errors.append(f"{filename}: {str(e)}")

        return jsonify({
            "success": True,
            "results": results,
            "errors": errors,
            "message": f"Generated insights for {len(results)} file(s)"
        }), 200

    except Exception as e:
        return jsonify({"error": f"Insight generation failed: {str(e)}"}), 500


@app.route('/api/generate-insights-from-existing', methods=['POST'])
def generate_insights_from_existing():
    """Generate insights from existing Excel files in the parsed directory."""
    # Setup output directories
    parsed_dir = app.config['OUTPUT_FOLDER'] / "parsed"
    integrated_dir = app.config['OUTPUT_FOLDER'] / "integrated"
    insights_dir = app.config['OUTPUT_FOLDER'] / "insights"
    for d in [parsed_dir, integrated_dir, insights_dir]:
        d.mkdir(parents=True, exist_ok=True)

    try:
        modules = get_parser_modules()
        results = []
        errors = []

        # Get all normalized Excel files from parsed directory
        excel_files = list(parsed_dir.glob('*_normalized.xlsx'))
        
        # If no normalized files, try regular Excel files
        if not excel_files:
            excel_files = list(parsed_dir.glob('*.xlsx'))
            excel_files = [f for f in excel_files if not f.name.endswith('_normalized.xlsx')]

        if not excel_files:
            return jsonify({
                "success": False,
                "results": [],
                "errors": ["No Excel files found in parsed directory. Please convert invoices first."],
                "message": "No files to process"
            }), 200

        for file_path in excel_files:
            filename = file_path.name
            base_name = file_path.stem.replace('_normalized', '')
            carrier = detect_carrier_from_name(filename)
            if not carrier:
                carrier = "UPS"  # Default for Excel files

            try:
                # Step 1: Integrate rates
                integrated_path = integrated_dir / f"{base_name}_integrated.xlsx"
                integrated_path, _ = modules["integrate_rates"](
                    excel_path=str(file_path),
                    out_excel=str(integrated_path),
                    carrier_hint=carrier,
                    seed=42,
                )

                # Step 2: Generate insights
                insights = modules["generate_insights"](str(integrated_path))
                insights_path = insights_dir / f"{base_name}_insights.json"
                with open(insights_path, 'w', encoding='utf-8') as f:
                    json.dump(insights, f, ensure_ascii=False, indent=2)

                # Step 3: Generate summary (try OpenAI, fallback to offline)
                brief_md = None
                brief_suffix = "_brief.md"
                try:
                    brief_md = modules["summarize_with_openai"](
                        insights,
                        company_name="Canada Post",
                        model=None,
                        temperature=0.25,
                        max_tokens=400,
                        max_sig_rows=6,
                    )
                    print(f"✓ Successfully generated OpenAI summary for {filename}")
                except ValueError as e:
                    # API key missing - use offline
                    print(f"⚠️  OpenAI API key not configured: {e}")
                    print(f"   Using offline summary for {filename}")
                    brief_md = modules["summarize_offline"](insights, company_name="Canada Post")
                    brief_suffix = "_brief_offline.md"
                except Exception as e:
                    # Other OpenAI errors - log and fallback
                    print(f"⚠️  OpenAI summarization failed for {filename}: {type(e).__name__}: {e}")
                    print(f"   Falling back to offline summary")
                    brief_md = modules["summarize_offline"](insights, company_name="Canada Post")
                    brief_suffix = "_brief_offline.md"

                brief_path = insights_dir / f"{base_name}{brief_suffix}"
                with open(brief_path, 'w', encoding='utf-8') as f:
                    f.write(brief_md or "")

                results.append({
                    "filename": filename,
                    "carrier": carrier,
                    "integrated_excel": str(integrated_path),
                    "insights_json": str(insights_path),
                    "brief_md": str(brief_path),
                    "insights": insights,
                })

            except Exception as e:
                errors.append(f"{filename}: {str(e)}")

        return jsonify({
            "success": True,
            "results": results,
            "errors": errors,
            "message": f"Generated insights for {len(results)} file(s)"
        }), 200

    except Exception as e:
        return jsonify({"error": f"Insight generation failed: {str(e)}"}), 500


@app.route('/api/process', methods=['POST'])
def process_files():
    """Process uploaded files through the full pipeline (legacy endpoint)."""
    data = request.get_json()
    run_id = data.get('run_id')
    if not run_id:
        return jsonify({"error": "run_id required"}), 400

    run_dir = app.config['UPLOAD_FOLDER'] / run_id
    if not run_dir.exists():
        return jsonify({"error": "Run directory not found"}), 404

    # Setup output directories
    parsed_dir = app.config['OUTPUT_FOLDER'] / "parsed"
    integrated_dir = app.config['OUTPUT_FOLDER'] / "integrated"
    insights_dir = app.config['OUTPUT_FOLDER'] / "insights"
    for d in [parsed_dir, integrated_dir, insights_dir]:
        d.mkdir(parents=True, exist_ok=True)

    try:
        modules = get_parser_modules()
        results = []
        errors = []

        # Get all files in run directory
        files = list(run_dir.glob('*'))
        base_prefix = datetime.now().strftime("%Y%m%d-%H%M%S")

        for file_path in files:
            if file_path.is_dir():
                continue

            filename = file_path.name
            base_name = file_path.stem
            carrier = detect_carrier_from_name(filename)

            # Skip if carrier unknown and it's a PDF
            if filename.lower().endswith('.pdf') and not carrier:
                errors.append(f"{filename}: Unknown carrier")
                continue

            # Default carrier for Excel files
            if filename.lower().endswith(('.xlsx', '.xls')) and not carrier:
                carrier = "UPS"

            try:
                # Step 1: Parse PDF if needed
                if filename.lower().endswith('.pdf'):
                    if carrier == "UPS":
                        parsed_xlsx = modules["run_ups_parser"](
                            [str(file_path)],
                            output_dir=str(parsed_dir),
                            output_basename=base_name
                        )
                    elif carrier == "Purolator":
                        parsed_xlsx = modules["run_purolator_parser"](
                            [str(file_path)],
                            output_dir=str(parsed_dir),
                            output_basename=base_name
                        )
                    else:
                        errors.append(f"{filename}: Unsupported carrier")
                        continue

                    excel_for_integration = normalize_parsed_excel(parsed_xlsx, carrier)
                else:
                    excel_for_integration = str(file_path)

                # Step 2: Integrate rates
                integrated_path = integrated_dir / f"{base_name}_integrated.xlsx"
                integrated_path, _ = modules["integrate_rates"](
                    excel_path=excel_for_integration,
                    out_excel=str(integrated_path),
                    carrier_hint=carrier,
                    seed=42,
                )

                # Step 3: Generate insights
                insights = modules["generate_insights"](str(integrated_path))
                insights_path = insights_dir / f"{base_name}_insights.json"
                with open(insights_path, 'w', encoding='utf-8') as f:
                    json.dump(insights, f, ensure_ascii=False, indent=2)

                # Step 4: Generate summary (try OpenAI, fallback to offline)
                brief_md = None
                brief_suffix = "_brief.md"
                try:
                    brief_md = modules["summarize_with_openai"](
                        insights,
                        company_name="Canada Post",
                        model=None,
                        temperature=0.25,
                        max_tokens=400,
                        max_sig_rows=6,
                    )
                except Exception:
                    brief_md = modules["summarize_offline"](insights, company_name="Canada Post")
                    brief_suffix = "_brief_offline.md"

                brief_path = insights_dir / f"{base_name}{brief_suffix}"
                with open(brief_path, 'w', encoding='utf-8') as f:
                    f.write(brief_md or "")

                results.append({
                    "filename": filename,
                    "carrier": carrier,
                    "parsed_excel": str(parsed_dir / f"{base_name}.xlsx") if filename.lower().endswith('.pdf') else None,
                    "integrated_excel": str(integrated_path),
                    "insights_json": str(insights_path),
                    "brief_md": str(brief_path),
                    "insights": insights,
                })

            except Exception as e:
                errors.append(f"{filename}: {str(e)}")

        return jsonify({
            "success": True,
            "results": results,
            "errors": errors,
            "message": f"Processed {len(results)} file(s)"
        }), 200

    except Exception as e:
        return jsonify({"error": f"Processing failed: {str(e)}"}), 500


@app.route('/api/download/<path:filepath>')
def download_file(filepath: str):
    """Download a processed file."""
    # Security: only allow downloads from output directory
    # Remove any leading slashes and normalize
    safe_path = filepath.lstrip('/').replace('\\', '/')
    if '..' in safe_path or safe_path.startswith('/'):
        return jsonify({"error": "Invalid path"}), 400

    file_path = app.config['OUTPUT_FOLDER'] / safe_path
    # Ensure the file is within the output directory (prevent directory traversal)
    try:
        file_path.resolve().relative_to(app.config['OUTPUT_FOLDER'].resolve())
    except ValueError:
        return jsonify({"error": "Invalid path"}), 400

    if not file_path.exists():
        return jsonify({"error": "File not found"}), 404

    return send_file(str(file_path), as_attachment=True)


@app.route('/static/<path:filename>')
def static_files(filename: str):
    """Serve static files from frontend directory."""
    return send_from_directory(app.static_folder, filename)


# Serve root static assets
@app.route('/screenshot.png')
def screenshot():
    return send_from_directory(app.static_folder, 'screenshot.png')


@app.route('/api/health')
def health_check():
    """Health check endpoint that shows environment status."""
    import os
    api_key_set = bool(os.getenv('OPENAI_API_KEY'))
    api_key_preview = "Set" if api_key_set else "Not Set"
    if api_key_set:
        key_value = os.getenv('OPENAI_API_KEY', '')
        # Show first 7 and last 4 characters for verification
        if len(key_value) > 11:
            api_key_preview = f"{key_value[:7]}...{key_value[-4:]}"
        else:
            api_key_preview = "Set (invalid format?)"
    
    return jsonify({
        "status": "ok",
        "openai_api_key": api_key_preview,
        "openai_configured": api_key_set,
        "flask_env": os.getenv('FLASK_ENV', 'not set'),
        "port": os.getenv('PORT', 'not set'),
    })


if __name__ == '__main__':
    # Use port from environment or default to 5001 for local dev (5000 is often used by AirPlay on macOS)
    # Production platforms (Render, Railway, etc.) will set PORT env var
    port_env = os.getenv('PORT')
    port = int(port_env) if port_env else 5001
    
    # Debug mode: enabled for local dev (when PORT is not set), disabled in production
    is_production = port_env is not None
    debug_mode = not is_production and os.getenv('FLASK_ENV') != 'production'
    
    print("=" * 60)
    print("Sales Enabler – Flask Web Application")
    print("=" * 60)
    print(f"Starting server on http://localhost:{port}")
    print(f"Debug mode: {debug_mode}")
    print("=" * 60)
    if not is_production:
        print(f"Open your browser to: http://localhost:{port}")
        print("=" * 60)
    app.run(debug=debug_mode, host='0.0.0.0', port=port)

