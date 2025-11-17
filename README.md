# Sales Enabler

**Analytics Co-pilot for Sales Teams**

Turn raw competitor data into lane-level insights and pricing guardrails in plain language, so reps can walk into every conversation prepared.

## Features

- 📄 **Parse competitor invoices** (UPS / Purolator PDFs) → structured Excel
- 🔄 **Integrate rates** into normalized workbooks with CPC/FedEx/competitor columns
- 📊 **Generate analytics** (price gaps, lane KPIs, offense/defense lanes)
- 🤖 **AI-powered summaries** via OpenAI (with offline fallback)
- 🌐 **Modern web interface** with drag-and-drop file uploads

## Quick Start

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Set Up Environment Variables (Optional)

Create a `.env` file in the project root:

```env
OPENAI_API_KEY=your_api_key_here
OPENAI_MODEL=gpt-4o-mini  # optional, defaults to gpt-4o-mini
```

If you don't set `OPENAI_API_KEY`, the system will use an offline summary generator.

### 3. Run the Web Application

```bash
python app.py
```

The application will start on `http://localhost:5000`

Open your browser and navigate to the URL to use the web interface.

## Usage

### Web Interface

1. **Upload Files**: Drag and drop or click to upload competitor invoice PDFs (UPS or Purolator)
2. **Process**: Click "Run Full Pipeline" to:
   - Parse PDFs → Excel
   - Integrate competitor rates
   - Generate analytics insights
   - Create sales-friendly summaries
3. **Download Results**: Get parsed Excel, integrated workbooks, insights JSON, and markdown briefs

### Command Line (Test.py)

For testing without the UI:

```bash
python Test.py --input-dir input --out-dir output --openai
```

Options:
- `--input-dir`: Directory with PDF/Excel files (default: `input/`)
- `--out-dir`: Output directory (default: `output/`)
- `--carrier`: Force carrier type (UPS or Purolator)
- `--openai`: Generate AI summaries (requires OPENAI_API_KEY)
- `--quiet`: Less console output

## Project Structure

```
Sales Enabler/
├── app.py                          # Flask web application (main entry point)
├── Test.py                         # CLI testing script
├── frontend/
│   ├── index.html                  # Main web interface
│   ├── styles.css                  # Styling
│   └── screenshot.png              # Dashboard screenshot
├── invoice_parser/
│   ├── ups_parser.py              # UPS PDF parser
│   └── purolator_parser.py        # Purolator PDF parser
├── integrations/
│   └── integrate_competitors_rate.py  # Rate integration logic
├── analytics_engine/
│   ├── insight_generator.py       # Analytics computation
│   ├── analytics_to_prompt.py     # LLM prompt builder
│   └── sales_summarizer.py        # OpenAI/offline summarization
├── input/                          # Upload directory
│   └── web_runs/                   # Per-run uploads (auto-created)
└── output/                         # Results directory
    ├── parsed/                     # Parsed Excel files
    ├── integrated/                # Integrated workbooks
    └── insights/                   # JSON insights + markdown briefs
```

## API Endpoints

- `GET /` - Main web interface
- `POST /api/upload` - Upload files (multipart/form-data)
- `POST /api/process` - Process uploaded files (JSON: `{"run_id": "..."}`)
- `GET /api/download/<path>` - Download processed files
- `GET /static/<filename>` - Static assets

## Development

### Running in Development Mode

The Flask app runs in debug mode by default:

```bash
python app.py
```

### Testing the Pipeline

Use `Test.py` to test the full pipeline without the web UI:

```bash
python Test.py --input-dir input --openai --print-brief
```

## Notes

- **File Size Limit**: 100MB per file (configurable in `app.py`)
- **Supported Formats**: PDF (UPS/Purolator invoices), Excel (.xlsx, .xls)
- **Carrier Detection**: Automatic from filename or PDF content
- **Offline Mode**: Works without OpenAI API key (uses template summaries)

## Migration from Streamlit

This Flask application replaces the previous Streamlit interface (`streamlit_app.py`). The new web app:

- ✅ Uses your custom HTML/CSS design
- ✅ Provides a modern, responsive interface
- ✅ Supports drag-and-drop file uploads
- ✅ Shows real-time progress and results
- ✅ Maintains all pipeline functionality from `Test.py`

You can remove `streamlit_app.py` if you no longer need it.

## License

Internal tool for CPC sales teams.

