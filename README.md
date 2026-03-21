# Biological Report Extractor

Extracts structured data from biological/environmental PDF reports using `pypdf` for text parsing and Google ADK (Gemini) for LLM-based extraction. Results are written to a formatted Excel file.

## Output

Two Excel sheets are generated:

- **Report Summary** — Title, Author, Client, Purpose, Publication Date, Location
- **Species List** — All observed/detected species with name and source section

## Requirements

```
pip install pypdf google-adk google-genai openpyxl python-dotenv
```

## Setup

Create a `.env` file in the project root:

```
GOOGLE_API_KEY=your_gemini_key
```

## Configuration

Edit the top of `my_agent/agent.py` to set:

```python
PDF_PATH = r"C:\path\to\your\report.pdf"
OUTPUT_EXCEL = "bio_report_output.xlsx"
GEMINI_MODEL = "gemini-2.0-flash"
```

## Usage

**Windows:**
```
run.bat
```

**Or directly:**
```
python my_agent/agent.py
```

When prompted, enter the section(s) containing directly observed species (e.g. `Appendix B: Plant Species Observed`), or press Enter to scan the full document.

## How It Works

1. **Parse** — `pypdf` extracts text from all PDF pages
2. **Extract** — Two LLM passes via Google ADK:
   - Pass 1: summary fields (Title, Author, Client, etc.)
   - Pass 2: species list (observed/detected only, with source section)
3. **Write** — Results saved to a styled `.xlsx` file

## Species Extraction Rules

The agent includes only species that were **directly observed or detected** on site. It excludes species listed as having potential to occur, not detected, or mentioned only in background literature.
