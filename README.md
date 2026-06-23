# Automated Revenue Extraction ETL Pipeline

A Python-based ETL workflow for extracting, cleaning, standardizing, and exporting operating revenue data from public-company annual reports.

## Project Overview

This project converts unstructured PDF-based financial filings into structured Excel outputs for downstream financial analysis.

The workflow covers:

1. **Extract**: Scan annual-report PDFs and identify revenue-related tables.
2. **Transform**: Clean extracted table values, detect reporting units, fix numeric formatting issues, and standardize revenue values.
3. **Load**: Export structured results into Excel for review and follow-up analysis.

## Repository Structure

```text
financial_report_revenue_etl_sanitized/
├── README.md
├── requirements.txt
├── .gitignore
└── src/
    ├── revenue_extraction_sanitized.py
    └── unit_conversion_sanitized.py
```

## Privacy and Data Handling

This GitHub-safe version removes or avoids:

- hardcoded local machine paths
- personal file names or private directories
- company-internal paths
- private environment details
- API keys or credentials
- proprietary input data

Only public or fully redacted annual reports should be used as sample inputs.

## Installation

```bash
pip install -r requirements.txt
```

## Usage

### Step 1: Extract revenue data from PDFs

```bash
python src/revenue_extraction_sanitized.py \
  --input-dir sample_data/annual_reports \
  --output-file output/revenue_extraction_output.xlsx
```

By default, the script does **not** move your PDF files. To move processed and failed files into separate folders, add:

```bash
--move-files
```

### Step 2: Standardize revenue units

```bash
python src/unit_conversion_sanitized.py \
  --input-file output/revenue_extraction_output.xlsx \
  --output-file output/revenue_extraction_output_converted.xlsx
```

## Example Output Fields

| Field | Purpose |
|---|---|
| File Name | Tracks the source annual report |
| Keyword Table Page | Records the page where the target table was found |
| Unit Text | Stores nearby extracted text used for unit detection |
| Unit | Stores the detected or standardized reporting unit |
| 2023 / 2022 / 2021 | Stores extracted revenue values |
| Modified | Flags whether numeric formatting correction was applied |

## Notes

The keyword-matching logic intentionally keeps Chinese revenue and unit terms because the source filings are Chinese-language PDFs. The code comments, function names, and output labels are written in English to make the project easier to review in a U.S. job-application context.
