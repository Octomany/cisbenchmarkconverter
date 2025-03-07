# CIS Benchmark Converter

**Author:** Maxime Beauchamp  
**LinkedIn:** [Maxime Beauchamp](https://www.linkedin.com/in/maxbeauchamp/)  
**Date Created:** 2024-11-06

Modifications:  
**Author:** UnBonWhisky / Flavien Fouqueray  
**LinkedIn:** [Flavien Fouqueray](https://www.linkedin.com/in/ffouqueray/)  
**Date Last Edition:** 2025-03-06

## Description

`CIS Benchmark Converter` is a Python script designed to extract recommendations from CIS Benchmark PDF documents and export them into CSV or Excel format. The output provides a structured, easy-to-read table format that simplifies compliance checks and reviews.

## Features

- Extracts recommendations, including key sections: Profile Applicability, Description, Rationale, Impact, Audit, Remediation, Default Value, References, and Additional Information.
- Supports both CSV (using `|` as a delimiter) and Excel output formats.
- Formats Excel output with styled headers, dropdowns for compliance status, and conditional formatting for easy review.

## Installation

1. Clone this repository.
2. Install dependencies using `requirements.txt`:

   ```bash
   pip install -r requirements.txt
   ```

## Usage

```bash
python cis_benchmark_converter.py -i path/to/input_file.pdf -o path/to/output_file -f [csv|excel]
```

### Arguments

- `-i, --input` : Path to the input CIS Benchmark PDF file (required).
- `-o, --output` : Path to the output file (defaults to the input file name with `.csv` or `.xlsx` extension).
- `-f, --format` : Output file format, either `csv` or `excel` (default: `excel`).

## Example

```bash
python cis_benchmark_converter.py -i ./CIS_AWS_Benchmark.pdf -o ./CIS_AWS_Benchmark.xlsx -f excel
```

## File Structure

The generated output includes the following columns:

- **Compliance Status** - Dropdown with "Compliant", "Non-Compliant", "To Review".
- **Number** - Recommendation number (e.g., 1.1.1).
- **Level** - Recommendation level (e.g., L1, L2).
- **Title** - Full title of the recommendation.
- Additional sections from the CIS Benchmark (Profile Applicability, Description, Rationale, etc.).

### Notes

- The script automatically skips sections labeled "CIS Controls" as they aren't part of the core recommendations.

## License

This script is provided under the MIT License. Respect CIS Benchmark copyright when using and sharing this tool.
