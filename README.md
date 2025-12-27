# V2 Overlooked & Returned – Supervisor Correction Tool

A hardened Go utility that reconciles driver data between V2 Overlooked & Returned reports and the ALLOPS workforce file, producing a validated master list and correcting supervisor names without corrupting data.

This tool is designed for real-world Excel exports where:
	•	Header rows are not always row 1
	•	Columns move between versions
	•	Invisible whitespace breaks naive parsing
	•	Duplicate records are common

⸻

## Features
	•	Header-row auto-detection (no hardcoded row numbers)
	•	Unicode-safe header normalization
	•	Correct supervisor mapping from ALLOPS
	•	Deduplicated master list (one row per colleague)
	•	In-place supervisor correction in V2 Data sheet
	•	Fails fast on ambiguous or invalid data
	•	Excel-order independent

⸻

## What This Tool Does
	1.	Reads driver names from the Data sheet in the V2 report
	2.	Matches them against Preferred First Name and Legal Last Name in ALLOPS
	3.	Builds a validated in-memory map containing:
	•	Colleague ID
	•	First Name
	•	Last Name
	•	Manager Name
	4.	Creates a new master list sheet containing one row per colleague
	5.	Corrects incorrect supervisor names in the V2 Data sheet
	6.	Writes a validated output file

⸻

## Why This Exists

Previous scripts failed because they:
	•	Assumed headers were on row 1
	•	Trusted Excel’s visual layout
	•	Allowed duplicate colleague IDs
	•	Accidentally read the wrong columns
	•	Silently corrupted supervisor data

This tool explicitly prevents all of those failure modes.

⸻

## Usage

### Requirements
	•	Go 1.20+
	•	Excel .xlsx files

### Install Dependency

go get github.com/xuri/excelize/v2

### Run

go run main.go \
  -allops ALLOPS_Workforce_Report.xlsx \
  -v2 "V2.Overlooked and Returned Report.xlsx"


⸻

### Output

V2.Overlooked_and_Returned_VALIDATED.xlsx

Contains:
	•	Updated supervisor names in the Data sheet
	•	A new master list sheet with:
	•	Colleague ID
	•	First Name
	•	Last Name
	•	Supervisor Name

Each colleague appears exactly once.

⸻

Data Integrity Guarantees

This tool will exit with an error if:
	•	Required columns cannot be unambiguously identified
	•	Header rows are ambiguous
	•	Manager name overlaps business area
	•	Files are malformed or incomplete

If the program completes successfully, the output data is valid by construction.

⸻

Expected Columns

ALLOPS Workforce File
	•	Preferred First Name
	•	Legal Last Name
	•	Colleague ID
	•	Manager - Name
	•	Business Area

V2 Overlooked & Returned File

Sheet: Data
	•	Driver Name
	•	Supervisor Name

Column order does not matter.

⸻

Safety Design
	•	No in-place modification of source files
	•	New output file is always written
	•	Header detection uses multiple anchors
	•	Supervisor data is validated before use
	•	Duplicate colleague IDs are prevented

⸻

Known Limitations
	•	Name matching is deterministic (no fuzzy matching)
	•	Middle names and suffixes are ignored
	•	Requires consistent first/last name spelling between files

These are intentional to avoid false-positive matches.

⸻

Possible Enhancements
	•	Fuzzy name matching
	•	CSV audit of unmatched drivers
	•	Unit tests for header detection
	•	Config file support (YAML)
	•	CI-friendly dry-run mode

⸻

## Ownership

Built for operational data correction where accuracy matters more than convenience.

If you modify this tool, do not weaken header detection or deduplication logic without understanding the data risks.

⸻

## License

Internal use / as-is.
Adapt as needed for your organization.