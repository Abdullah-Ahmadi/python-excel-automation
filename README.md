# Excel Batch Printing Automation

## Problem
Daily Excel documents need to be updated and printed based on driver data.
Doing this manually is repetitive and error-prone.

## Solution
This Python script automates the entire workflow:
- Reads driver data from CSV
- Updates date values in an Excel template
- Selectively prints sheets based on active records
- Logs all operations

## Input
- Excel template file
- CSV file containing driver information

## Output
- Printed Excel sheets
- Execution log file

## How to Run
1. Install requirements
2. Place input files in the `input/` folder
3. Run `python run.py`
