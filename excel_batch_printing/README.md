# Excel-Based Document Automation and Batch Printing with Python

![Python](https://img.shields.io/badge/Python-3.x-blue?logo=python&logoColor=white)
![Automation](https://img.shields.io/badge/Type-Excel%20Automation-success)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20Excel-lightgrey)
![License](https://img.shields.io/badge/License-MIT-yellow)
![Status](https://img.shields.io/badge/Status-Stable-brightgreen)

## Overview

Many operational workflows still rely on repetitive, rule-based tasks that consume a significant amount of time. When such tasks recur daily in the same form—such as preparing and printing Excel-based documents—the manual effort becomes inefficient and prone to errors.

This project demonstrates a Python-based automation designed to eliminate this repetitiveness. It automates the preparation and batch printing of Excel documents using a predefined template and structured input data, allowing the entire process to run reliably with minimal user interaction.

In its current implementation, the automation is used to print daily trip passes for drivers in a logistics operation. However, the underlying approach can be applied to many Excel-driven workflows that rely on templates, external data sources, and repetitive daily execution.

## What the Automation Does

The automation executes a complete end-to-end workflow in a controlled and reliable manner:

- Opens an Excel workbook in headless mode to avoid unnecessary UI overhead
- Updates date values in a centralized template cell, allowing linked sheets to refresh automatically
- Reads structured input data from a CSV file and validates its format
- Filters records according to business rules (e.g., active vs. inactive entries)
- Matches data records to corresponding Excel sheets
- Performs batch printing of matched sheets without manual intervention
- Records all operations, warnings, and errors in a persistent log file
- Provides user feedback through message dialogs upon completion or failure
- Cleans up all resources by properly closing files and terminating the Excel process

**Result:** a repeatable, production-ready Excel document automation that can be executed daily without manual interaction with Excel.

## Key Features

- Headless execution to prevent UI interruptions and ensure smooth background operation
- Integration with external CSV data for dynamic and data-driven processing
- Business-rule-based date handling to support real-world operational requirements
- Batch printing of Excel-based documents for efficient daily execution
- Detailed execution logging for traceability, auditing, and troubleshooting
- Graceful error handling to prevent unexpected crashes and provide clear feedback
- Proper resource cleanup to avoid memory leaks and ensure system stability

## Technical Stack

- **Python** – Core language used to implement the automation logic
- **xlwings** – Headless Excel automation, workbook manipulation, and batch printing
- **csv** – Reading and validating structured input data
- **logging** – Persistent execution logging for monitoring and troubleshooting
- **tkinter** – Lightweight user notifications via message dialogs
- **datetime** – Date handling and business-rule-based adjustments
- **pathlib** – Dynamic and executable-safe file path resolution
- **sys / time** – Runtime environment detection and execution control

## Why This Matters

- Reduces repetitive manual work by automating Excel-based document workflows
- Minimizes human error in daily operational processes
- Saves time by executing batch operations without user intervention
- Demonstrates a production-ready approach to Excel automation
- Can be adapted to similar template-driven workflows across different industries
- Designed with reliability and repeatability in mind for daily execution

## Typical Use Cases

- Automating daily Excel-based reports
- Batch printing documents based on CSV data
- Updating template-driven Excel files automatically
- Reducing manual Excel work in small business operations

## How It Is Used

1. Place the script (or compiled executable), Excel template, and CSV file in the same folder
2. Update the Excel template as needed
3. Maintain driver or record data in the CSV file
4. Run the script
5. Review printed output and execution logs

The automation is designed to run with minimal user interaction once configured.

## Requirements & Notes

- Microsoft Excel must be installed on the system
- Intended for controlled, internal-use environments
- Designed to run silently without disrupting the user
- Built with reliability and repeatability in mind for daily execution

> Note: This project is provided as a demonstration of Excel automation techniques and may require adaptation for use in different environments.

## Author

**Abdulla Ahmadi**
Python Automation • Excel Workflow Automation
