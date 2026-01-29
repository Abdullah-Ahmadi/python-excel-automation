import csv
import datetime
import logging
import sys
import time
import tkinter as tk
import xlwings as xw

from pathlib import Path
from tkinter import messagebox


# Initial declarations
EXCEL_FILE_NAME = "Trip Pass Format.xlsx"
CSV_FILE_NAME = "Drivers.csv"
SHEET_NAME = "Sheet1"
DATE_CELL = "G1"
REQUIRED_CSV_HEADERS = {"Name", "Plate", "Status", "Sheet Name"}
FRIDAY_WEEKDAY = 4


# Basic logging configuration
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler("script.log", mode="w", encoding="utf-8")
    ]
)


def get_base_folder():
    """Returns the folder containing the EXE or the .py file"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent


def show_error(message, root):
    """Record errors and display error message boxes"""
    logging.error(message)
    if root:
        messagebox.showerror("Error", message)


def show_info(message, root):
    """Record different operations and display an info message box"""
    logging.info(message)
    if root:
        messagebox.showinfo("Information", message)


def get_drivers(csv_path):
    """Read drivers' info from a CSV file"""
    with open(csv_path, encoding="utf-8") as file:
        drivers = list(csv.DictReader(file))

    if not drivers or not REQUIRED_CSV_HEADERS.issubset(drivers[0]):
        raise ValueError(
            f"Invalid CSV format. Required headers: {REQUIRED_CSV_HEADERS}"
        )

    return drivers


def main():
    app = None
    wb = None
    root = None

    try:
        # Initial configuration of GUI tools
        try:
            root = tk.Tk()
            root.withdraw()
        except tk.TclError:
            root = None
            logging.warning("Tkinter unavailable.")

        # Resolve file paths
        base_folder = get_base_folder()
        excel_path = base_folder / EXCEL_FILE_NAME
        csv_path = base_folder / CSV_FILE_NAME

        # Logging opening operations
        logging.info(f"Base folder: {base_folder}")
        logging.info(f"Excel file: {excel_path} (exists={excel_path.exists()})")
        logging.info(f"CSV file: {csv_path} (exists={csv_path.exists()})")

        if not excel_path.exists():
            show_error(f"{EXCEL_FILE_NAME} not found.", root)
            return

        if not csv_path.exists():
            show_error(f"{CSV_FILE_NAME} not found.", root)
            return

        # Start Excel in headless mode
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False

        try:
            app.api.Application.PrintCommunication = False
        except Exception:
            logging.warning("PrintCommunication not available.")

        time.sleep(1)

        # Open workbook
        wb = app.books.open(str(excel_path))
        logging.info("Excel workbook opened.")

        # Update date
        sheet = wb.sheets[SHEET_NAME]
        today = datetime.datetime.now()
        adjusted_date = today + datetime.timedelta(days=1)

        if adjusted_date.weekday() == FRIDAY_WEEKDAY:
            adjusted_date += datetime.timedelta(days=1)

        sheet.range(DATE_CELL).value = adjusted_date
        logging.info(f"Updated date to {adjusted_date}")

        # Read drivers
        drivers = get_drivers(csv_path)

        active_drivers = {
            (
                d["Name"].strip(),
                d["Plate"].strip(),
                d["Sheet Name"].strip()
            )
            for d in drivers
            if d["Status"].strip().lower() == "active"
        }

        logging.info(f"Active drivers: {len(active_drivers)}")

        if not active_drivers:
            show_info("No active drivers found.", root)
            return

        # Print sheets
        sheets_by_name = {s.name: s for s in wb.sheets}
        printed = 0

        for name, plate, sheet_name in active_drivers:
            if sheet_name not in sheets_by_name:
                logging.warning(f"Sheet '{sheet_name}' not found for {name}")
                continue

            try:
                sheets_by_name[sheet_name].api.PrintOut(
                    Copies=1,
                    Preview=False,
                    Collate=True
                )
                printed += 1
                logging.info(f"Printed '{sheet_name}' for {name} ({plate})")
            except Exception as e:
                logging.error(f"Print failed for {name}: {e}")

        show_info(
            f"{printed} printed successfully out of {len(active_drivers)} drivers.",
            root
        )

    except Exception as e:
        logging.exception("Fatal error occurred.")
        show_error(str(e), root)

    finally:
        # Cleaning up
        if wb:
            wb.close()
            logging.info("Workbook closed.")

        if app:
            app.quit()
            logging.info("Excel closed.")

        logging.info("Script finished.")


if __name__ == "__main__":
    main()
