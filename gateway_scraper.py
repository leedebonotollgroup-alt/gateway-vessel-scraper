import os
import re
import sqlite3
import pandas as pd
from datetime import datetime
from bs4 import BeautifulSoup
from playwright.sync_api import sync_playwright
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

# --- Configuration ---
GATEWAY_URL = "https://www.gatewaycontainerline.com.au/schedule"
DB_FILE = "gateway_vessel_schedule.db"
EXCEL_OUTPUT = "Gateway_Vessel_Schedule.xlsx"

# Environment Variables for Security
GATEWAY_USERNAME = os.getenv("GATEWAY_USERNAME", "leede")
GATEWAY_PASSWORD = os.getenv("GATEWAY_PASSWORD", "8g*HL#TJkZRYe?4")


def clean_text(text):
    """Utility to clean up scraped string content."""
    if not text:
        return ""
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def parse_datetime(date_str):
    """Parses various date/time formats into pandas datetime objects."""
    if not date_str or date_str in ["-", "N/A", "TBA", ""]:
        return pd.NaT

    date_str = clean_text(date_str)

    formats = [
        "%d/%m/%Y %H:%M",
        "%d/%m/%Y %H:%M:%S",
        "%d/%m/%Y",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d",
    ]

    for fmt in formats:
        try:
            return pd.to_datetime(date_str, format=fmt)
        except (ValueError, TypeError):
            continue

    return pd.to_datetime(date_str, errors="coerce")


def extract_gateway_schedule():
    """Logs into Gateway using Playwright and extracts the vessel schedule table."""
    print("Starting Playwright browser...")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        )
        page = context.new_page()

        print(f"Navigating to Gateway Schedule page: {GATEWAY_URL}")
        page.goto(GATEWAY_URL, wait_until="networkidle")

        # Check for login inputs
        username_input = page.query_selector(
            "input[name='username'], input[name='user'], input[type='text']"
        )
        password_input = page.query_selector(
            "input[name='password'], input[type='password']"
        )

        if username_input and password_input:
            print("Login fields detected. Authenticating...")
            username_input.fill(GATEWAY_USERNAME)
            password_input.fill(GATEWAY_PASSWORD)

            submit_btn = page.query_selector(
                "input[type='submit'], button[type='submit'], button:has-text('Login')"
            )
            if submit_btn:
                submit_btn.click()
            else:
                page.keyboard.press("Enter")

            page.wait_for_load_state("networkidle")
            print("Authentication complete.")

        # Ensure page content loaded
        page.wait_for_timeout(3000)
        html_content = page.content()
        browser.close()

    print("Parsing page HTML with BeautifulSoup...")
    soup = BeautifulSoup(html_content, "html.parser")

    table = soup.find("table")
    if not table:
        print(
            "Warning: No table element found on the schedule page. Retrying generic table search..."
        )
        tables = soup.find_all("table")
        if tables:
            table = tables[0]
        else:
            print("Error: Could not locate vessel schedule table.")
            return pd.DataFrame()

    headers = []
    header_row = table.find("tr")
    if header_row:
        headers = [
            clean_text(th.get_text())
            for th in header_row.find_all(["th", "td"])
            if clean_text(th.get_text())
        ]

    rows = []
    for tr in table.find_all("tr")[1:]:
        cells = [clean_text(td.get_text()) for td in tr.find_all("td")]
        if cells and len(cells) >= 3:
            rows.append(cells)

    if not rows:
        print("Error: No data rows extracted from schedule table.")
        return pd.DataFrame()

    # Determine column mapping dynamically or use default schema
    if headers and len(headers) == len(rows[0]):
        df = pd.DataFrame(rows, columns=headers)
    else:
        default_cols = [
            "Facility",
            "Vessel",
            "Voyage",
            "Lloyds Number",
            "Service",
            "ETA",
            "ETD",
            "Export Receive Start",
            "Export Haz Receive Start",
            "Export Empty Receive Start",
            "Export Reefer Cutoff",
            "Export Cargo Cutoff",
            "Export Haz Cutoff",
            "Export Empty Cutoff",
            "First Avail.",
            "First Free",
            "Import Storage Start",
            "Status",
        ]
        if len(rows[0]) == len(default_cols):
            df = pd.DataFrame(rows, columns=default_cols)
        else:
            df = pd.DataFrame(rows)

    # Standardize Column Names
    col_rename = {}
    for col in df.columns:
        c_lower = str(col).lower()
        if "lloyd" in c_lower or "imo" in c_lower:
            col_rename[col] = "Lloyds Number"
        elif "vessel" in c_lower:
            col_rename[col] = "Vessel"
        elif "voyage" in c_lower:
            col_rename[col] = "Voyage"
        elif "eta" in c_lower:
            col_rename[col] = "ETA"
        elif "etd" in c_lower:
            col_rename[col] = "ETD"

    df = df.rename(columns=col_rename)

    # Clean Lloyds Number Column
    if "Lloyds Number" in df.columns:
        df["Lloyds Number"] = (
            df["Lloyds Number"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
        )
    else:
        print("Warning: 'Lloyds Number' column not found. Creating empty column for key tracking.")
        df["Lloyds Number"] = ""

    # Parse DateTime Columns
    datetime_cols = [
        "ETA",
        "ETD",
        "Export Receive Start",
        "Export Haz Receive Start",
        "Export Empty Receive Start",
        "Export Reefer Cutoff",
        "Export Cargo Cutoff",
        "Export Haz Cutoff",
        "Export Empty Cutoff",
        "First Avail.",
        "First Free",
        "Import Storage Start",
    ]

    for col in datetime_cols:
        if col in df.columns:
            df[col] = df[col].apply(parse_datetime)

    df["Scrape Timestamp"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    print(f"Successfully scraped {len(df)} records from Gateway.")
    return df


def merge_with_historical_data(df_new):
    """
    Merges newly scraped records with existing database records based on 'Lloyds Number'.
    Preserves historical records while overwriting matching Lloyds Numbers with fresh data.
    """
    if df_new.empty:
        return df_new

    if not os.path.exists(DB_FILE):
        print("No existing database found. Starting fresh historical record.")
        return df_new

    try:
        conn = sqlite3.connect(DB_FILE)
        df_old = pd.read_sql("SELECT * FROM gateway_vessel_schedule", conn)
        conn.close()

        if df_old.empty:
            return df_new

        print(f"Found {len(df_old)} existing historical records in database.")

        # Ensure Lloyds Numbers are string types for clean comparison
        df_new["Lloyds Number"] = df_new["Lloyds Number"].astype(str).str.strip()
        df_old["Lloyds Number"] = df_old["Lloyds Number"].astype(str).str.strip()

        # Extract set of non-empty Lloyds Numbers from the new scrape batch
        new_lloyds = set(df_new[df_new["Lloyds Number"] != ""]["Lloyds Number"])

        # Retain old records whose Lloyds Numbers are NOT in the new batch
        df_old_retained = df_old[~df_old["Lloyds Number"].isin(new_lloyds)]

        # Combine old retained records with all new records
        df_combined = pd.concat([df_old_retained, df_new], ignore_index=True)

        # Convert date columns back to proper datetime objects after DB read
        datetime_cols = [
            "ETA",
            "ETD",
            "Export Receive Start",
            "Export Haz Receive Start",
            "Export Empty Receive Start",
            "Export Reefer Cutoff",
            "Export Cargo Cutoff",
            "Export Haz Cutoff",
            "Export Empty Cutoff",
            "First Avail.",
            "First Free",
            "Import Storage Start",
        ]
        for col in datetime_cols:
            if col in df_combined.columns:
                df_combined[col] = pd.to_datetime(df_combined[col], errors="coerce")

        print(
            f"Historical Merge Complete: Preserved {len(df_old_retained)} old records, updated/added {len(df_new)} records (Total: {len(df_combined)} rows)."
        )
        return df_combined

    except Exception as e:
        print(
            f"Warning: Could not read existing database for merge ({e}). Proceeding with current scrape only."
        )
        return df_new


def export_to_excel_and_sqlite(df_new):
    """Merges historical data, then saves to SQLite and exports formatted Excel Table."""
    if df_new.empty:
        print("No data extracted. Skipping export.")
        return

    # Merge with existing historical data before saving
    df_final = merge_with_historical_data(df_new)

    print(f"\n1. Saving {len(df_final)} total records to SQLite database: {DB_FILE}...")

    # Write to SQLite Database
    conn = sqlite3.connect(DB_FILE)
    df_to_sql = df_final.copy()
    
    # Format datetime columns as ISO strings for SQLite compatibility
    for col in df_to_sql.select_dtypes(include=["datetime64[ns]", "datetime64"]):
        df_to_sql[col] = df_to_sql[col].dt.strftime("%Y-%m-%d %H:%M:%S")

    df_to_sql.to_sql("gateway_vessel_schedule", conn, if_exists="replace", index=False)
    conn.close()

    print(f"2. Exporting formatted Excel Table to: {EXCEL_OUTPUT}...")

    with pd.ExcelWriter(
        EXCEL_OUTPUT,
        engine="openpyxl",
        datetime_format="YYYY-MM-DD HH:MM",
        date_format="YYYY-MM-DD",
    ) as writer:
        sheet_name = "Gateway Schedule"
        df_final.to_excel(writer, sheet_name=sheet_name, index=False)

        ws = writer.sheets[sheet_name]
        max_row = len(df_final) + 1
        max_col = len(df_final.columns)
        col_letter = get_column_letter(max_col)

        # Apply Native Excel Table Formatting
        tab_range = f"A1:{col_letter}{max_row}"
        table = Table(displayName="GatewayVesselScheduleList", ref=tab_range)
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium9", showRowStripes=True
        )
        ws.add_table(table)

        # Auto-adjust column widths
        for col in ws.columns:
            max_len = 0
            for cell in col:
                if isinstance(cell.value, datetime):
                    val_str = cell.value.strftime("%Y-%m-%d %H:%M")
                else:
                    val_str = str(cell.value or "")
                max_len = max(max_len, len(val_str))

            col_letter_idx = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter_idx].width = max(max_len + 3, 12)

    print(f"\nSUCCESS: Exported {len(df_final)} historical and current rows to Excel and SQLite!")


def main():
    """Main execution function."""
    print("=" * 60)
    print("Gateway Container Line Schedule Scraper - Historical Upsert Mode")
    print("=" * 60)

    df_new = extract_gateway_schedule()
    export_to_excel_and_sqlite(df_new)


if __name__ == "__main__":
    main()
