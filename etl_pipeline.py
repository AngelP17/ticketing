"""
ETL Pipeline: Excel to SQLite Sync
===================================
Monitors the Excel file and automatically syncs data to a relational SQLite database.
This is a Data Engineering pattern that works identically on Windows and Mac.

Run standalone: python etl_pipeline.py
Or as a Docker service: see docker-compose.yml
"""

import pandas as pd
from sqlalchemy import create_engine, text
import os
import time
from datetime import datetime

# --- CONFIGURATION ---
EXCEL_FILE = 'tickets.xlsx'
SHEET_NAME = 'IT Service Tickets'
DB_URL = 'sqlite:///local_infra.db'  # Cross-platform SQLite
CHECK_INTERVAL = 60  # Check for changes every 60 seconds


def run_etl():
    """
    Extracts data from Excel, Transforms it (cleaning), 
    and Loads it into a relational SQLite database.
    """
    if not os.path.exists(EXCEL_FILE):
        print(f"[{datetime.now()}] Warning: {EXCEL_FILE} not found. Waiting...")
        return False

    print(f"[{datetime.now()}] Starting ETL process...")

    try:
        # 1. EXTRACT
        # Read the Excel file - Pandas works the same on Mac and Windows
        df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
        
        # Rename columns to match our schema
        column_mapping = {
            'Ticket ID': 'ticket_id',
            'Title': 'title',
            'Status': 'status',
            'Priority': 'priority',
            'Request Type': 'request_type',
            'Staff Assigned': 'staff_assigned',
            'Requester': 'requester',
            'Date Opened': 'date_opened',
            'Days Open': 'days_open',
            'Description': 'description',
            'Resolution Notes': 'resolution_notes'
        }
        
        # Only rename columns that exist
        df = df.rename(columns={k: v for k, v in column_mapping.items() if k in df.columns})

        # 2. TRANSFORM (Data Cleaning)
        # Engineering step: Ensuring data integrity
        
        # Generate ticket_id if not present (based on row number)
        if 'ticket_id' not in df.columns:
            df['ticket_id'] = [f"IT-2025{i:04d}" for i in range(1, len(df) + 1)]
        
        # Remove empty rows or rows missing a Title
        if 'title' in df.columns:
            df = df.dropna(subset=['title'])
        
        # Fill missing values to prevent database NULL errors
        if 'status' in df.columns:
            df['status'] = df['status'].fillna('Open')
        if 'priority' in df.columns:
            df['priority'] = df['priority'].fillna('Low')
        if 'description' in df.columns:
            df['description'] = df['description'].fillna('')
        if 'resolution_notes' in df.columns:
            df['resolution_notes'] = df['resolution_notes'].fillna('')
        if 'request_type' in df.columns:
            df['request_type'] = df['request_type'].fillna('')
        if 'staff_assigned' in df.columns:
            df['staff_assigned'] = df['staff_assigned'].fillna('')
        if 'requester' in df.columns:
            df['requester'] = df['requester'].fillna('')
        
        # Convert Date column to standard ISO format for SQL compatibility
        if 'date_opened' in df.columns:
            df['date_opened'] = pd.to_datetime(df['date_opened'], errors='coerce')
            df['date_opened'] = df['date_opened'].dt.strftime('%Y-%m-%d')
            df['date_opened'] = df['date_opened'].fillna('')

        # 3. LOAD
        # Engineering step: Relational storage
        engine = create_engine(DB_URL)
        
        # 'replace' ensures the DB is always in sync with the source of truth (Excel)
        df.to_sql('tickets', con=engine, if_exists='replace', index=False)
        
        # Verify the data was loaded
        with engine.connect() as conn:
            result = conn.execute(text("SELECT COUNT(*) FROM tickets"))
            count = result.fetchone()[0]
        
        print(f"[{datetime.now()}] ✓ Successfully synced {count} records to {DB_URL}")
        return True

    except Exception as e:
        print(f"[{datetime.now()}] ✗ ETL Failed: {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    """Main entry point - runs the ETL loop."""
    print("=" * 50)
    print(" ETL SERVICE - Excel to SQLite Pipeline")
    print(f" Watching: {EXCEL_FILE}")
    print(f" Database: {DB_URL}")
    print(f" Interval: {CHECK_INTERVAL} seconds")
    print("=" * 50)
    print()
    
    # Run immediately on startup
    run_etl()
    
    # Then run on interval
    while True:
        time.sleep(CHECK_INTERVAL)
        run_etl()


if __name__ == "__main__":
    main()
