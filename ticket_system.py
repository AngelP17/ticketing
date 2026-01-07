
import sqlite3
import pandas as pd
import os
from datetime import datetime, time

class TicketManager:
    def __init__(self, db_path='local_infra.db', schema_path='schema.sql'):
        self.db_path = db_path
        self.schema_path = schema_path
        self._initialize_db()

    def _initialize_db(self):
        """Creates the database and tickets table if they don't exist."""
        with sqlite3.connect(self.db_path) as conn:
            with open(self.schema_path, 'r') as f:
                conn.executescript(f.read())

    def import_csv(self, file_path):
        """
        Imports tickets from an Excel file (named csv for compatibility) into the SQLite database.
        Assumes data starts after a header row (approx row 46 in the specific file).
        """
        print(f"Loading data from {file_path}...")
        
        try:
            # Load Excel, finding the header at row 45 (0-indexed)
            # We use header=45 so that 'Ticket ID', etc become columns.
            df = pd.read_excel(file_path, header=45)
            
            # The columns based on inspection:
            # 0: NaN
            # 1: Ticket ID
            # 2: Title
            # 3: Status
            # 4: Priority
            # 5: Date
            # 6: NaN (Location?) -> Default to HQ
            # 7: Status/Request Type?
            
            # Filter valid rows: Ticket ID must not be NaN and must start with "IT-"
            # We rename columns by index to be sure
            df.columns.values[1] = 'ticket_id'
            df.columns.values[2] = 'title'
            df.columns.values[3] = 'status'
            df.columns.values[4] = 'priority'
            df.columns.values[5] = 'date_opened'
            
            # Select only rows with valid Ticket IDs
            df = df[df['ticket_id'].astype(str).str.startswith('IT-', na=False)]
            
            print(f"Found {len(df)} tickets to import.")

            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                
                for _, row in df.iterrows():
                    # Extract & Clean
                    t_id = str(row['ticket_id']).strip()
                    title = str(row['title']).strip()
                    status = str(row['status']).strip()
                    prio = str(row['priority']).strip()
                    
                    # Date handling
                    date_val = row['date_opened']
                    if pd.isnull(date_val):
                        date_str = datetime.now().strftime('%Y-%m-%d')
                    else:
                        try:
                            # Handle datetime.time objects (default to today's date + time)
                            if isinstance(date_val, time):
                                date_str = datetime.combine(datetime.today(), date_val).strftime('%Y-%m-%d')
                            else:
                                date_str = pd.to_datetime(date_val).strftime('%Y-%m-%d')
                        except Exception:
                            # Fallback if conversion fails
                            date_str = datetime.now().strftime('%Y-%m-%d')
                        
                    # Default Location since column 6 is NaN
                    loc = 'HQ' 
                    
                    # Upsert (Replace if exists)
                    cursor.execute("""
                        INSERT OR REPLACE INTO tickets 
                        (ticket_id, title, status, priority, date_opened, location, description)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                    """, (t_id, title, status, prio, date_str, loc, title))
                
                conn.commit()
                print("Import successful.")
                
        except Exception as e:
            print(f"Error importing data: {e}")
            import traceback
            traceback.print_exc()

    def create_ticket(self, ticket_data):
        """Creates a new ticket in the database."""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO tickets (ticket_id, title, status, priority, date_opened, location, description)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (
                ticket_data['ticket_id'], 
                ticket_data['title'], 
                ticket_data.get('status', 'Open'), 
                ticket_data.get('priority', 'Low'), 
                ticket_data.get('date_opened', datetime.now().strftime('%Y-%m-%d')), 
                ticket_data.get('location', 'HQ'), 
                ticket_data.get('description', '')
            ))
            conn.commit()
            return ticket_data

    def update_ticket(self, ticket_id, ticket_data):
        """Updates an existing ticket."""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            # Dynamic update query
            fields = []
            values = []
            for key, val in ticket_data.items():
                if key != 'ticket_id': # Don't update PK
                    fields.append(f"{key} = ?")
                    values.append(val)
            
            if not fields:
                return False

            values.append(ticket_id)
            query = f"UPDATE tickets SET {', '.join(fields)} WHERE ticket_id = ?"
            cursor.execute(query, values)
            conn.commit()
            return cursor.rowcount > 0

    def delete_ticket(self, ticket_id):
        """Deletes a ticket."""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute("DELETE FROM tickets WHERE ticket_id = ?", (ticket_id,))
            conn.commit()
            return cursor.rowcount > 0
