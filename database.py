import sqlite3
import os
from datetime import datetime
import json

# 1. Define the persistent directory
PERSISTENT_DIR = "/app/data"
DB_NAME = "safari_quotes.db"

# 2. Determine the path
# If we are on Railway (the directory /app exists), use the volume
if os.path.exists("/app"):
    # Ensure the data directory actually exists inside the container
    if not os.path.exists(PERSISTENT_DIR):
        try:
            os.makedirs(PERSISTENT_DIR, exist_ok=True)
        except Exception:
            pass # Fallback to local if permissions fail
    
    DB_PATH = os.path.join(PERSISTENT_DIR, DB_NAME)
else:
    # Local development path
    DB_PATH = DB_NAME

def init_db():
    # Final check: if we are using the volume path, make sure the folder exists
    if DB_PATH.startswith(PERSISTENT_DIR):
        os.makedirs(PERSISTENT_DIR, exist_ok=True)
        
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS quotes 
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, 
                  client_name TEXT, 
                  country TEXT, 
                  date_generated TEXT, 
                  config_json TEXT)''')
    conn.commit()
    conn.close()

def save_quote_data(client_name, country, config_dict):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    now = datetime.now().strftime("%d/%m/%Y %H:%M")
    config_str = json.dumps(config_dict)
    c.execute("INSERT INTO quotes (client_name, country, date_generated, config_json) VALUES (?, ?, ?, ?)",
              (client_name, country, now, config_str))
    conn.commit()
    conn.close()

def delete_quote(quote_id):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("DELETE FROM quotes WHERE id = ?", (quote_id,))
    conn.commit()
    conn.close()

def search_quotes(query):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("SELECT id, client_name, country, date_generated, config_json FROM quotes WHERE client_name LIKE ?", 
              ('%' + query + '%',))
    results = c.fetchall()
    conn.close()
    return results