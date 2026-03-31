import sqlite3
import os
from datetime import datetime
import json

# This checks if we are on Railway or running locally
# Locally it uses the old filename, on Railway it uses the Volume path
DB_PATH = "/app/data/safari_quotes.db" if os.path.exists("/app/data") else "safari_quotes.db"

def init_db():
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