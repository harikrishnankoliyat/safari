import sqlite3
from datetime import datetime
import json

def init_db():
    conn = sqlite3.connect('safari_quotes.db')
    c = conn.cursor()
    # config_json stores the text data instead of a heavy Word file
    c.execute('''CREATE TABLE IF NOT EXISTS quotes 
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, 
                  client_name TEXT, 
                  country TEXT, 
                  date_generated TEXT, 
                  config_json TEXT)''')
    conn.commit()
    conn.close()

def save_quote_data(client_name, country, config_dict):
    conn = sqlite3.connect('safari_quotes.db')
    c = conn.cursor()
    now = datetime.now().strftime("%d/%m/%Y %H:%M")
    # Convert the quote dictionary to a small text string
    config_str = json.dumps(config_dict)
    c.execute("INSERT INTO quotes (client_name, country, date_generated, config_json) VALUES (?, ?, ?, ?)",
              (client_name, country, now, config_str))
    conn.commit()
    conn.close()

def delete_quote(quote_id):
    conn = sqlite3.connect('safari_quotes.db')
    c = conn.cursor()
    c.execute("DELETE FROM quotes WHERE id = ?", (quote_id,))
    conn.commit()
    conn.close()

def search_quotes(query):
    conn = sqlite3.connect('safari_quotes.db')
    c = conn.cursor()
    c.execute("SELECT id, client_name, country, date_generated, config_json FROM quotes WHERE client_name LIKE ?", 
              ('%' + query + '%',))
    results = c.fetchall()
    conn.close()
    return results