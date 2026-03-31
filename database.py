import psycopg2
import os
from datetime import datetime
import json

# Railway provides DATABASE_URL automatically from your Variables screen
DATABASE_URL = os.getenv("DATABASE_URL")

def get_connection():
    # sslmode='require' is necessary for secure Railway connections
    return psycopg2.connect(DATABASE_URL, sslmode='require')

def init_db():
    conn = get_connection()
    cur = conn.cursor()
    # Postgres uses SERIAL for auto-incrementing IDs
    cur.execute('''CREATE TABLE IF NOT EXISTS quotes 
                 (id SERIAL PRIMARY KEY, 
                  client_name TEXT, 
                  country TEXT, 
                  date_generated TEXT, 
                  config_json TEXT)''')
    conn.commit()
    cur.close()
    conn.close()

def save_quote_data(client_name, country, config_dict):
    conn = get_connection()
    cur = conn.cursor()
    now = datetime.now().strftime("%d/%m/%Y %H:%M")
    config_str = json.dumps(config_dict)
    # Postgres uses %s placeholders instead of ?
    cur.execute("INSERT INTO quotes (client_name, country, date_generated, config_json) VALUES (%s, %s, %s, %s)",
              (client_name, country, now, config_str))
    conn.commit()
    cur.close()
    conn.close()

def delete_quote(quote_id):
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("DELETE FROM quotes WHERE id = %s", (quote_id,))
    conn.commit()
    cur.close()
    conn.close()

def search_quotes(query):
    conn = get_connection()
    cur = conn.cursor()
    # ILIKE searches both the client_name and the config_json (Tour Code)
    cur.execute("""
        SELECT id, client_name, country, date_generated, config_json 
        FROM quotes 
        WHERE client_name ILIKE %s OR config_json ILIKE %s
    """, ('%' + query + '%', '%' + query + '%'))
    results = cur.fetchall()
    cur.close()
    conn.close()
    return results