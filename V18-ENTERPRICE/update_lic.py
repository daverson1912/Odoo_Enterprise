import psycopg2

def update_expiration():
    try:
        conn = psycopg2.connect(host="localhost", port="5432", user="odoo", password="odoo123", dbname="odoo_v18")
        cur = conn.cursor()
        cur.execute("UPDATE ir_config_parameter SET value = '2026-12-31 23:59:59' WHERE key = 'database.expiration_date'")
        conn.commit()
        print("Licencia actualizada a 2026-12-31 23:59:59")
        cur.close()
        conn.close()
    except Exception as e:
        print("Error al actualizar: ", e)

if __name__ == "__main__":
    update_expiration()
