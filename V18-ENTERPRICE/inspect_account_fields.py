import psycopg2

conn = psycopg2.connect(
    dbname='odoo_v18',
    user='odoo',
    password='odoo123',
    host='localhost',
    port=5432
)
cur = conn.cursor()

# Consultar líneas de impuesto y sus relaciones
cur.execute("""
    SELECT 
        aml.id, 
        aml.balance, 
        aml.tax_line_id, 
        t.name, 
        t.amount, 
        t.amount_type
    FROM account_move_line aml 
    JOIN account_tax t ON aml.tax_line_id = t.id 
    WHERE aml.display_type = 'tax' 
    LIMIT 30
""")
rows = cur.fetchall()

print("=" * 80)
print("Líneas de Impuesto en Movimientos Contables:")
print("=" * 80)
for r in rows:
    print(f"AML ID: {r[0]} | Balance: {r[1]} | Tax ID: {r[2]} | Name: {r[3]} | Amount: {r[4]} | Type: {r[5]}")

conn.close()
