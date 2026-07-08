company = env['res.company'].search([('name', 'ilike', 'Prueba')], limit=1)
if company:
    print(f"FOUND_COMPANY: {company.name} (ID: {company.id})")
else:
    print("COMPANY_NOT_FOUND")
