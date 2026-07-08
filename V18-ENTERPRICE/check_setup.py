company = env['res.company'].browse(6)
accounts = env['account.account'].search([('company_id', '=', 6)])
journals = env['account.journal'].search([('company_id', '=', 6)])
print("COMPANY NAME:", company.name)
print("ACCOUNTS COUNT:", len(accounts))
print("JOURNALS COUNT:", len(journals))
for a in accounts[:10]:
    print(f"ACCOUNT: {a.code} - {a.name}")
for j in journals:
    print(f"JOURNAL: {j.name} - {j.code}")
