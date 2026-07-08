# Inspect journals of company 6
journals = env['account.journal'].search([('company_id', '=', 6)])
for j in journals:
    print(f"JOURNAL: {j.name} (Code: {j.code}, Type: {j.type}, ID: {j.id})")

# Inspect POS Payment Methods of company 3
pos_pm_3 = env['pos.payment.method'].search([('company_id', '=', 3)])
for pm in pos_pm_3:
    print(f"POS PM 3: {pm.name} (Journal: {pm.journal_id.name if pm.journal_id else 'None'}, RecAcc: {pm.receivable_account_id.code if pm.receivable_account_id else 'None'})")
