pos_pm_3 = env['pos.payment.method'].search([('company_id', '=', 3)])
for pm in pos_pm_3:
    print(f"POS PM 3: {pm.name} | Journal: {pm.journal_id.name if pm.journal_id else 'None'} | RecAcc: {pm.receivable_account_id.code if pm.receivable_account_id else 'None'} | OutstandingAcc: {pm.outstanding_account_id.code if pm.outstanding_account_id else 'None'}")
