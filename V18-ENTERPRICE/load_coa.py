company = env['res.company'].browse(6)
print("Loading CoA 'generic_coa' for company:", company.name)
try:
    env['account.chart.template'].try_loading('generic_coa', company)
    print("CoA 'generic_coa' loaded successfully!")
except Exception as e:
    print("Failed to load CoA 'generic_coa':", str(e))
env.cr.commit()
