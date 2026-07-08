mapping = env['account.chart.template']._get_chart_template_mapping(get_all=True)
for code in mapping.keys():
    print("CODE:", code)
