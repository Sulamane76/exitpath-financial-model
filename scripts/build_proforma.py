import pandas as pd, numpy as np, yaml

# Load assumptions
with open('data/assumptions.yaml') as f:
    scenarios = yaml.safe_load(f)

years = [2025, 2026, 2027, 2028, 2029]
market_price = 10000
company_price = 30000
ready_price = 50000
go_fee_pct = 0.015
deal_size = 100_000_000

for scenario, cfg in scenarios.items():
    operator_customers = np.linspace(cfg['operator_start'],
                                     cfg['operator_end_2025'],
                                     len(years))
    market_fit_arr = operator_customers * market_price
    company_fit_arr = operator_customers * np.array(cfg['graduation_rate']) * company_price
    ready_arr = np.array(cfg['ready_customers']) * ready_price
    operator_arr = market_fit_arr + company_fit_arr + ready_arr

    transaction_rev = np.array(cfg['go_probability']) * deal_size * go_fee_pct
    investor_services = np.array(cfg['investor_services'])
    investor_licenses = np.array(cfg['investor_licenses'])

    total_rev = operator_arr + transaction_rev + investor_services + investor_licenses

    df = pd.DataFrame({
        'Year': years,
        'Operator ARR': operator_arr.round(0),
        'Transaction Revenue': transaction_rev.round(0),
        'Investor Services': investor_services,
        'Investor Licenses': investor_licenses,
        'Total Revenue': total_r_



