from fredapi import Fred

# Vervang met je eigen API-sleutel
fred = Fred(api_key='2c224f339cc10721c3896180e9fdb66d')

# Haal een kleine dataset op
m2_data = fred.get_series('M2SL', observation_start='2024-01-01', observation_end='2025-08-14')
print(m2_data.head())