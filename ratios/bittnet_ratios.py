import os
import pandas as pd
import numpy as np

# Reloading the Excel file paths
files = {
    2020: r"D:\Fabiz\Licenta\risk_and_return_calc\ratios\reports\2020.xlsx",
    2021: r"D:\Fabiz\Licenta\risk_and_return_calc\ratios\reports\2021.xlsx",
    2022: r"D:\Fabiz\Licenta\risk_and_return_calc\ratios\reports\2022.xlsx",
    2023: r"D:\Fabiz\Licenta\risk_and_return_calc\ratios\reports\2023.xlsx",
    2024: r"D:\Fabiz\Licenta\risk_and_return_calc\ratios\reports\2024.xlsx",
}

net_profits = []
liquidity_results = []
leverage_results = []
profitability_results = []
efficiency_results = []
dupont_results = []

def find_value(df, keyword):
    """
    Searches column A (column 0) for a keyword and returns the value from column B (column 1) in the same row.
    """
    for i in range(len(df)):
        cell = str(df.iat[i, 0]).lower()  # look only in column A
        if keyword.lower() in cell:
            try:
                value = str(df.iat[i, 1]).replace(',', '')  # take the value from column B
                return float(value)
            except:
                return None
    return None

# Processing each Excel file
for year, path in files.items():
    xls = pd.ExcelFile(path)
    bs = xls.parse("Balance Sheet", header=None)
    isheet = xls.parse("Income Statement", header=None)

    # Extracting net profit for the year
    net_profit = find_value(isheet, "net profit")
    net_profits.append(net_profit)

    # Liquidity calculations
    current_assets = find_value(bs, "total current assets")
    current_liabilities = find_value(bs, "total current liabilities")
    cash = find_value(bs, "cash and cash equivalents")

    # Leverage calculations
    total_debt = find_value(bs, "total liabilities")
    ebit = find_value(isheet, "ebit")
    interest_expense = abs(find_value(isheet, "interest expense"))

    # Profitability variables
    net_profit = find_value(isheet, "net profit")
    gross_margin = find_value(isheet, "gross margin")
    revenues = find_value(isheet, "revenues from customers")
    total_assets = find_value(bs, "total assets")
    total_equity = find_value(bs, "total equity")

    # Efficiency variables
    cogs = abs(find_value(isheet, "cost of sales"))
    inventory = find_value(bs, "inventories")
    receivables = find_value(bs, "trade receivables and other receivables")


    #CALCULATIONS

    # Convert net profits to a NumPy array for further calculations
    net_profits_array = np.array(net_profits)

    # Calculate statistics for net profits
    mean_np = np.mean(net_profits_array)
    var_np = np.var(net_profits_array)
    std_dev = np.std(net_profits_array)


    # Liquidity calculations
    current_ratio = current_assets / current_liabilities if current_assets and current_liabilities else None
    quick_ratio = ((current_assets - inventory) / current_liabilities) if current_assets and inventory and current_liabilities else None
    cash_ratio = cash / current_liabilities if cash and current_liabilities else None

    # Leverage calculations
    debt_ratio = total_debt / total_assets if total_debt and total_assets else None
    debt_to_equity = total_debt / total_equity if total_debt and total_equity else None
    times_interest_earned = ebit / interest_expense if ebit and interest_expense else None

    # Profitability calculations
    roa = (net_profit / total_assets) * 100 if total_assets else None
    roe = (net_profit / total_equity) * 100 if total_equity else None
    gross_profit_margin = (gross_margin / revenues) * 100 if revenues else None

    # Efficiency calculations
    inventory_turnover = cogs / inventory if inventory else None
    asset_turnover = revenues / total_assets if total_assets else None
    receivables_turnover = revenues / receivables if receivables else None

    # DuPont components
    net_profit_margin = net_profit / revenues if revenues else None
    asset_turnover_dupont = revenues / total_assets if total_assets else None
    equity_multiplier = total_assets / total_equity if total_equity else None

    # Final DuPont ROE calculation
    dupont = (
        net_profit_margin * asset_turnover_dupont * equity_multiplier
        if net_profit_margin and asset_turnover_dupont and equity_multiplier
        else None
    )

    # Append results to respective lists

    liquidity_results.append({
        "Year": year,
        "Current Ratio": current_ratio,
        "Quick Ratio": quick_ratio,
        "Cash Ratio": cash_ratio
    })

    leverage_results.append({
        "Year": year,
        "Debt Ratio": debt_ratio,
        "Debt to Equity": debt_to_equity,
        "Times Interest Earned": times_interest_earned
    })

    profitability_results.append({
        "Year": year,
        "ROA (%)": roa,
        "ROE (%)": roe,
        "Gross Profit Margin (%)": gross_profit_margin
    })

    efficiency_results.append({
        "Year": year,
        "Inventory Turnover": inventory_turnover,
        "Asset Turnover": asset_turnover,
        "Receivables Turnover": receivables_turnover
    })

    dupont_results.append({
        "Year": year,
        "Net Profit Margin": net_profit_margin,
        "Asset Turnover": asset_turnover_dupont,
        "Equity Multiplier": equity_multiplier,
        "DuPont": dupont
    })

    
# Convert to DataFrames and round
df_stats = pd.DataFrame({
    "Metric": ["Mean Net Profit", "Variance Net Profit", "Standard Deviation Net Profit"],
    "Value": [mean_np, var_np, std_dev]
})
df_liquidity = pd.DataFrame(liquidity_results).round(3)
df_leverage = pd.DataFrame(leverage_results).round(3)
df_profitability = pd.DataFrame(profitability_results).round(3)
df_efficiency = pd.DataFrame(efficiency_results).round(3)
df_dupont = pd.DataFrame(dupont_results).round(3)

# Combine into one Excel file with separate sheets
output_path = "outputs/ratios.xlsx"

if os.path.exists(output_path):
    mode = 'a'  # append safely
else:
    mode = 'w'  # create new file if not exists

with pd.ExcelWriter(output_path, mode=mode, if_sheet_exists='replace') as writer:
    df_stats.to_excel(writer, sheet_name="Risk Statistics", index=False)
    df_liquidity.to_excel(writer, sheet_name="Liquidity Ratios", index=False)
    df_leverage.to_excel(writer, sheet_name="Leverage Ratios", index=False)
    df_profitability.to_excel(writer, sheet_name="Profitability Ratios", index=False)
    df_efficiency.to_excel(writer, sheet_name="Efficiency Ratios", index=False)
    df_dupont.to_excel(writer, sheet_name="DuPont Analysis", index=False)


print("Ratios: ")
#print(df_stats)
# print(df_liquidity)
# print(df_leverage)
# print(df_profitability)
# print(df_efficiency)
print(df_dupont)