import pandas as pd
import openpyxl
from openpyxl.chart import LineChart, BarChart, Reference

# -------------------------------
# 1. Utility function: American to Decimal Odds
# -------------------------------
"""Converts American odds to decimal odds."""
def american_to_decimal(odds):
    if pd.isna(odds):
        return 0
    return 1 + (odds / 100 if odds > 0 else 100 / abs(odds))

# -------------------------------
# 2. Excel file path
# -------------------------------
file_path = "Bet_Tracker.xlsx"

# -------------------------------
# 3. Create or load Bet Log
# -------------------------------
try:
    bet_log_df = pd.read_excel(file_path, sheet_name="Bet Log")
except FileNotFoundError:
    columns = ["Date", "Sportsbook", "Bet Type", "Selection", "Stake ($)", "Odds", "Result", "Decimal Odds"]
    bet_log_df = pd.DataFrame(columns=columns)
    bet_log_df.to_excel(file_path, sheet_name="Bet Log", index=False)

# -------------------------------
# 4. Parse Date column (short format like 9/11/25)
# -------------------------------
bet_log_df['Date'] = pd.to_datetime(bet_log_df['Date'], errors='coerce').dt.date

# -------------------------------
# 5. Convert American Odds to Decimal Odds
# -------------------------------
bet_log_df["Decimal Odds"] = bet_log_df["Odds"].apply(american_to_decimal)


# -------------------------------
# 6. Write to Excel and insert Payout & Net PnL formulas
# -------------------------------
with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
    bet_log_df.to_excel(writer, sheet_name="Bet Log", index=False)
    workbook = writer.book
    worksheet = writer.sheets['Bet Log']

    # Format Date column
    date_format = workbook.add_format({'num_format': 'mm/dd/yy'})
    worksheet.set_column('A:A', 12, date_format)

    # Insert formulas for Payout ($) and Net PnL ($)
    worksheet.write('I1', 'Payout ($)')
    worksheet.write('J1', 'Net PnL ($)')
    for row in range(2, len(bet_log_df) + 2):
        worksheet.write_formula(f'I{row}', f'=IF(G{row}="Win", F{row}*H{row}, 0)')
        worksheet.write_formula(f'J{row}', f'=I{row}-F{row}')


# -------------------------------
# 7. Load workbook for dashboard
# -------------------------------
wb = openpyxl.load_workbook(file_path)
ws_log = wb["Bet Log"]

# Remove old Dashboard
if "Dashboard" in wb.sheetnames:
    del wb["Dashboard"]
ws_dash = wb.create_sheet("Dashboard")

# -------------------------------
# 8. KPI Section
# -------------------------------
ws_dash["A1"] = "KPI"
ws_dash["B1"] = "Value"
ws_dash["A2"] = "Total PnL ($)"
ws_dash["B2"] = "=SUM('Bet Log'!J2:J1000)"
ws_dash["A3"] = "Total Stake ($)"
ws_dash["B3"] = "=SUM('Bet Log'!E2:E1000)"
ws_dash["A4"] = "Win %"
ws_dash["B4"] = '=COUNTIF(\'Bet Log\'!G2:G1000,"Win")/COUNTA(\'Bet Log\'!G2:G1000)'
ws_dash["A5"] = "ROI (%)"
ws_dash["B5"] = "=B2/B3"

# -------------------------------
# 9. Cumulative PnL Column
# -------------------------------
ws_log.cell(row=1, column=11).value = "Cumulative PnL ($)"  # Column K
for i in range(2, ws_log.max_row + 1):
    ws_log.cell(row=i, column=11).value = f"=SUM($J$2:J{i})"

# -------------------------------
# 10. Charts
# -------------------------------
# Line Chart: Cumulative PnL Over Time
line = LineChart()
line.title = "Cumulative Net PnL Over Time"
line.x_axis.title = "Date"
line.y_axis.title = "Cumulative Net PnL ($)"

dates = Reference(ws_log, min_col=1, min_row=2, max_row=ws_log.max_row)
cumulative = Reference(ws_log, min_col=11, min_row=1, max_row=ws_log.max_row)
line.add_data(cumulative, titles_from_data=True)
line.set_categories(dates)
line.height = 10
line.width = 20
ws_dash.add_chart(line, "A8")

# Bar Chart: Net PnL by Sportsbook
bar = BarChart()
bar.title = "Net PnL by Sportsbook"
bar.x_axis.title = "Sportsbook"
bar.y_axis.title = "Net PnL ($)"

sportsbooks = Reference(ws_log, min_col=2, min_row=2, max_row=ws_log.max_row)
net_pnl = Reference(ws_log, min_col=10, min_row=1, max_row=ws_log.max_row)  # Column J
bar.add_data(net_pnl, titles_from_data=True)
bar.set_categories(sportsbooks)
bar.height = 10
bar.width = 20
ws_dash.add_chart(bar, "L8")

# -------------------------------
# 11. Save workbook
# -------------------------------
wb.save(file_path)
print(f"Bet Tracker Dashboard ready: {file_path}")