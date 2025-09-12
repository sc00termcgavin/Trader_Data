import pandas as pd
import openpyxl
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.styles import numbers

# -------------------------------
# 1. Utility function: American to Decimal Odds
# -------------------------------
def american_to_decimal(odds):
    """Converts American odds to decimal odds."""
    if pd.isna(odds):
        return 0
    if odds > 0:
        return 1 + odds / 100
    else:
        return 1 + 100 / abs(odds)

# -------------------------------
# 2. Excel file path
# -------------------------------
file_path = "Simple_Bet_Tracker.xlsx"

# -------------------------------
# 3. Check if Excel exists, else create empty Bet Log
# -------------------------------
try:
    bet_log_df = pd.read_excel(file_path, sheet_name="Bet Log")
except FileNotFoundError:
    bet_log_df = pd.DataFrame(columns=["Date", "Sportsbook", "Stake ($)", "Odds", "Result"])
    bet_log_df.to_excel(file_path, sheet_name="Bet Log", index=False)

# -------------------------------
# 4. Parse Date column (short format like 9/11/25)
# -------------------------------
bet_log_df['Date'] = pd.to_datetime(bet_log_df['Date'], errors='coerce', dayfirst=False).dt.date

# -------------------------------
# 5. Convert American Odds to Decimal Odds
# -------------------------------
bet_log_df["Decimal Odds"] = bet_log_df["Odds"].apply(american_to_decimal)

# -------------------------------
# 6. Calculate Payout and Net PnL
# -------------------------------
# bet_log_df["Payout ($)"] = bet_log_df.apply(
#     lambda x: x["Stake ($)"] * x["Decimal Odds"] if x["Result"] == "Win" else 0, axis=1
# )
# bet_log_df["Net PnL ($)"] = bet_log_df["Payout ($)"] - bet_log_df["Stake ($)"]

# We can write Excel formulas directly to update when "Result" Changes
# Column G (Payout ($))          Column H (Net PnL ($)):
# =IF(E2="Win", F2*D2, 0)        =G2 - F2
# Where,
# E2 = Result
# F2 = Stake ($)
# D2 = Decimal odds
# For bonus bets always use Stake = 0 for net PnL = payout if Win, else 0

# -------------------------------
# 6. Write to Excel and insert formulas for Payout & Net PnL
# type “Win” or “Loss” in the Result column and Net PnL updates instantly.
# -------------------------------
with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
    bet_log_df.to_excel(writer, sheet_name="Bet Log", index=False)
    workbook  = writer.book
    worksheet = writer.sheets['Bet Log']

    # Format Date
    date_format = workbook.add_format({'num_format': 'mm/dd/yy'})
    worksheet.set_column('A:A', 12, date_format)

    # Write formulas for Payout & Net PnL
    for row in range(2, len(bet_log_df)+2):  # Excel rows start at 1 + header
        # Payout ($) = IF(Result="Win", Stake*DecimalOdds, 0)
        worksheet.write_formula(f'G{row}', f'=IF(F{row}="Win", C{row}*E{row}, 0)')
        # Net PnL ($) = Payout - Stake
        worksheet.write_formula(f'H{row}', f'=G{row}-C{row}')

# -------------------------------
# 7. Save updated Bet Log to Excel with Date formatting
# -------------------------------
# with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
#     bet_log_df.to_excel(writer, sheet_name="Bet Log", index=False)
#     workbook  = writer.book
#     worksheet = writer.sheets['Bet Log']
#     # Format Date column (A) as MM/DD/YY
#     date_format = workbook.add_format({'num_format': 'mm/dd/yy'})
#     worksheet.set_column('A:A', 12, date_format)

# -------------------------------
# 8. Load workbook and add Dashboard
# -------------------------------
wb = openpyxl.load_workbook(file_path)
ws_log = wb["Bet Log"]

# Remove old Dashboard if it exists
if "Dashboard" in wb.sheetnames:
    del wb["Dashboard"]
ws_dash = wb.create_sheet("Dashboard")

# -------------------------------
# 9. KPI Section
# -------------------------------
ws_dash["A1"] = "KPI"
ws_dash["B1"] = "Value"
ws_dash["A2"] = "Total PnL ($)"
ws_dash["B2"] = "=SUM('Bet Log'!G2:G1000)"  # Net PnL column
ws_dash["A3"] = "Total Stake ($)"
ws_dash["B3"] = "=SUM('Bet Log'!C2:C1000)"
ws_dash["A4"] = "Win %"
ws_dash["B4"] = '=COUNTIF(\'Bet Log\'!E2:E1000,"Win")/COUNTA(\'Bet Log\'!E2:E1000)'
ws_dash["A5"] = "ROI (%)"
ws_dash["B5"] = "=B2/B3"

# -------------------------------
# 10. Create Cumulative PnL Column
# -------------------------------
ws_log.cell(row=1, column=8).value = "Cumulative PnL ($)"  # column H
for i in range(2, ws_log.max_row + 1):
    ws_log.cell(row=i, column=8).value = f"=SUM($G$2:G{i})"

# -------------------------------
# 11. Charts
# -------------------------------

# Line Chart: Cumulative PnL Over Time
line = LineChart()
line.title = "Cumulative Net PnL Over Time"
line.x_axis.title = "Date"
line.y_axis.title = "Cumulative Net PnL ($)"

dates = Reference(ws_log, min_col=1, min_row=2, max_row=ws_log.max_row)
cumulative_pnl = Reference(ws_log, min_col=9, min_row=1, max_row=ws_log.max_row)
line.add_data(cumulative_pnl, titles_from_data=True)
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
net_pnl = Reference(ws_log, min_col=8, min_row=1, max_row=ws_log.max_row)
bar.add_data(net_pnl, titles_from_data=True)
bar.set_categories(sportsbooks)
bar.height = 10
bar.width = 20
ws_dash.add_chart(bar, "L8")


# -------------------------------
# 12. Save Excel
# -------------------------------
wb.save(file_path)
print(f"Bet Tracker Dashboard ready: {file_path}")