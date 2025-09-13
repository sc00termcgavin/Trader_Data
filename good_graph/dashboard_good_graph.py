# import openpyxl
# from openpyxl.chart import LineChart, BarChart, Reference

# file_path = "Bet_Tracker.xlsx"

# # -------------------------------
# # 1. Create Excel with headers if not exist
# # -------------------------------
# try:
#     wb = openpyxl.load_workbook(file_path)
#     ws_log = wb["Bet Log"]
# except FileNotFoundError:
#     wb = openpyxl.Workbook()
#     ws_log = wb.active
#     ws_log.title = "Bet Log"
#     headers = [
#         "Date", "Sportsbook", "Bet Type", "Selection",
#         "Stake ($)", "Odds", "Result", "Bonus",
#         "Decimal Odds", "Payout ($)", "Net PnL ($)", "Cumulative PnL ($)"
#     ]
#     ws_log.append(headers)
#     wb.save(file_path)

# # -------------------------------
# # 2. Add Dashboard sheet
# # -------------------------------
# if "Dashboard" in wb.sheetnames:
#     del wb["Dashboard"]
# ws_dash = wb.create_sheet("Dashboard")

# # KPI Section
# ws_dash["A1"] = "KPI"
# ws_dash["B1"] = "Value"
# ws_dash["A2"] = "Total PnL ($)"
# ws_dash["B2"] = "=SUM('Bet Log'!K2:K1000)"
# ws_dash["A3"] = "Total Stake ($)"
# ws_dash["B3"] = "=SUM('Bet Log'!E2:E1000)"
# ws_dash["A4"] = "Win %"
# ws_dash["B4"] = '=COUNTIF(\'Bet Log\'!G2:G1000,"Win")/COUNTA(\'Bet Log\'!G2:G1000)'
# ws_dash["A5"] = "ROI (%)"
# ws_dash["B5"] = "=B2/B3"

# # -------------------------------
# # 3. Charts
# # -------------------------------
# # Line Chart: Cumulative PnL
# line = LineChart()
# line.title = "Cumulative Net PnL Over Time"
# line.x_axis.title = "Date"
# line.y_axis.title = "Cumulative Net PnL ($)"
# dates = Reference(ws_log, min_col=1, min_row=2, max_row=ws_log.max_row)
# cumulative = Reference(ws_log, min_col=12, min_row=2, max_row=ws_log.max_row)
# line.add_data(cumulative, titles_from_data=True)
# line.set_categories(dates)
# line.height = 10
# line.width = 20
# ws_dash.add_chart(line, "A8")

# # Bar Chart: Net PnL by Sportsbook
# bar = BarChart()
# bar.title = "Net PnL by Sportsbook"
# bar.x_axis.title = "Sportsbook"
# bar.y_axis.title = "Net PnL ($)"
# sportsbooks = Reference(ws_log, min_col=2, min_row=2, max_row=ws_log.max_row)
# net_pnl = Reference(ws_log, min_col=11, min_row=2, max_row=ws_log.max_row)  # Net PnL column
# bar.add_data(net_pnl, titles_from_data=True)
# bar.set_categories(sportsbooks)
# bar.height = 10
# bar.width = 20
# ws_dash.add_chart(bar, "L8")

# wb.save(file_path)
# print(f"✅ Bet Tracker Dashboard ready: {file_path}")