import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import CellIsRule
from datetime import datetime
import os

FILE_PATH = "Bet_Tracker.xlsx"

# -------------------------------
# Utility: Convert American odds to Decimal
# -------------------------------
def american_to_decimal(odds):
    if odds == 0 or odds is None:
        return 0
    return 1 + (odds / 100 if odds > 0 else 100 / abs(odds))

# -------------------------------
# Log a single bet
# -------------------------------
def log_bet(date, sportsbook, bet_type, selection, odds, stake=0, result="", bonus=False):
    """
    Append a single bet to Bet Tracker.
    Bonus bets: stake is 0, payout = real money won.
    """
    # Load workbook or create if not exist
    try:
        wb = openpyxl.load_workbook(FILE_PATH)
        ws = wb["Bet Log"]
    except FileNotFoundError:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Bet Log"
        headers = [
            "Date", "Sportsbook", "Bet Type", "Selection",
            "Stake ($)", "Odds", "Result", "Bonus",
            "Decimal Odds", "Payout ($)", "Net PnL ($)", "Cumulative PnL ($)"
        ]
        ws.append(headers)

    next_row = ws.max_row + 1

    # Actual stake for bonus bets
    actual_stake = 0 if bonus else stake
    dec_odds = american_to_decimal(odds)

    # Append data
    ws[f"A{next_row}"] = date
    ws[f"B{next_row}"] = sportsbook
    ws[f"C{next_row}"] = bet_type
    ws[f"D{next_row}"] = selection
    ws[f"E{next_row}"] = actual_stake
    ws[f"F{next_row}"] = odds
    ws[f"G{next_row}"] = result
    ws[f"H{next_row}"] = bonus
    ws[f"I{next_row}"] = dec_odds

    # -------------------------------
    # Payout Formula (handles Win, Push, Loss, Bonus)
    # -------------------------------
    if bonus:
        ws[f"J{next_row}"] = f'=IF(G{next_row}="Win",{stake}*(I{next_row}-1),IF(G{next_row}="Push",E{next_row},0))'
    else:
        ws[f"J{next_row}"] = f'=IF(G{next_row}="Win",E{next_row}*I{next_row},IF(G{next_row}="Push",E{next_row},0))'

    # Net PnL
    ws[f"K{next_row}"] = f'=IF(G{next_row}="", "", J{next_row}-E{next_row})'

    # Cumulative PnL
    ws[f"L{next_row}"] = f'=IF(K{next_row}="", "", SUM(K$2:K{next_row}))'

    # Conditional formatting for Net PnL
    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    ws.conditional_formatting.add(f"K2:K{ws.max_row}",
                                  CellIsRule(operator='greaterThan', formula=['0'], fill=green_fill))
    ws.conditional_formatting.add(f"K2:K{ws.max_row}",
                                  CellIsRule(operator='lessThan', formula=['0'], fill=red_fill))

    # -------------------------------
    # Dashboard - update KPIs
    # -------------------------------
    if "Dashboard" in wb.sheetnames:
        ws_dash = wb["Dashboard"]
        # Clear old values (optional)
        for row in ws_dash.iter_rows(min_row=2, max_col=2, max_row=20):
            for cell in row:
                cell.value = None
    else:
        ws_dash = wb.create_sheet("Dashboard")
        ws_dash["A1"], ws_dash["B1"] = "Metric", "Value"

    # Metrics
    ws_dash["A2"], ws_dash["B2"] = "Total PnL ($)", f"=SUM('Bet Log'!K2:K{ws.max_row})"
    ws_dash["A3"], ws_dash["B3"] = "Total Stake ($)", f"=SUM('Bet Log'!E2:E{ws.max_row})"
    ws_dash["A4"], ws_dash["B4"] = "Wins", f'=COUNTIF(\'Bet Log\'!G2:G{ws.max_row},"Win")'
    ws_dash["A5"], ws_dash["B5"] = "Total Bets", f'=COUNTA(\'Bet Log\'!G2:G{ws.max_row})'
    ws_dash["A6"], ws_dash["B6"] = "Pending Bets", f'=COUNTIF(\'Bet Log\'!G2:G{ws.max_row},"")'
    ws_dash["A7"], ws_dash["B7"] = "Win %", f"=IF(B5=0,0,B4/B5)"
    ws_dash["A8"], ws_dash["B8"] = "ROI (%)", f"=IF(B3=0,0,B2/B3)"

    wb.save(FILE_PATH)
    print(f"✅ Logged bet in row {next_row}: {bet_type} - {selection} ({sportsbook})")


# -------------------------------
# Interactive Mode
# -------------------------------
if __name__ == "__main__":
    print("📊 Bet Logger - Enter your bets (type 'done' at sportsbook to stop)\n")

    while True:
        sportsbook = input("Sportsbook (or 'done' to quit): ").strip()
        if sportsbook.lower() == "done":
            break

        # Default date = today
        date = input(f"Date (MM/DD/YY, default today {datetime.today().strftime('%m/%d/%y')}): ").strip()
        if date == "":
            date = datetime.today().strftime("%m/%d/%y")

        bet_type = input("Bet Type: ").strip()
        selection = input("Selection: ").strip()
        odds = int(input("Odds (e.g., -110, +125): ").strip())
        stake = float(input("Stake ($): ").strip())
        result = input("Result (Win/Loss/Push/blank): ").strip()
        bonus = input("Bonus bet? (y/n): ").strip().lower() == "y"

        log_bet(date, sportsbook, bet_type, selection, odds, stake, result, bonus)

    print("✅ All bets logged successfully!")


# # -------------------------------
# # Bets
# # -------------------------------

# # 1. Fanatics $10 moneyline loss gave $100 bonus bets
# log_bet(
#     date="9/5/25",
#     sportsbook="Fanatics",
#     bet_type="Moneyline",
#     selection="Kansas City Chiefs ML vs LA Chargers",
#     odds=-170,
#     stake=10,
#     result="Loss",
#     bonus=True
# )


# # 3. Hard Rock $80 EVEN loss (+125)
# log_bet(
#     date="9/11/25",
#     sportsbook="Hard Rock",
#     bet_type="Total Points Odd/Even",
#     selection="Washington Commanders vs Green Bay Packers - Even",
#     odds=125,
#     stake=80,
#     result="Loss",
#     bonus=False
# )

# # 4. Fanatics $100 bonus bet on Total Points ODD (-125) -> won $80
# log_bet(
#     date="9/11/25",
#     sportsbook="Fanatics",
#     bet_type="Total Points Odd/Even",
#     selection="Washington Commanders vs Green Bay Packers - Odd",
#     odds=-125,
#     stake=100,  # real money won from bonus
#     result="Win",
#     bonus=True
# )

# # 5. Hard Rock Bets $25 bonus bet on 4-leg parlay (+364) -> potential win of 91
# log_bet(
#     date="9/12/25",
#     sportsbook="Hard Rock",
#     bet_type="4-leg parlay, Bills, Ravens, Bangels, Chiefs = win",
#     selection="Bills, Ravens, Bangels, Chiefs = win",
#     odds=+364,
#     stake=25,  # real money won from bonus
#     result="",
#     bonus=True
# )

# # 6. Hard Rock Bets $25 bonus bet on Atlanta spread of -5.54-leg parlay (+364) -> potential win of 93.7
# log_bet(
#     date="9/12/25",
#     sportsbook="Hard Rock",
#     bet_type="Point-Spread",
#     selection="Falcons -5.5",
#     odds=+375,
#     stake=25,  # real money won from bonus
#     result="",
#     bonus=True
# )

# #7 Hard Rock Bets $25 bonus bet on Jahmyr Gibbs to have 5+ Q1 Recieving Yards (+125) -> win of 31.25
# log_bet(
#     date="9/13/25",
#     sportsbook="Hard Rock",
#     bet_type="To have 5+ Q1 Recieving Yards",
#     selection="Jahmyr Gibbs",
#     odds=+125,
#     stake=25,  # real money won from bonus
#     result="",
#     bonus=True
# )