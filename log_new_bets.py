import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import CellIsRule
import pandas as pd

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
    # Load workbook
    wb = openpyxl.load_workbook(FILE_PATH)
    ws = wb["Bet Log"]

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

    # Payout formula
    # For bonus bets: payout = (stake * (decimal odds - 1)) if Win, else 0
    ws[f"J{next_row}"] = f'=IF(G{next_row}="Win", IF(H{next_row}=TRUE, {stake}*(I{next_row}-1), E{next_row}*I{next_row}), 0)'


    # Net PnL
    ws[f"K{next_row}"] = f'=J{next_row}-E{next_row}'

    # Cumulative PnL
    ws[f"L{next_row}"] = f'=SUM(K$2:K{next_row})'

    # Conditional formatting for Net PnL
    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    ws.conditional_formatting.add(f"K2:K{ws.max_row}",
                                  CellIsRule(operator='greaterThan', formula=['0'], fill=green_fill))
    ws.conditional_formatting.add(f"K2:K{ws.max_row}",
                                  CellIsRule(operator='lessThan', formula=['0'], fill=red_fill))

    wb.save(FILE_PATH)
    print(f"✅ Logged bet in row {next_row}: {bet_type} - {selection} ({sportsbook})")

# -------------------------------
# Bets
# -------------------------------

# 1. Fanatics $10 moneyline loss
log_bet(
    date="9/11/25",
    sportsbook="Fanatics",
    bet_type="Moneyline",
    selection="Kansas City Chiefs ML vs LA Chargers",
    odds=-170,
    stake=10,
    result="Loss",
    bonus=False
)

# 2. Hard Rock $5 ML win
log_bet(
    date="9/11/25",
    sportsbook="Hard Rock",
    bet_type="Moneyline",
    selection="Reds vs Padraes",
    odds=-950,
    stake=5,
    result="Win",
    bonus=False
)

# 3. Hard Rock $80 EVEN loss (+125)
log_bet(
    date="9/11/25",
    sportsbook="Hard Rock",
    bet_type="Total Points Odd/Even",
    selection="Washington Commanders vs Green Bay Packers - Even",
    odds=125,
    stake=80,
    result="Loss",
    bonus=False
)

# 4. Fanatics $100 bonus bet on Total Points ODD (-125) -> won $80
log_bet(
    date="9/11/25",
    sportsbook="Fanatics",
    bet_type="Total Points Odd/Even",
    selection="Washington Commanders vs Green Bay Packers - Odd",
    odds=-125,
    stake=80,  # real money won from bonus
    result="Win",
    bonus=True
)

# 5. Hard Rock Bets $25 bonus bet on 4-leg parlay (+364) -> potential win of 90.9
log_bet(
    date="9/12/25",
    sportsbook="Hard Rock",
    bet_type="4-leg parlay, Bills, Ravens, Bangels, Chiefs = win",
    selection="Bills, Ravens, Bangels, Chiefs = win",
    odds=+364,
    stake=25,  # real money won from bonus
    result="",
    bonus=True
)

# 6. Hard Rock Bets $25 bonus bet on 4-leg parlay (+364) -> potential win of 90.9
log_bet(
    date="9/12/25",
    sportsbook="Hard Rock",
    bet_type="Point-Spread",
    selection="Falcons -5.5",
    odds=+375,
    stake=25,  # real money won from bonus
    result="",
    bonus=True
)