# import openpyxl

# FILE_PATH = "Bet_Tracker.xlsx"

# def american_to_decimal(odds):
#     """Convert American odds to decimal odds."""
#     if odds is None:
#         return 0
#     return 1 + (odds / 100 if odds > 0 else 100 / abs(odds))

# def log_bet(date, sportsbook, bet_type, selection, odds, stake=0, result="", bonus=False):
#     """
#     Log a new bet into Bet_Tracker.xlsx.

#     Bonus bets:
#       - Stake in the sheet remains 0
#       - Payout calculated as stake * (decimal odds - 1)
#       - Net PnL = Payout
#     """
#     # Load workbook
#     wb = openpyxl.load_workbook(FILE_PATH)
#     ws = wb["Bet Log"]

#     next_row = ws.max_row + 1
#     dec_odds = american_to_decimal(odds)

#     # Add main values
#     ws[f"A{next_row}"] = date
#     ws[f"B{next_row}"] = sportsbook
#     ws[f"C{next_row}"] = bet_type
#     ws[f"D{next_row}"] = selection
#     ws[f"E{next_row}"] = 0 if bonus else stake
#     ws[f"F{next_row}"] = odds
#     ws[f"G{next_row}"] = result
#     ws[f"H{next_row}"] = bonus
#     ws[f"I{next_row}"] = dec_odds

#     # Payout formula
#     if bonus:
#         # For bonus bet, use the original stake in formula, not E column
#         ws[f"J{next_row}"] = f'=IF(G{next_row}="Win",{stake}*(I{next_row}-1),0)'
#     else:
#         ws[f"J{next_row}"] = f'=IF(G{next_row}="Win",E{next_row}*I{next_row},0)'

#     # Net PnL
#     ws[f"K{next_row}"] = f'=J{next_row}-E{next_row}'

#     # Cumulative PnL
#     ws[f"L{next_row}"] = f'=SUM(K$2:K{next_row})'

#     wb.save(FILE_PATH)
#     print(f"✅ Added bet to row {next_row}: {selection} ({sportsbook})")


# # -------------------------------
# # Bets
# # -------------------------------

# #1. Fanatics $10 moneyline loss -> gives $100 bonus
# log_bet(
#     date="9/5/25",
#     sportsbook="Fanatics",
#     bet_type="Money Line",
#     selection="Kansas City Chiefs vs LA Chargers",
#     odds=-170,
#     stake=10,
#     result="Loss",
#     bonus=False
# )

# #2. Hard Rock $5 moneyline win on Reds vs Padres (-950)
# log_bet(
#     date="9/11/25",
#     sportsbook="Hard Rock Bet",
#     bet_type="Money Line",
#     selection="Cincinnati Reds vs San Diego Padres",
#     odds=-950,
#     stake=5,
#     result="Win",
#     bonus=False
# )

# # 3. Hard Rock $80 cash bet EVEN on Total Points (lost +125)
# log_bet(
#     date="9/11/25",
#     sportsbook="Hard Rock Bet",
#     bet_type="Total Points Odd/Even",
#     selection="Washington Commanders vs Green Bay Packers - Even",
#     odds=125,
#     stake=80,
#     result="Loss",
#     bonus=False
# )

# # # 4. Fanatics $100 bonus bet on Total Points ODD (-125) -> won $80
# log_bet(
#     date="9/11/25",
#     sportsbook="Fanatics",
#     bet_type="Total Points Odd/Even",
#     selection="Washington Commanders vs Green Bay Packers - Odd",
#     odds=-125,
#     stake=100,
#     result="Win",
#     bonus=True
# )

# ---------------------------------------------------------------------------------
# ---------------------------------------------------------------------------------
# ---------------------------------------------------------------------------------
# ---------------------------------------------------------------------------------
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
    if bonus:
        # For bonus bets, payout = actual profit won (real money), ignore stake
        ws[f"J{next_row}"] = f'=IF(G{next_row}="Win",{stake},0)'
    else:
        ws[f"J{next_row}"] = f'=IF(G{next_row}="Win",E{next_row}*I{next_row},0)'

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
# Example bets
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