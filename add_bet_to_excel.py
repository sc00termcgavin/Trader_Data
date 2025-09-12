# 1. Automatically treat bonus bets as $0 cost.
# 2.	Prefill Decimal Odds from the American odds you enter.
# 3.	Append new rows to your Bet Log.
# 4.	Excel formulas handle Payout, Net PnL, and Cumulative PnL automatically.

import openpyxl

FILE_PATH = "Simple_Bet_Tracker.xlsx"

def american_to_decimal(odds):
    """Convert American odds to Decimal Odds"""
    if odds > 0:
        return 1 + odds / 100
    else:
        return 1 + 100 / abs(odds)

def add_bet_to_excel(date, sportsbook, bet_type, selection, odds, stake=0, payout=None, result="Pending", bonus=False):
    """
    Append a new bet to Simple_Bet_Tracker.xlsx
    
    Parameters:
    - date: string "MM/DD/YY" or datetime
    - sportsbook: string
    - bet_type: string
    - selection: string
    - odds: int (American odds, e.g., +375, -120)
    - stake: float, default 0 if bonus=True
    - payout: float, optional
    - result: "Win", "Loss", or "Pending"
    - bonus: bool, if True stake is treated as 0
    """

    # Adjust stake for bonus bets
    if bonus:
        stake = 0

    decimal_odds = american_to_decimal(odds)

    # Open workbook
    wb = openpyxl.load_workbook(FILE_PATH)
    ws = wb["Bet Log"]

    # Find next empty row
    next_row = ws.max_row + 1

    # Write data
    ws.cell(row=next_row, column=1).value = date
    ws.cell(row=next_row, column=2).value = sportsbook
    ws.cell(row=next_row, column=3).value = stake
    ws.cell(row=next_row, column=4).value = odds
    ws.cell(row=next_row, column=5).value = decimal_odds
    ws.cell(row=next_row, column=6).value = result

    # Payout formula (Excel handles automatically)
    ws.cell(row=next_row, column=7).value = f'=IF(F{next_row}="Win", C{next_row}*E{next_row}, 0)'

    # Net PnL formula
    ws.cell(row=next_row, column=8).value = f'=G{next_row}-C{next_row}'

    wb.save(FILE_PATH)
    print(f"✅ Added bet: {selection} ({sportsbook})")