import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import CellIsRule
from datetime import datetime
import os

FILE_PATH = "Bet_Tracker.xlsx"

HEADERS = [
    "Date",
    "Sportsbook",
    "League",
    "Market",
    "Pick",
    "Stake ($)",
    "Odds",
    "Result",
    "Bonus",
    "Decimal Odds",
    "Payout ($)",
    "Net PnL ($)",
    "Cumulative PnL ($)"
]

# -------------------------------
# Utility: Convert American odds to Decimal
# -------------------------------
def american_to_decimal(odds):
    if odds == 0 or odds is None:
        return 0
    return 1 + (odds / 100 if odds > 0 else 100 / abs(odds))

# -------------------------------
# Header maintenance
# -------------------------------
def ensure_bet_log_headers(ws):
    header_values = [cell.value for cell in ws[1]] if ws.max_row >= 1 else []

    if not any(header_values):
        ws.append(HEADERS)
        return

    rename_map = {"Bet Type": "Market", "Selection": "Pick"}
    for idx, value in enumerate(header_values, start=1):
        if value in rename_map:
            ws.cell(row=1, column=idx, value=rename_map[value])

    header_values = [cell.value for cell in ws[1]]

    if "League" not in header_values:
        try:
            sportsbook_idx = header_values.index("Sportsbook") + 1
        except ValueError:
            sportsbook_idx = 2
        ws.insert_cols(sportsbook_idx + 1)
        ws.cell(row=1, column=sportsbook_idx + 1, value="League")

    for idx, header in enumerate(HEADERS, start=1):
        ws.cell(row=1, column=idx, value=header)


def log_bet(date, sportsbook, league, market, pick, odds, stake=0, result="", bonus=False):
    """
    Append a single bet to Bet Tracker.
    Payout, Net PnL, and Cumulative PnL are calculated in Python and stored as numeric values.
    """
    # Load workbook or create if not exist
    try:
        wb = openpyxl.load_workbook(FILE_PATH)
        ws = wb["Bet Log"]
    except FileNotFoundError:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Bet Log"
        ws.append(HEADERS)
    else:
        ensure_bet_log_headers(ws)

    next_row = ws.max_row + 1

    # Actual stake for bonus bets
    actual_stake = 0 if bonus else stake
    dec_odds = american_to_decimal(odds)

    # -------------------------------
    # Compute values directly in Python
    # -------------------------------
    payout = None
    net_pnl = None
    cumulative_pnl = None

    result_clean = (result or "").strip()

    if result_clean and result_clean != "Open":
        if bonus:
            if result_clean == "Win":
                payout = stake * (dec_odds - 1)
            elif result_clean == "Push":
                payout = actual_stake
            else:  # Loss
                payout = 0
        else:
            if result_clean == "Win":
                payout = actual_stake * dec_odds
            elif result_clean == "Push":
                payout = actual_stake
            else:  # Loss
                payout = 0

        net_pnl = payout - actual_stake
        # Previous cumulative PnL + current net
        if next_row > 2 and ws[f"M{next_row-1}"].value not in (None, ""):
            cumulative_pnl = ws[f"M{next_row-1}"].value + net_pnl
        else:
            cumulative_pnl = net_pnl

    # -------------------------------
    # Append data into worksheet
    # -------------------------------
    ws[f"A{next_row}"] = date
    ws[f"B{next_row}"] = sportsbook
    ws[f"C{next_row}"] = league
    ws[f"D{next_row}"] = market
    ws[f"E{next_row}"] = pick
    ws[f"F{next_row}"] = actual_stake
    ws[f"G{next_row}"] = odds
    ws[f"H{next_row}"] = result_clean
    ws[f"I{next_row}"] = bonus
    ws[f"J{next_row}"] = dec_odds
    ws[f"K{next_row}"] = payout
    ws[f"L{next_row}"] = net_pnl
    ws[f"M{next_row}"] = cumulative_pnl

    # Conditional formatting for Net PnL
    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    ws.conditional_formatting.add(f"L2:L{ws.max_row}",
                                  CellIsRule(operator='greaterThan', formula=['0'], fill=green_fill))
    ws.conditional_formatting.add(f"L2:L{ws.max_row}",
                                  CellIsRule(operator='lessThan', formula=['0'], fill=red_fill))

    # -------------------------------
    # Dashboard - update KPIs
    # -------------------------------
    if "Dashboard" in wb.sheetnames:
        ws_dash = wb["Dashboard"]
        for row in ws_dash.iter_rows(min_row=2, max_col=2, max_row=20):
            for cell in row:
                cell.value = None
    else:
        ws_dash = wb.create_sheet("Dashboard")
        ws_dash["A1"], ws_dash["B1"] = "Metric", "Value"

    ws_dash["A2"], ws_dash["B2"] = "Total PnL ($)", f"=SUM('Bet Log'!L2:L{ws.max_row})"
    ws_dash["A3"], ws_dash["B3"] = "Total Stake ($)", f"=SUM('Bet Log'!F2:F{ws.max_row})"
    ws_dash["A4"], ws_dash["B4"] = "Wins", f'=COUNTIF(\'Bet Log\'!H2:H{ws.max_row},"Win")'
    ws_dash["A5"], ws_dash["B5"] = "Total Bets", f'=COUNTA(\'Bet Log\'!H2:H{ws.max_row})'
    ws_dash["A6"], ws_dash["B6"] = "Pending Bets", f'=COUNTIF(\'Bet Log\'!H2:H{ws.max_row},"")'
    ws_dash["A7"], ws_dash["B7"] = "Win %", f"=IF(B5=0,0,B4/B5)"
    ws_dash["A8"], ws_dash["B8"] = "ROI (%)", f"=IF(B3=0,0,B2/B3)"

    wb.save(FILE_PATH)
    wb.close()
    

    print(f"✅ Logged bet in row {next_row}: {league} {market} - {pick} ({sportsbook})")


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

        league = input("League (e.g., NFL, NBA): ").strip()
        market = input("Market: ").strip()
        pick = input("Pick: ").strip()
        odds = int(input("Odds (e.g., -110, +125): ").strip())
        stake = float(input("Stake ($): ").strip())
        result = input("Result (Win/Loss/Push/blank): ").strip()
        bonus = input("Bonus bet? (y/n): ").strip().lower() == "y"

        log_bet(date, sportsbook, league, market, pick, odds, stake, result, bonus)

    print("✅ All bets logged successfully!")
