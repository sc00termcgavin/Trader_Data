# Sports Bet Tracking

This is a project that provides a python-based betting tracker and dashboard for personal use. It allows you to log bets, track Net PnL, visualize cumulative PnL over time, and analyze performance by sportsbook.

1. `dashboard.py`
   1. Main dashboard generator: reads Bet Log Excel, writes formulas, creates cumulative PnL and charts.

2. `add_bet_to_excel`
   1. Helper function to append new bets to Excel, with bonus bet logic and decimal odds calculation.

3. `log_new_bets.py`
   1. Example script showing how to call add_bet_to_excel to add bets.

- Excel formulas automatically recalc Payout, Net PnL, and Cumulative PnL.
- Charts in the Dashboard sheet auto-update when you modify results.

---

## 1. Dependencies

Install the required Python packages using pip:

```bash
pip install -r requirements.txt
```

## 2. Running the Dashboard

The `dashboard.py` script generates your Excel dashboard with:

- Notes:
  - Cumulative Net PnL over time  
  - Net PnL by sportsbook  
  - KPIs: Total PnL, Total Stake, Win %, ROI  

To run it, open a terminal in your project folder and execute:

```bash
python dashboard.py
```

## 3. Logging New bets

The `add_bet_to_excel.py` script appends new bets to the Excel Bet Log without manually Editing Excel

- Notes:
  - bonus=True → stake treated as $0
  - bonus=False → stake is your actual cash amount
  - After entering "Win" or "Loss" in Excel, PnL and dashboard charts update automatically.

