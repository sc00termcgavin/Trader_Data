# Sports Bet Tracking

This is a project that provides a python-based betting tracker and dashboard for personal use. It allows you to log bets, track Net PnL, visualize cumulative PnL over time, and analyze performance by sportsbook.

1. `dashboard.py`
   1. Main dashboard generator: reads Bet Log Excel, writes formulas, creates cumulative PnL and charts.

2. `log_new_bets`
   1. Helper function to append new bets (bonus/cash, decimal odds, formulas auto-filled).

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

```bash
python log_new_bets.py
```
