# 📊 Sports Bet Tracker

A Python-based betting tracker and dashboard that logs bets, tracks Net PnL, and visualizes your performance over time.  
You can use it in two ways:
- **Command-line scripts** (`log_new_bets.py` & `dashboard.py`)  
- **Streamlit web app** (`app.py`) for a self-hosted dashboard with an interactive form

---

## ✨ Features

- Log bets with metadata: **Date, Sportsbook, League, Market, Pick, Odds, Stake, Result, Bonus**
- Automatically calculates:
  - **Payout ($)**
  - **Net PnL ($)**
  - **Cumulative PnL ($)**
- Track performance with KPIs:
  - **Total PnL ($)**
  - **Total Stake ($)**
  - **Wins / Total Bets / Pending Bets**
  - **Win %**
  - **ROI (%)**
- Excel Dashboard (`dashboard.py`) generates:
  - Line chart: Cumulative Net PnL over time
  - Bar chart: Net PnL by Sportsbook
- Web App (`app.py`) built with Streamlit:
  - Sidebar form to log bets (NFL, NBA, UFC, etc.)
  - Real-time table of bets
  - KPIs updated instantly

---

## ⚙️ Installation

Clone the repo and install dependencies:

```bash
git clone <your-repo-url>
cd Trader_Data
python -m venv .venv
source .venv/bin/activate   # On Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

---

1. Log Bets via CLI
```bash
python log_new_bets.py
```
> Interactive prompts let you enter bets one at a time.

2. Generate Excel Dashboard
```bash
python dashboard.py
```
> Creates/updates Bet_Tracker.xlsx with KPIs and charts.

3. Run the Web App
```bash
streamlit run app.py
```
> Open http://localhost:8501 to access:
> * A form to log new bets
> * A table of past bets
> * KPIs: Total PnL, Total Stake, Win %, ROI %


---

## ToDO

- [ ]Deployable Docker image for easy hosting