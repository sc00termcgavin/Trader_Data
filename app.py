import streamlit as st
from datetime import datetime
import openpyxl
from log_new_bets import log_bet  # reuse your existing function

FILE_PATH = "Bet_Tracker.xlsx"

st.set_page_config(page_title="📊 Bet Tracker", layout="wide")

st.title("📊 Bet Tracker Dashboard")

# -------------------------------
# Sidebar - Log New Bet
# -------------------------------
st.sidebar.header("➕ Log a New Bet")

with st.sidebar.form("log_bet_form"):
    sportsbook = st.text_input("Sportsbook")

    # League - required dropdown
    league = st.selectbox(
        "League",
        ["Select...", "NFL", "NBA", "MLB", "NHL", "EPL", "UFC", "Other"],
        index=0
    )

    # Market - required dropdown
    market = st.selectbox(
        "Market",
        ["Select...", "Moneyline", "Spread", "Total", "Prop", "Parlay"],
        index=0
    )

    pick = st.text_input("Pick / Wager")
    odds = st.number_input("Odds (American)", step=1, value=-110)
    stake = st.number_input("Stake ($)", step=1.0, value=10.0)
    result = st.selectbox("Result", ["", "Win", "Loss", "Push"])
    bonus = st.checkbox("Bonus Bet?")
    date = st.date_input("Date", datetime.today())

    submitted = st.form_submit_button("Log Bet")

    if submitted:
        # Validation
        if not sportsbook:
            st.error("❌ Sportsbook is required.")
        elif league == "Select...":
            st.error("❌ Please select a League.")
        elif market == "Select...":
            st.error("❌ Please select a Market.")
        elif not pick:
            st.error("❌ Pick / Wager is required.")
        else:
            log_bet(
                date.strftime("%m/%d/%y"),
                sportsbook,
                league,
                market,
                pick,
                odds,
                stake,
                result,
                bonus
            )
            st.success(f"✅ Logged bet: {league} - {market} - {pick} ({sportsbook})")

# -------------------------------
# Main Dashboard
# -------------------------------
try:
    wb = openpyxl.load_workbook(FILE_PATH, data_only=True)
    ws_log = wb["Bet Log"]

    st.subheader("📑 Bet Log")

    header_row = next(ws_log.iter_rows(min_row=1, max_row=1, values_only=True), None)
    headers = []
    if header_row:
        for idx, value in enumerate(header_row):
            headers.append(value if value not in (None, "") else f"Column {idx + 1}")

    max_col = len(headers) if headers else ws_log.max_column
    data_rows = list(ws_log.iter_rows(min_row=2, max_col=max_col, values_only=True))
    table_data = []

    if headers:
        for row in data_rows:
            if not any(cell not in (None, "") for cell in row):
                continue
            row_dict = {
                headers[idx]: (row[idx] if idx < len(row) else None)
                for idx in range(len(headers))
            }
            table_data.append(row_dict)

    if table_data:
        st.dataframe(table_data, width='stretch')
    else:
        st.info("No bets logged yet. Add a wager using the sidebar form.")

    # # Show KPIs
    # ws_dash = wb["Dashboard"]
    # st.subheader("📈 KPIs")
    # kpis = {}
    # for row in ws_dash.iter_rows(min_row=2, max_col=2, values_only=True):
    #     if row[0]:
    #         kpis[row[0]] = row[1]

    # col1, col2, col3, col4 = st.columns(4)
    # with col1: st.metric("Total PnL ($)", kpis.get("Total PnL ($)", 0))
    # with col2: st.metric("Total Stake ($)", kpis.get("Total Stake ($)", 0))

    # win_pct = kpis.get("Win %")
    # roi_pct = kpis.get("ROI (%)")

    # with col3: st.metric("Win %", f"{round((win_pct or 0) * 100, 2)}%")
    # with col4: st.metric("ROI (%)", f"{round((roi_pct or 0) * 100, 2)}%")

    # ---- KPIs computed directly from Bet Log (no Dashboard dependency) ----
    def _to_number(x):
        try:
            return float(x)
        except (TypeError, ValueError):
            return 0.0

    st.subheader("📈 KPIs")

    # Build safe list of dict rows keyed by header names (already done above as table_data)
    rows = table_data  # alias for readability

    total_pnl = sum(_to_number(r.get("Net PnL ($)")) for r in rows)
    total_stake = sum(_to_number(r.get("Stake ($)")) for r in rows)

    result_vals = [(r.get("Result") or "").strip() for r in rows]
    wins = sum(1 for v in result_vals if v == "Win")
    total_bets = sum(1 for v in result_vals if v not in ("", None))
    pending_bets = sum(1 for v in result_vals if v in ("", None))

    win_pct = (wins / total_bets) if total_bets > 0 else 0.0
    roi_pct = (total_pnl / total_stake) if total_stake > 0 else 0.0

    col1, col2, col3, col4 = st.columns(4)
    with col1: st.metric("Total PnL ($)", f"{total_pnl:,.2f}")
    with col2: st.metric("Total Stake ($)", f"{total_stake:,.2f}")
    with col3: st.metric("Win %", f"{win_pct*100:.2f}%")
    with col4: st.metric("ROI (%)", f"{roi_pct*100:.2f}%")

    wb.close()

except FileNotFoundError:
    st.warning("⚠️ No Bet Tracker file found yet. Log your first bet to create one.")
