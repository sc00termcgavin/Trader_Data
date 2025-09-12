from add_bet_to_excel import add_bet_to_excel

# For normal cash bets, set bonus=False and provide the stake.
# Example: 4-leg parlay bonus bet
add_bet_to_excel(
    date="9/12/25",
    sportsbook="Hard Rock",
    bet_type="Parlay (4-leg)",
    selection="Bills ML, Ravens ML, Bengals ML, Chiefs ML",
    odds=264,  # Approx American odds from payout
    bonus=True
)

# Example: Falcons -5.5 bonus bet
add_bet_to_excel(
    date="9/12/25",
    sportsbook="Hard Rock",
    bet_type="Point Spread",
    selection="Falcons -5.5 vs Vikings",
    odds=375,
    bonus=True
)

# Example: normal cash bet
add_bet_to_excel(
    date="9/13/25",
    sportsbook="Fanatics",
    bet_type="Moneyline",
    selection="Packers ML",
    odds=-120,
    stake=50,
    bonus=False
)
