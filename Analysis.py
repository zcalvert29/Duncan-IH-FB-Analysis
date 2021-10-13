import pandas as pd

def run_analysis(input_file, other_team, output_file):
    game = pd.read_excel(input_file)[['O/D', 'Down', 'Distance', 'Yard Line', 'R/P/S', 'QB',
                                'Ball Carrier', 'Result', 'Description']]
    # Isolate offensive plays
    offense = game[(game['O/D'] == 'O') & (game['R/P/S'] != 'S')]
    # Read in Player Roster and merge to offensive plays
    roster = pd.read_excel("Roster.xlsx")
    offense = pd.merge(offense, roster, left_on="QB", right_on="Jersey").rename(columns={'Player': 'QB Name'}).drop(
        'Jersey', axis=1)
    offense = pd.merge(offense, roster, left_on="Ball Carrier", right_on="Jersey").rename(
        columns={'Player': 'Ball Carrier Name'}).drop('Jersey', axis=1)
    # Isolate pass plays and rush plays
    passes = offense[offense['R/P/S'] == 'P']
    rushes = offense[offense['R/P/S'] == 'R']
    # Isolate defensive plays
    defense = game[(game['O/D'] == 'D') & (game['R/P/S'] != 'S')]
    # Calculate total yards for O and D and total plays
    total_yards_offense = offense['Result'].sum()
    total_yards_defense = defense['Result'].sum()
    offense_plays = len(offense)
    defense_plays = len(defense)
    # How many completions did our offense have?
    completions = passes[passes['Result'] != 0]
    # Passing and Rushing Yards/Plays
    passing_yards = passes['Result'].sum()
    pass_plays = len(passes)
    rushing_yards = rushes['Result'].sum()
    rush_plays = len(rushes)

    print(f"Duncan's Total Yards of Offense: {total_yards_offense}")
    print(f"# of Plays Run by Duncan: {offense_plays}")
    print(f"Duncan Offensive Efficiency: {round(total_yards_offense / offense_plays, 2)} Yards Per Play\n")
    print(f"{other_team}'s Total Yards of Offense: {total_yards_defense}")
    print(f"# of Plays Run by {other_team}: {defense_plays}")
    print(f"{other_team}'s Offensive Efficiency: {round(total_yards_defense / defense_plays, 2)} Yards Per Play")

    print("\nDuncan Hall's In-Depth Offensive Stats\n")
    print(f"Total Passing Yards: {passing_yards}")
    print(f"Passing Plays: {pass_plays}")
    print(f"Pass Efficiency: {round(passing_yards / pass_plays, 2)} Yards Per Play\n")
    print(f"Total Rushing Yards: {rushing_yards}")
    print(f"Rushing Plays: {rush_plays}")
    print(f"Rush Efficiency: {round(rushing_yards / rush_plays, 2)} Yards Per Play")

    qb = passes.groupby("QB Name")['Result'].sum().reset_index(name="Passing Yards")
    qb['Completions'] = len(completions)
    qb['Attempts'] = len(passes)

    receiving_yards = passes.groupby("Ball Carrier Name")['Result'].sum().sort_values(ascending=False).reset_index(
        name="Receiving Yards")
    receiver_targets = passes.groupby('Ball Carrier Name').size().sort_values(ascending=False).reset_index(
        name="Targets")
    receiver_catches = completions.groupby('Ball Carrier Name').size().sort_values(ascending=False).reset_index(
        name="Catches")
    receiver = receiver_catches.merge(receiving_yards, on="Ball Carrier Name")
    receiver = receiver.merge(receiver_targets, on="Ball Carrier Name")

    rusher_yards = rushes.groupby('Ball Carrier Name')['Result'].sum().sort_values(ascending=False).reset_index(
        name="Rushing Yards")
    rusher_carries = rushes.groupby('Ball Carrier Name').size().sort_values(ascending=False).reset_index(name="Carries")
    rush = rusher_carries.merge(rusher_yards, on="Ball Carrier Name")

    writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
    qb.to_excel(writer, sheet_name='QB')
    rush.to_excel(writer, sheet_name='RB')
    receiver.to_excel(writer, sheet_name="WR")
    offense.to_excel(writer, sheet_name="Offensive Plays")
    defense.to_excel(writer, sheet_name="Defensive Plays")
    writer.save()
    writer.close()