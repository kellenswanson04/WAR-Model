from tkinter import filedialog
import pandas as pd

# Load CSV
team = "LEA_AVG"
csv_file = filedialog.askopenfilename()
df = pd.read_csv(csv_file)

# Filter to team pitchers
df = df[df['PitcherTeam'] == team]

# Clean or prepare data
df = df[df['TaggedPitchType'].notna()]  # Remove rows without valid pitches

# Optional: normalize PitchCall names
df['PitchCall'] = df['PitchCall'].fillna("Undefined")

# Keep only pitch calls relevant to BB/K/IP
valid_calls = ['BallCalled', 'StrikeCalled', 'StrikeSwinging', 'FoulBall', 'InPlay', 'HitByPitch']
df = df[df['PitchCall'].isin(valid_calls)]

hit_results = ['Single', 'Double', 'Triple', 'HomeRun']
df['isHit'] = df['PlayResult'].isin(hit_results)

# Tag outcomes
df['isK'] = df['KorBB'] == "Strikeout"
df['isBB'] = df['KorBB'] == "Walk"
df['isIP'] = df['PitchCall'] == "InPlay"

# Estimate innings pitched (3 batters per inning from InPlay)
def estimate_ip(group):
    in_play = group['PitchCall'] == 'InPlay'
    hbp = group['PitchCall'] == 'HitByPitch'
    strikeout = group['KorBB'] == 'Strikeout'
    walk = group['KorBB'] == 'Walk'
    
    batters_faced = (in_play | hbp | strikeout | walk).sum()
    return batters_faced / 3.0

df['isHR'] = df['PlayResult'] == 'HomeRun'

def to_baseball_innings(ip):
    full = int(ip)
    frac = ip - full
    if frac < 1/3:
        return full + 0.0
    elif frac < 2/3:
        return full + 0.1
    else:
        return full + 0.2

# Aggregate by pitcher
pitcher_stats = df.groupby('Pitcher').apply(lambda g: pd.Series({
    'IP': estimate_ip(g),
    'K': g['isK'].sum(),
    'BB': g['isBB'].sum(),
    'HR': g['isHR'].sum(),
    'H': g['isHit'].sum(),
    'TotalPitches': g.shape[0]
})).reset_index()

# FIP and WAR estimation
FIP_constant = 3.75
replacement_fip = 7.44

pitcher_stats['WHIP'] = (pitcher_stats['BB'] + pitcher_stats['H']) / pitcher_stats['IP']
pitcher_stats['FIP'] = (13 * pitcher_stats['HR']+ 3 * pitcher_stats['BB'] - 2 * pitcher_stats['K']) / pitcher_stats['IP'] + FIP_constant
pitcher_stats['Runs_Prevented'] = pitcher_stats['IP'] * (replacement_fip - pitcher_stats['FIP']) / 9
pitcher_stats['WAR'] = pitcher_stats['Runs_Prevented'] / 9

# Sort by WAR
pitcher_stats['IP'] = pitcher_stats['IP'].apply(to_baseball_innings)
pitcher_stats['K'] = pitcher_stats['K'].astype(int)
pitcher_stats['BB'] = pitcher_stats['BB'].astype(int)
pitcher_stats['HR'] = pitcher_stats['HR'].astype(int)
pitcher_stats['FIP'] = pitcher_stats['FIP'].round(2)
pitcher_stats['WHIP'] = pitcher_stats['WHIP'].round(2)
pitcher_stats['WAR'] = pitcher_stats['WAR'].round(2)
pitcher_stats = pitcher_stats.sort_values(by='WAR', ascending=False)

# Display final stats
print(pitcher_stats[['Pitcher', 'IP', 'K', 'BB', 'HR', 'FIP', 'WHIP', 'WAR']])