from pptx import Presentation, util, text
from pptx.util import Inches, Pt
from tkinter import filedialog
import matplotlib.pyplot as plt
import pandas as pd
from PIL import Image
import re
from pptx.enum.text import PP_ALIGN
from pptx.enum.text import MSO_ANCHOR

team = ("ORE_BEA")
csv_file =filedialog.askopenfilename()
df = pd.read_csv(csv_file)

#total_runs = df['RunsScored'].sum()
#runs_per_pa = total_runs / 273516
#numerator = (
#    0.689 * df['KorBB'].eq('BB').sum() +
#    0.720 * (df['PitchCall'] == 'HitByPitch').sum() +
#    0.844 * (df['PlayResult'] == 'Single').sum() +
#    1.261 * (df['PlayResult'] == 'Double').sum() +
#    1.601 * (df['PlayResult'] == 'Triple').sum() +
#    2.072 * (df['PlayResult'] == 'HomeRun').sum()
#)
#new_woba = numerator / 273516
#scale = new_woba / runs_per_pa
#print(scale)
#AVG OPS = .785
#AVG SLG = .418
#PA = 273516

# Define constants
W_OBA_WEIGHTS = {'BB': 0.69, 'HBP': 0.72, '1B': 0.89, '2B': 1.27, '3B': 1.62, 'HR': 2.10}
W_OBA_SCALE = 1.264
LEAGUE_WOBA = 0.276
RUNS_PER_WIN = 8.734
LEAGUE_OBP = 0.367  
LEAGUE_SLG = 0.418

# Map PlayResult to hit outcomes
hit_map = {
    'Single': '1B', 'Double': '2B', 'Triple': '3B', 'HomeRun': 'HR'
}
df['HitType'] = df['PlayResult'].map(hit_map)
df['BB'] = df['KorBB'].apply(lambda x: 1 if x == 'Walk' else 0)
df['HBP'] = df['PitchCall'].apply(lambda x: 1 if x == 'HitByPitch' else 0)

# Filter hitters by selected team
team_hitters_df = df[df['BatterTeam'] == team].copy()

# Get total plate appearances
team_hitters_df['PA_event'] = team_hitters_df['PitchCall'].isin([
    'StrikeSwinging', 'StrikeCalled', 'FoulBall', 'InPlay',
    'BallCalled', 'BallIntentional', 'HitByPitch'
]) | team_hitters_df['PlayResult'].notna()
team_hitters_df = team_hitters_df[team_hitters_df['PA_event']]

# Group and count outcomes
grouped = team_hitters_df.groupby('Batter')
results = []

for batter, group in grouped:
    pos = input(f"Enter primary position for {batter}: ").strip()

    # Filter only final pitches that end a plate appearance
    plate_appearances = ((group['PitchCall'] == "InPlay") | (group['KorBB'].isin(['Strikeout', 'Walk'])) | (group['PitchCall'] == 'HitByPitch')).sum()
    
    counts = {
        '1B': (group['HitType'] == '1B').sum(),
        '2B': (group['HitType'] == '2B').sum(),
        '3B': (group['HitType'] == '3B').sum(),
        'HR': (group['HitType'] == 'HR').sum(),
        'BB': group['BB'].sum(),
        'HBP': group['HBP'].sum(),
        'PA': plate_appearances
    }
    
    # Add calculated components
    hits = counts['1B'] + counts['2B'] + counts['3B'] + counts['HR']
    total_bases = counts['1B'] + 2 * counts['2B'] + 3 * counts['3B'] + 4 * counts['HR']
    ab = counts['PA'] - counts['BB'] - counts['HBP']  # Subtract SF too if you have it

    # Calculate SLG, OBP, OPS
    slg = total_bases / ab if ab > 0 else 0
    obp = (hits + counts['BB'] + counts['HBP']) / (ab + counts['BB'] + counts['HBP']) if (ab + counts['BB'] + counts['HBP']) > 0 else 0
    ops = slg + obp



    pos_adj_table = {
        'C': 7.5, 'SS': 7.5, '2B': 5, '3B': 2.5, 'CF': 2.5,
        'LF': -2.5, 'RF': -2.5, '1B': -5, 'DH': -6
    }
    pos_adj = pos_adj_table.get(pos.upper(), 0) * (counts['PA'] / 300)
    num = sum(counts[k] * W_OBA_WEIGHTS[k] for k in W_OBA_WEIGHTS)
    denom = counts['PA']
    woba = num / denom if denom > 0 else 0
    wraa = (woba - LEAGUE_WOBA) / W_OBA_SCALE * counts['PA']
    repl = 20 * (counts['PA'] / 250)
    war = (wraa + repl + pos_adj) / RUNS_PER_WIN

    ops_plus = 100 * ((obp / LEAGUE_OBP) + ((slg / LEAGUE_SLG) - 1)) if LEAGUE_OBP > 0 and LEAGUE_SLG > 0 else 100


    total_swings = ((group['PitchCall'] == 'InPlay') | 
                (group['PitchCall'] == 'FoulBall') | 
                (group['PitchCall'] == 'FoulBallFieldable') | 
                (group['PitchCall'] == 'FoulBallNotFieldable') | 
                (group['PitchCall'] == 'StrikeSwinging')).sum()

    # Example: mark all pitches in Zones 1–9 as strikes
    good_swings = ((group['PlateLocHeight'] >= 1.524166) & 
                (group['PlateLocHeight'] <= 3.3775) &
                (group['PlateLocSide'] >= -0.830833) & 
                (group['PlateLocSide'] <= 0.830833) &
                ((group['PitchCall'] == 'InPlay') | 
                (group['PitchCall'] == 'FoulBall') | 
                (group['PitchCall'] == 'StrikeSwinging'))).sum()
    
    seager_approx = good_swings / total_swings if total_swings > 0 else 0



    results.append({
        'Batter': batter,
        'Position': pos,
        'PA': counts['PA'],
        'SLG': round(slg, 3),
        'OBP': round(obp, 3),
        'OPS': round(ops, 3),
        'wOBA': round(woba, 3),
        'wRAA': round(wraa, 2),
        'Replacement Runs': round(repl, 2),
        'Position Adjusment': round(pos_adj, 2),
        'SEAGER': round(seager_approx, 3),
        'OPS+': round(ops_plus, 1),
        'WAR': round(war, 2)
    })

# Convert to DataFrame and display
war_df = pd.DataFrame(results)
war_df = war_df.sort_values(by='WAR', ascending=False)
print(war_df)

#PAs in 99, Runs per win, woba scale