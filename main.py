from os import listdir
from os.path import isfile, join
import pandas
from IPython.display import display
import functools

PLAY_MONTHS = ['Januar', 'Februar', 'Marts', 'April', 'Maj', 'Juni', 'Juli', 'August', 'Oktober', 'November']

def get_xslx_files():
    files = []
    for file in listdir('./'):
        if file.endswith('.xlsx') and 'Stilling' in file:
            files.append('./'+file)
    return files

def load_xslx_as_pandas(files):
    master_frame = pandas.DataFrame([], columns=["Holdnavn", "Måned", "Point", "Værtshus"])
    for file in files:
        bar_name = file.replace('./Stilling-', '').replace('.xlsx', '')
        sub_frame = pandas.read_excel(file)
        for index, row in sub_frame.iterrows():
            for month in PLAY_MONTHS:
                points = row[month]
                if pandas.notna(points):
                    master_frame.loc[len(master_frame.index)] = [row['Holdnavn'], month, points, bar_name]
    return master_frame

def raise_duplicate_errors(df):
    duplicate_series = df.duplicated(subset=['Holdnavn', 'Måned'], keep=False)
    error = False
    for index, value in duplicate_series.items():
        if value:
            print(f'Holdet {df.iloc[index][0]} har fået point fra flere værtshuse i {df.iloc[index][1]}, {df.iloc[index][2]} point fra {df.iloc[index][3]}')
            error = True
    if error:
        print('Da der er fejl kan dokumentet ikke genereres')
        return
    else:
        return df

def find_qualifiers(df):
    team_names = df["Holdnavn"].unique()
    top_10 = find_top_10(df, team_names)
    taken = top_10['Holdnavn'].tolist()
    top_5 = find_top_5(df, team_names, taken)


def find_top_10(df, team_names):
    top_10_df = pandas.DataFrame([], columns=["Holdnavn", "Total"])
    for team_name in team_names:
        team_score_df = df.loc[df['Holdnavn'] == team_name]
        team_score = team_score_df['Point'].sum()
        top_10_df.loc[len(top_10_df.index)] = [team_name, team_score]
    return top_10_df.sort_values(by=['Total'], ascending=False, ignore_index=True).head(10)

def find_top_5(df, team_names, taken):
    top_5 = pandas.DataFrame([], columns=["Holdnavn", "Værtshus", "Værtshus Total"])
    bar_names = df["Værtshus"].unique()
    eligible_teams = list(filter(lambda x: not x in taken, team_names))
    print("team_names", team_names, "\ntaken", taken, "\neligible_teams", eligible_teams)
    for bar_name in bar_names:
        for team_name in eligible_teams:
            pass

    

files = get_xslx_files()
raw_frame = load_xslx_as_pandas(files)
sanitized_frame = raise_duplicate_errors(raw_frame)
res = find_qualifiers(sanitized_frame)

