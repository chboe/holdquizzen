from tkinter import messagebox
from os import listdir
from os.path import isfile, join
import pandas
from IPython.display import display
import functools
import numpy as np

PLAY_MONTHS = ['Januar', 'Februar', 'Marts', 'April', 'Maj', 'Juni', 'Juli', 'August', 'Oktober', 'November']

def get_xslx_files():
    files = []
    for file in listdir('./'):
        if file.endswith('.xlsx') and 'Stilling' in file:
            files.append('./'+file)
    return files

def load_xslx_as_pandas(files):
    master_frame = pandas.DataFrame([], columns=["Holdnavn", "Måned", "Point", "Værtshus"])
    error = False
    error_msg = []
    for file in files:
        bar_name = file.replace('./Stilling-', '').replace('.xlsx', '')
        try:
            with open(file, 'rb') as f:
                sub_frame = pandas.read_excel(f)
                for index, row in sub_frame.iterrows():
                    for month in PLAY_MONTHS:
                        points = row[month]
                        if pandas.notna(points):
                            master_frame.loc[len(master_frame.index)] = [row['Holdnavn'], month, points, bar_name]
        except IOError as e:
            return master_frame, error, error_msg
    return master_frame, error, error_msg

def raise_duplicate_errors(df):
    duplicate_series = df.duplicated(subset=['Holdnavn', 'Måned'], keep=False)
    error = False
    error_msg = []
    for index, value in duplicate_series.items():
        if value:
            error_msg.append(f'Holdet {df.iloc[index][0]} har fået point fra flere værtshuse i {df.iloc[index][1]}, {df.iloc[index][2]} point fra {df.iloc[index][3]}')
            error = True
    return df, error, error_msg

def find_qualifiers(df):
    team_names = df["Holdnavn"].unique()
    overall_top_teams = find_top_10_teams(df, team_names)
    taken = overall_top_teams['Holdnavn'].tolist()
    bar_totals = find_bar_totals(df, team_names, taken)
    with pandas.ExcelWriter('./Resultater.xlsx') as writer:
        overall_top_teams.index = np.arange(1, len(overall_top_teams) + 1)
        overall_top_teams.to_excel(writer, sheet_name="Top hold")
        for tuple in bar_totals:
            bar_name, bar_scores, bar_index = tuple[0], tuple[1], tuple[2]
            bar_scores.index = np.arange(1, len(bar_scores) + 1)
            bar_scores.to_excel(writer, sheet_name=bar_name)

def find_top_10_teams(df, team_names):
    top_df = pandas.DataFrame([], columns=["Holdnavn", "Total", "Fejl"])
    for team_name in team_names:
        team_score_df = df.loc[df['Holdnavn'] == team_name]
        team_score = team_score_df['Point'].sum()
        top_df.loc[len(top_df.index)] = [team_name, team_score, ""]
    top_df = top_df.sort_values(by=['Total'], ascending=False, ignore_index=True)
    top_df, index = resolve_equal_score_error(top_df, 'Total', 9)
    return top_df.head(index)

def resolve_equal_score_error(df, key, index):
    cut_off_point = df.loc[index][key]
    equal_score_teams_error = df.loc[df[key] == cut_off_point]['Holdnavn']
    if df.iloc[index+1][key] == cut_off_point:
        if len(equal_score_teams_error) > 1:
            for team_name in equal_score_teams_error:
                df.loc[df['Holdnavn'] == team_name, 'Fejl'] = f'Der er delt {index+1}. plads for hold med {cut_off_point} point'
    i = index
    while df.iloc[i][key] == cut_off_point:
        i += 1
    return df, i

def resolve_multi_qualified_teams(res):
    master_frame = pandas.DataFrame([], columns=["Holdnavn", "Værtshuse"])
    for bar_name, df, index in res:
        for team_name in df.iloc[:index]['Holdnavn'].values.tolist():
            master_frame.loc[len(master_frame.index)] = [team_name, bar_name]

    new_res = []
    fixed_teams = []
    duplicates = master_frame[master_frame.duplicated(subset=['Holdnavn'], keep=False)]
    for team_name in duplicates['Holdnavn'].values.tolist():
        if team_name not in fixed_teams:
            for bar_name, df, index in res:
                duplicate_bar_names = duplicates.loc[duplicates['Holdnavn'] == team_name]['Værtshuse'].values.tolist()
                current_error = df.loc[df['Holdnavn'] == team_name, 'Fejl'].values[0]
                if len(current_error) > 1:
                    df.loc[df['Holdnavn'] == team_name, 'Fejl'] = f'{current_error}, kvalificeret i både {" og ".join(duplicate_bar_names)}'
                else:
                    df.loc[df['Holdnavn'] == team_name, 'Fejl'] = f'Kvalificeret i både {" og ".join(duplicate_bar_names)}'
                fixed_teams.append(team_name)
                new_res.append([bar_name, df, index])
    return new_res

def find_bar_totals(df, team_names, taken):
    total_bar_team_points = pandas.DataFrame([], columns=["Holdnavn", "Værtshus", "Værtshus Total", "Fejl"])
    bar_names = df["Værtshus"].unique()
    eligible_teams = list(filter(lambda x: not x in taken, team_names))
    for bar_name in bar_names:
        for team_name in eligible_teams:
            bar_team_points = df.loc[(df['Holdnavn'] == team_name) & (df['Værtshus'] == bar_name)]
            points = bar_team_points['Point'].sum()
            total_bar_team_points.loc[len(total_bar_team_points.index)] = [team_name, bar_name, points, '']
    res = []
    for bar_name in bar_names:
        bar_df = total_bar_team_points.loc[total_bar_team_points['Værtshus'] == bar_name]
        bar_df = bar_df.sort_values(by=['Værtshus Total'], ascending=False, ignore_index=True).drop(columns='Værtshus')
        bar_df, index = resolve_equal_score_error(bar_df, 'Værtshus Total', 4)
        res.append([bar_name, bar_df, index])

    res = resolve_multi_qualified_teams(res)

    return res

files = get_xslx_files()
raw_frame, error, error_msg = load_xslx_as_pandas(files)
if error:
    msg = "\n".join(error_msg)
    messagebox.showerror(title='Fejl', message=msg)
else:
    sanitized_frame, error, error_msg = raise_duplicate_errors(raw_frame)
    if error:
        msg = "\n".join(error_msg)
        messagebox.showerror(title='Fejl', message=msg+'\nHusk at gemme excel filerne bagefter!')
    else:
        find_qualifiers(sanitized_frame)

