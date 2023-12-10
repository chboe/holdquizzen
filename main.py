import os
import sys
from tkinter import messagebox
from os import listdir
import pandas
import numpy as np

PLAY_MONTHS = ['Januar', 'Februar', 'Marts', 'April', 'Maj', 'Juni', 'Juli', 'August', 'September', 'Oktober', 'November']


def get_xslx_files():
    files = []
    for file in listdir('./'):
        if file.endswith('.xlsx') and 'Stilling' in file:
            files.append('./' + file)
    return files


def load_xslx_as_pandas(files):
    score_frame = pandas.DataFrame([], columns=["Holdnavn", "Måned", "Point", "Quizsted"])
    winnings_frame = pandas.DataFrame([], columns=["Holdnavn", "Quizsted", "Samlet Gevinst"])
    error = False
    error_msg = []
    for file in files:
        bar_name = file.replace('./Stilling-', '').replace('.xlsx', '')
        _, file_extension = os.path.splitext(file)
        try:
            with open(file, 'rb') as f:
                if file_extension == '.xlsx':
                    sub_frame = pandas.read_excel(f, engine='openpyxl')
                elif file_extension == '.xls':
                    sub_frame = pandas.read_excel(f, engine='xlrd')
                for index, row in sub_frame.iterrows():
                    winnings_frame.loc[len(winnings_frame.index)] = [row['Holdnavn'], bar_name, row['Samlet Gevinst']]
                    for month in PLAY_MONTHS:
                        points = row[month]
                        if pandas.notna(points):
                            score_frame.loc[len(score_frame.index)] = [row['Holdnavn'], month, points, bar_name]
        except IOError as e:
            return score_frame, error, error_msg
    return score_frame, winnings_frame, error, error_msg


def raise_points_frame_errors(points_frame):
    duplicate_series = points_frame.duplicated(subset=['Holdnavn', 'Måned'], keep=False)
    error = False
    error_msg = []
    for index, value in duplicate_series.items():
        if value:
            error_msg.append(
                f'Holdet {points_frame.iloc[index][0]} har fået point fra flere Quizsteder i {points_frame.iloc[index][1]}, {points_frame.iloc[index][2]} point fra {points_frame.iloc[index][3]}')
            error = True
    return points_frame, error, error_msg

def add_in_top_10(bar_totals, overall_top_teams):
    result_dfs = []
    overall_top_10_points = list(set(overall_top_teams.sort_values(by='Total', ascending=False).head(10)['Total']))
    cut_off_point = max(overall_top_10_points[0], 0)
    yes_teams = overall_top_teams[overall_top_teams['Total'] > cut_off_point]
    maybe_teams = overall_top_teams[overall_top_teams['Total'] == cut_off_point]


    for bar_name, bar_df in bar_totals:
        bar_df.loc[bar_df['Holdnavn'].isin(yes_teams['Holdnavn']), 'Ligger i top 10'] = 'Ja'

        if len(maybe_teams) > 1:
            bar_df.loc[bar_df['Holdnavn'].isin(maybe_teams['Holdnavn']), 'Ligger i top 10'] = 'Måske'
        else:
            bar_df.loc[bar_df['Holdnavn'].isin(maybe_teams['Holdnavn']), 'Ligger i top 10'] = 'Ja'
        result_dfs.append((bar_name, bar_df))

    return result_dfs


def generate_output(points_frame, winnings_frame):
    team_names = points_frame["Holdnavn"].unique()
    overall_top_teams = find_top_teams(points_frame, winnings_frame, team_names)
    bar_totals = find_bar_totals(points_frame, winnings_frame, team_names)
    bar_totals = add_in_top_10(bar_totals, overall_top_teams)
    try:
        with pandas.ExcelWriter('./Resultater.xlsx') as writer:
            overall_top_teams.index = np.arange(1, len(overall_top_teams) + 1)
            overall_top_teams.to_excel(writer, sheet_name="Top hold")
            for tuple in bar_totals:
                bar_name, bar_scores = tuple[0], tuple[1]
                bar_scores.index = np.arange(1, len(bar_scores) + 1)
                bar_scores.to_excel(writer, sheet_name=bar_name)
    except IOError as e:
        messagebox.showerror(title='Note', message='Luk "Resultater.xlsx" før du forsøger at generere resultaterne!')


def find_top_teams(points_frame, winnings_frame, team_names):
    top_df = pandas.DataFrame([], columns=["Holdnavn", "Total", "Højeste Score", "Antal Deltagelser", "Gennemsnit", "Quizsted", "Samlet Gevinst", "Note"])
    for team_name in team_names:
        team_score_df = points_frame.loc[points_frame['Holdnavn'] == team_name]
        team_score = team_score_df['Point'].sum()
        if len(team_score_df['Point'].index) > 0:
            max_score = team_score_df[team_score_df['Point'] == team_score_df['Point'].max()]['Point'].values.tolist()[
                0]
        else:
            max_score = 0
        participations = len(team_score_df.index)
        amount_visited_series = points_frame.loc[points_frame['Holdnavn'] == team_name]['Quizsted'].value_counts().sort_values(ascending=False)
        most_visited = amount_visited_series.where(amount_visited_series == amount_visited_series[0]).dropna().index.values.tolist()
        average = round(max_score/max(participations, 1), 2)
        top_df.loc[len(top_df.index)] = [team_name, team_score, max_score, participations, average, " og ".join(most_visited), winnings_frame.loc[winnings_frame['Holdnavn'] == team_name]["Samlet Gevinst"].sum(), ""]
    top_df = top_df.sort_values(by=['Antal Deltagelser'], ascending=False, ignore_index=True)
    top_df = top_df.sort_values(by=['Højeste Score'], ascending=False, ignore_index=True)
    top_df = top_df.sort_values(by=['Total'], ascending=False, ignore_index=True)
    top_df = resolve_equal_score_error(top_df, 'Total')
    return top_df


def resolve_equal_score_error(df, key):
    df = df.sort_values(by=[key], ascending=False, ignore_index=True)
    df = df.reset_index(drop=True)
    points = df[key].values.tolist()
    for point in points:
        same_points = df.loc[df[key] == point]
        if len(same_points.index) > 1:
            note = f'Der er delt {same_points.index[0] + 1}-{same_points.index[-1] + 1}. plads for hold med {int(point)} point'
            df.loc[df[key] == point, 'Note'] = note

    return df


def resolve_multi_qualified_teams(res):
    master_frame = pandas.DataFrame([], columns=["Holdnavn", "Quizsteder"])
    for bar_name, df in res:
        for team_name in df['Holdnavn'].values.tolist():
            master_frame.loc[len(master_frame.index)] = [team_name, bar_name]
    new_res = []
    fixed_teams = []
    duplicates = master_frame[master_frame.duplicated(subset=['Holdnavn'], keep=False)]
    for bar_name, df in res:
        for team_name in df['Holdnavn'].unique():
            if team_name in duplicates['Holdnavn'].values.tolist():
                if team_name not in fixed_teams:
                    duplicate_bar_names = duplicates.loc[duplicates['Holdnavn'] == team_name]['Quizsteder'].values.tolist()
                    current_error = df.loc[df['Holdnavn'] == team_name]['Note'].values.tolist()[0]
                    if len(current_error) > 1:
                        df.loc[df[
                                   'Holdnavn'] == team_name, 'Note'] = f'{current_error}, deltager i både {" og ".join(duplicate_bar_names)}'
                    else:
                        df.loc[
                            df['Holdnavn'] == team_name, 'Note'] = f'Deltager i både {" og ".join(duplicate_bar_names)}'
                    fixed_teams.append(team_name)
                    new_res.append([bar_name, df])
            else:
                new_res.append([bar_name, df])
    return new_res


def find_bar_totals(points_frame, winnings_frame, team_names):
    total_bar_team_points = pandas.DataFrame([], columns=["Holdnavn", "Quizsted", "Quizsted Total", "Højeste Score",
                                                          "Antal Deltagelser", "Gennemsnit", "Gevinst", "Ligger i top 10", "Note"])
    bar_names = points_frame["Quizsted"].unique()
    for bar_name in bar_names:
        for team_name in team_names:
            bar_team_points = points_frame.loc[(points_frame['Holdnavn'] == team_name) & (points_frame['Quizsted'] == bar_name)]
            points = bar_team_points['Point'].sum()
            if len(bar_team_points['Point'].index) > 0:
                max_score = \
                bar_team_points[bar_team_points['Point'] == bar_team_points['Point'].max()]['Point'].values.tolist()[0]
            else:
                max_score = 0
            participations = len(bar_team_points.index)
            average = round(max_score / max(participations, 1), 2)
            total_bar_team_points.loc[len(total_bar_team_points.index)] = [team_name, bar_name, points, max_score,
                                                                           participations, average, winnings_frame.loc[(winnings_frame['Holdnavn'] == team_name) & (winnings_frame['Quizsted'] == bar_name), "Samlet Gevinst"].sum(), '', '']
    res = []
    for bar_name in bar_names:
        bar_df = total_bar_team_points.loc[total_bar_team_points['Quizsted'] == bar_name]
        bar_df = bar_df.sort_values(by=['Antal Deltagelser'], ascending=False, ignore_index=True)
        bar_df = bar_df.sort_values(by=['Højeste Score'], ascending=False, ignore_index=True)
        bar_df = bar_df.sort_values(by=['Quizsted Total'], ascending=False, ignore_index=True).drop(columns='Quizsted')
        bar_df = resolve_equal_score_error(bar_df, 'Quizsted Total')
        bar_df = bar_df[bar_df['Quizsted Total'] != 0]
        res.append([bar_name, bar_df])

    res = resolve_multi_qualified_teams(res)
    return res


files = get_xslx_files()
if len(files) == 0:
    messagebox.showerror(title='Note', message='Der blev ikke fundet nogle excel-filer i mappen!')
    sys.exit(0)
points_frame, winnings_frame, error, error_msg = load_xslx_as_pandas(files)
if error:
    msg = "\n".join(error_msg)
    messagebox.showerror(title='Note', message=msg + '\nPrøv at lukke excel-filerne!')
    sys.exit(0)
sanitized_points_frame, error, error_msg = raise_points_frame_errors(points_frame)
if error:
    msg = "\n".join(error_msg)
    messagebox.showerror(title='Note', message=msg + '\nHusk at gemme excel filerne bagefter!')
    sys.exit(0)
generate_output(sanitized_points_frame, winnings_frame)
