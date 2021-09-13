from tkinter import filedialog
from time import sleep
import tkinter as tk
import pandas as pd
import numpy as np
import subprocess
import sys
import os

root = tk.Tk()
root.withdraw()


def empty_string_to_null(input_object):
    if pd.isna(input_object):
        return np.nan
    elif str(input_object).lower() in ('', 'nan', 'nat', 'none'):
        return np.nan
    elif isinstance(input_object, str) and any([input_object.isspace(), not input_object]):
        return np.nan
    elif input_object is None:
        return np.nan
    return input_object


def get_master_from_xlsx(path_to_master_file):
    master_all_sheets = pd.ExcelFile(path_to_master_file)
    master_dataframe = master_all_sheets.parse('Schedule', header=None, usecols='B:G')
    if len(list(master_dataframe)) == 5:
        master_dataframe['points'] = ''
    master_dataframe.columns = ['visitors_choice', 'visitors', 'name', 'home_choice', 'home', 'points']
    master_dataframe = master_dataframe.applymap(empty_string_to_null)

    master_dataframe['visitor_potential_game'] = pd.notna(master_dataframe['visitors'])
    master_dataframe['home_potential_game'] = pd.notna(master_dataframe['home'])
    master_dataframe['not_visitor_home'] = np.logical_not(master_dataframe['home'].str.upper().str.contains('TEAM'))
    master_dataframe['is_a_game'] = master_dataframe['visitor_potential_game'] & master_dataframe['home_potential_game'] & master_dataframe['not_visitor_home']

    master_dataframe['says_football'] = master_dataframe['visitors_choice'].str.upper().str.contains('FOOTBALL')
    master_dataframe['above_says_football'] = master_dataframe['says_football'].shift(1)
    master_dataframe['says_week'] = master_dataframe['visitors_choice'].str.upper().str.contains('WEEK')
    master_dataframe['is_week_number'] = master_dataframe['above_says_football'] & master_dataframe['says_week']

    master_dataframe['dad_marked_visitor'] = pd.notna(master_dataframe['visitors_choice'])
    master_dataframe['dad_marked_home'] = pd.notna(master_dataframe['home_choice'])
    master_dataframe['dad_marked_something'] = master_dataframe['dad_marked_visitor'] | master_dataframe['dad_marked_home']
    master_dataframe['dad_marked_nothing'] = np.logical_not(master_dataframe['dad_marked_something'])

    master_dataframe['visitor_won'] = master_dataframe['is_a_game'] & master_dataframe['dad_marked_visitor']
    master_dataframe['home_won'] = master_dataframe['is_a_game'] & master_dataframe['dad_marked_home']

    master_dataframe['complete_game'] = master_dataframe['is_a_game'] & master_dataframe['dad_marked_something']
    master_dataframe['incomplete_game'] = master_dataframe['is_a_game'] & master_dataframe['dad_marked_nothing']

    print('\nOkay, I think the winners are:\n')
    sleep(0.5)

    games_not_completed = len(master_dataframe[master_dataframe['incomplete_game'] == True])
    if games_not_completed > 0:
        print('It looks like I have', games_not_completed, 'unfinished games.')
        if input('Does that seem right? (y/n) ') != 'y':
            print('Please exit the program and correct the file.\n')
            sys.exit()

    total_points_correct = 0
    master_dataframe['is_tie_breaker'] = np.where(master_dataframe['visitors_choice'].str.contains('Total Combined Points')
                                                  & master_dataframe['visitors_choice'].notna(),
                                                  True,
                                                  False)

    for index, row in master_dataframe.iterrows():

        visitor_team = row.visitors
        home_team = row.home
        visitor_won = row.visitor_won
        home_won = row.home_won

        if visitor_won & home_won:
            winner = 'Tie between the ' + visitor_team + ' and the ' + home_team
            print(winner)

        elif visitor_won & np.logical_not(home_won):
            winner = visitor_team
            print(winner)

        elif np.logical_not(visitor_won) & home_won:
            winner = home_team
            print(winner)

        if row.is_week_number:
            week = str(master_dataframe.at[index, 'visitors_choice']).strip()
            week_number = int(''.join(filter(str.isdigit, week)))
            print(f'WEEK {week_number}\n')

        if row.is_tie_breaker:
            try:
                for column_to_try in ('points', 'visitors_choice'):
                    guess_cell = empty_string_to_null(str(master_dataframe.at[index, column_to_try]).strip().split('.0')[0])
                    if pd.notna(guess_cell):
                        break

                total_points_correct = int(''.join(filter(str.isdigit, guess_cell)))
                print('\nTotal Points Combined:', total_points_correct, '\n')

                sleep(0.5)
                # I now have which games are correct, as well as the correct point total.
                if input('Is this what you have? Enter to continue, type anything if not. ') != '':
                    print('\nPlease correct the file and restart the program.')
                    sleep(0.5)
                    sys.exit()

            except ValueError:
                if input('\nHmmm.. I don\'t see any points for Monday. Does that sound right? (y/n) ') != 'y':
                    print('\nPlease correct the file and restart the program.\n')
                    sleep(0.5)
                    sys.exit()

                else:
                    # I should have some correct games, and now correct points is zero
                    if input('\nAlright, so continuing like normal. Is the above what you have? Enter to continue, type anything if not. ') != '':
                        print('\nPlease correct the file and restart the program.')
                        sys.exit()

    return master_dataframe, week_number, total_points_correct


def potentially_inspect(dataframe, sheet, filename_with_xlsx, look_at=None):
    if look_at:
        if not look_at.endswith('.xlsx'):
            look_at += '.xlsx'
        if look_at == filename_with_xlsx:
            satisfied = False
            while not satisfied:
                try:
                    results_name = f'Inspection of {look_at}'
                    with pd.ExcelWriter(results_name) as writer:
                        dataframe.to_excel(writer, sheet_name=sheet, index=False)
                        format_excel_worksheet(writer.sheets[sheet], dataframe)

                    if sys.platform == "win32":
                        os.startfile(results_name)
                    else:
                        opener = "open" if sys.platform == "darwin" else "xdg-open"
                        subprocess.call([opener, results_name])
                except Exception:
                    pass
                else:
                    satisfied = True


def grade_participant(master_dataframe, results_dataframe, filename_with_xlsx, path, total_points_correct, look_at=None):
    filename_w_o_xlsx = filename_with_xlsx.split('.xlsx')[0]

    participant_all_sheets = pd.ExcelFile(path + '/' + filename_with_xlsx)
    for sheet in set(participant_all_sheets.sheet_names).difference({'Weekly Results', 'WeeklyResults', 'Export Summary'}):
        try:
            participant_dataframe = participant_all_sheets.parse(sheet, header=None, usecols='B:G')
            if len(list(participant_dataframe)) == 5:
                participant_dataframe['points'] = ''
            participant_dataframe.columns = ['visitors_choice', 'visitors', 'name', 'home_choice', 'home', 'points']
            participant_dataframe = participant_dataframe.applymap(empty_string_to_null)

            participant_dataframe['is_correct'] = ''
            participant_dataframe['extra_stuff --->'] = ' '

            participant_dataframe['is_a_game'] = master_dataframe['is_a_game']
            participant_dataframe['dad_marked_something'] = master_dataframe['dad_marked_something']
            participant_dataframe['visitor_won'] = master_dataframe['visitor_won']
            participant_dataframe['home_won'] = master_dataframe['home_won']
            participant_dataframe['is_tie_breaker'] = master_dataframe['is_tie_breaker']
            participant_dataframe['complete_game'] = master_dataframe['complete_game']
            participant_dataframe['incomplete_game'] = master_dataframe['incomplete_game']

            participant_dataframe['p_marked_visitor'] = pd.notna(participant_dataframe['visitors_choice'])
            participant_dataframe['p_marked_home'] = pd.notna(participant_dataframe['home_choice'])

            participant_dataframe['p_visitor_chosen'] = participant_dataframe['is_a_game'] & participant_dataframe['p_marked_visitor']
            participant_dataframe['p_home_chosen'] = participant_dataframe['is_a_game'] & participant_dataframe['p_marked_home']

            participant_name = str(participant_dataframe.at[0, 'name']).strip()

            total_points_guessed = -1000
            points_off = -1000
            points_off_sort = -1001

            for index, row in participant_dataframe.iterrows():

                visitor_won = row.visitor_won
                home_won = row.home_won

                complete_game = row.complete_game

                if complete_game:

                    # determine outcome

                    if visitor_won and home_won:
                        outcome = 'Tie'
                    elif visitor_won and not home_won:
                        outcome = 'Visitor'
                    elif not visitor_won and home_won:
                        outcome = 'Home'
                    else:
                        outcome = 'No game chosen yet'

                    # determine choice

                    picked_visitor = row.p_marked_visitor
                    picked_home = row.p_marked_home

                    if picked_visitor and picked_home:
                        choice = 'Tie'
                    elif picked_visitor and not picked_home:
                        choice = 'Visitor'
                    elif not picked_visitor and picked_home:
                        choice = 'Home'
                    else:
                        choice = 'No choice made'

                    is_correct = (outcome == choice)
                    participant_dataframe.at[index, 'is_correct'] = is_correct

                if row.is_tie_breaker:
                    for column_to_try in ('points', 'visitors_choice'):
                        guess_cell = empty_string_to_null(str(participant_dataframe.at[index, column_to_try]).strip().split('.0')[0])
                        if pd.notna(guess_cell):
                            break
                    try:
                        total_points_guessed = int(''.join(filter(str.isdigit, guess_cell)))
                    except ValueError:
                        total_points_guessed = -1000

                    if total_points_guessed != -1000:
                        try:
                            points_off = int(abs(total_points_guessed - total_points_correct))
                        except ValueError:
                            points_off = -1000

                    if points_off == -1000:
                        points_off_sort = np.inf
                    else:
                        points_off_sort = points_off

            p_games_correct = len(participant_dataframe[participant_dataframe['complete_game'] & participant_dataframe['is_correct'] == True])
            p_games_incorrect = len(participant_dataframe[participant_dataframe['complete_game'] & ~(participant_dataframe['is_correct'] == True)])

            participant_score_row = pd.DataFrame({
                'Sorting Name'   : filename_w_o_xlsx,
                'Name on Sheet'  : participant_name,
                'Correct'        : p_games_correct,
                'Incorrect'      : p_games_incorrect,
                'Points Guessed' : total_points_guessed,
                'Points off Sort': points_off_sort,
                'Points off'     : points_off
            },
                index=[0])

            potentially_inspect(participant_dataframe, sheet, filename_with_xlsx, look_at)
            return results_dataframe.append(participant_score_row)
        except Exception as e:
            print(f'Unable to parse {sheet} within {filename_with_xlsx}. The exception is {e}')
            pass


def format_excel_worksheet(worksheet, dataframe):
    for i, col in enumerate(list(dataframe)):
        iterate_length = dataframe[col].astype(str).str.len().max()
        header_length = len(col)
        max_size = max(iterate_length, header_length) + 1
        worksheet.set_column(i, i, max_size)


def conditional_format(worksheet, workbook, column_format_range, winning_number_of_games):
    if workbook:
        colors_dictionary = {
            '0'                    : {
                'bg_color'  : '#FFC7CE',
                'font_color': '#9C0006'
            },
            winning_number_of_games: {
                'bg_color'  : '#C6EFCE',
                'font_color': '#006100'
            }
        }
        for if_equals, format_dictionary in colors_dictionary.items():
            excel_format = workbook.add_format(format_dictionary)
            worksheet.conditional_format(column_format_range, {
                'type'    : 'cell',
                'criteria': '=',
                'value'   : if_equals,
                'format'  : excel_format
            })


# def export_logic_oldish(logic_dataframe, week_number):
#     logic_name = 5
#     with pd.ExcelWriter(logic_name, engine='xlsxwriter') as logic_writer:
#         logic_dataframe.to_excel(logic_writer, sheet_name='Schedule', index=False)
#         format_excel_worksheet(logic_writer.sheets['Schedule'], logic_dataframe)


# def export_results_old(results_dataframe, week_number):
#     results_name = f'Results for Week {week_number}.xlsx'
#     with pd.ExcelWriter(results_name) as report_writer:
#         results_dataframe.to_excel(report_writer, sheet_name='Results', index=False)
#         format_excel_worksheet(report_writer.sheets['Results'], results_dataframe)


def remove_inbetween_quotations(name):
    try:
        index_for_first_quotation = name.find('"')
        index_for_second_quotation = name.find('"', index_for_first_quotation + 1)
        return name[:index_for_first_quotation] + name[index_for_second_quotation + 1:]
    except Exception:
        return name


def remove_inbetween_open_and_close_paren(name):
    try:
        index_for_open_paren = name.find('(')
        index_for_close_paren = name.find(')', index_for_open_paren + 1)
        return name[:index_for_open_paren] + name[index_for_close_paren + 1:]
    except Exception:
        return name


def remove_and_following(name, and_phrase):
    try:
        index_for_and = name.find(and_phrase)
        index_for_following = name.find(' ', index_for_and + 1)
        return name[:index_for_and] + name[index_for_following + 1:]
    except Exception:
        return name


def quotation_cleaner(name):
    while '"' in name:
        name = remove_inbetween_quotations(name)
    return name


def paren_cleaner(name):
    while '(' in name and ')' in name:
        name = remove_inbetween_open_and_close_paren(name)
    return name


def and_cleaner(name):
    for and_phrase in (' and ', ' & '):
        while and_phrase in name:
            name = remove_and_following(name, and_phrase=and_phrase)
    return name


def get_first_three_first_and_last(name):
    name = name.strip()
    for cleaner in (quotation_cleaner, paren_cleaner, and_cleaner):
        name = cleaner(name)

    formatted_name = ''
    name_split = list(filter(None, [word.strip() for word in name.split(' ')]))
    for i, word in enumerate(name_split):
        formatted_name += ' ' if 0 < i < len(name_split) else ''
        if i < 2:
            if word.startswith('Chri') and word not in ('Chris', 'Christopher'):
                formatted_name += 'Chri'
            else:
                formatted_name += word[:3]
        elif i == 2:
            formatted_name += word[0]
    return formatted_name



def get_letter_from_column(dataframe, week_string):
    for i, col in enumerate(list(dataframe)):
        if col == week_string:
            return 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'[i]


def get_filename_and_sheetname(label):
    if label.endswith('.xlsx'):
        return label, label.split('xlsx')[0]
    else:
        return label + '.xlsx', label


def export_excel(dataframe, label):
    filename, sheetname = get_filename_and_sheetname(label)
    print(f'Now exporting: {filename}')

    with pd.ExcelWriter(filename) as report_writer:
        dataframe.to_excel(report_writer, sheet_name=sheetname, index=False)
        format_excel_worksheet(report_writer.sheets[sheetname], dataframe)


def export_results(path_to_masterfile, label, week_number, winning_number_of_games, results_dataframe):
    filename, sheetname = get_filename_and_sheetname(label)
    print(f'Now exporting: {filename}')

    weekly_results = pd.ExcelFile(path_to_masterfile).parse('Weekly Results')
    week_string = f'Week {week_number:02}'
    weekly_results['first_three_first_and_last'] = weekly_results[list(weekly_results)[0]].apply(get_first_three_first_and_last)

    dataframe = pd.merge(weekly_results, results_dataframe[['Sorting Name', 'Name on Sheet', 'Correct']], how='outer', left_on='first_three_first_and_last', right_on='Sorting Name')
    dataframe.rename(columns=lambda x: x.strip(), inplace=True)
    with pd.ExcelWriter(filename) as writer:
        dataframe[week_string] = dataframe['Correct']
        dataframe['Totals'] = dataframe['Totals'] + dataframe['Correct']
        dataframe[list(dataframe)[0]].fillna('  ' + dataframe['Name on Sheet'], inplace=True)

        for col in (week_string, 'Totals'):
            dataframe[col].fillna(0, inplace=True)

        dataframe.to_excel(writer, sheet_name=sheetname, index=False)

        letter = get_letter_from_column(dataframe, week_string)
        format_excel_worksheet(writer.sheets[sheetname], dataframe)
        conditional_format(worksheet=writer.sheets[sheetname],
                           workbook=writer.book,
                           column_format_range=f'{letter}1:{letter}{len(dataframe)}',
                           winning_number_of_games=winning_number_of_games)


def main():
    results_dataframe = pd.DataFrame()

    print('\nWelcome to TSFL, the Tom Smith Football League.\n')
    ready_answer = input('Are you ready for some foootballlll? (y/n) ')
    sleep(0.5)

    # ready_answer = 'y'
    # ready_answer = 'inspect'
    if ready_answer.lower() == 'y':
        look_at = None
    elif ready_answer.lower() == 'inspect':
        look_at = input('\nWhich file do you want to take a look at? ')
        satisfied = True if look_at else False
        while not satisfied:
            look_at = input('\nSorry, I didn\'t catch that. Which file do you want to take a look at? ')
            satisfied = True if look_at else False

    else:
        print('Alright, restart when you\'re ready!')
        sys.exit()

    input('\nLet\'s get your answer sheet! Cool? Press enter to continue. ')
    path_to_masterfile = filedialog.askopenfilename()
    # print(path_to_masterfile)
    # path_to_masterfile = '/Users/spencer.smith/Documents/Self/Python_Football/picks/Tom Smi.xlsx'
    grading_dataframe, week_number, total_points_correct = get_master_from_xlsx(path_to_masterfile)

    input('\nGreat! Now let\'s go to this week\'s folder! Press enter when you\'re ready.\n')
    path = filedialog.askdirectory()
    # print(path)
    # path = '/Users/spencer.smith/Documents/Self/Python_Football/picks'
    directory = os.fsencode(path)

    print('Awesome... here we go!')
    sleep(1.5)

    print('\nReady?\n')
    sleep(1)
    print('3..\n')
    sleep(1)
    print('2..\n')
    sleep(1)
    print('1..\n')
    sleep(1)

    files_parsed = []
    master_filename = path_to_masterfile.split('/')[-1]

    for file in sorted(os.listdir(directory)):
        filename = os.fsdecode(file)

        if all([filename.endswith('.xlsx'), filename != master_filename, not filename.startswith('~$')]):
            try:
                results_dataframe = grade_participant(master_dataframe=grading_dataframe,
                                                      results_dataframe=results_dataframe,
                                                      filename_with_xlsx=filename,
                                                      path=path,
                                                      total_points_correct=total_points_correct,
                                                      look_at=look_at)

                files_parsed += [filename]
            except Exception:
                print(f'We went through: {files_parsed}')
                files_parsed = []
                print(f'Unable to parse: {filename}')
                pass

    for col in ('Points Guessed', 'Points off'):
        results_dataframe[col] = results_dataframe[col].astype(int).replace(-1000, 'Error')

    results_dataframe = results_dataframe.sort_values(['Correct', 'Points off Sort'], ascending=[False, True])
    results_dataframe.drop(columns='Points off Sort', inplace=True)

    winners_dataframe = results_dataframe[results_dataframe['Correct'] == results_dataframe['Correct'].max()]

    print('Congratulations to... ')
    sleep(1)
    print('\n', winners_dataframe.to_string(), '\n')
    sleep(1.5)

    print('Here are your full results: ')
    sleep(0.5)
    print('\nTotal Points Combined:', total_points_correct, '\n')

    print('\n', results_dataframe.to_string(), '\n\n')

    export_excel(grading_dataframe, f'Scoring Logic for Week {week_number}.xlsx')

    try:
        export_results(path_to_masterfile=path_to_masterfile,
                       label=f'Results for Week {week_number}',
                       week_number=week_number,
                       winning_number_of_games=winners_dataframe['Correct'].max(),
                       results_dataframe=results_dataframe)
    except Exception as e:
        print(f'We were unable to nicely format the scores for you, and the error was {e}.')
        print(f'But your results and logic files should survive unscathed.')
        export_excel(grading_dataframe, f'Results for Week {week_number}.xlsx')

    sleep(1.5)
    print('\nLove you always Dad, Spence.\n')
    sleep(5)
    input('Press enter to close.\n')


if __name__ == '__main__':
    main()
