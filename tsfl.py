import pandas as pd
import sys
import os
import numpy as np
import tkinter as tk
from tkinter import filedialog
import time
test_mode = 1

root = tk.Tk()
root.withdraw()

def uni_type_fixer(dataframe, value, cast_to):
    return dataframe[value].astype(cast_to)

try:
    report_df = pd.DataFrame()

    print("")
    print("Welcome to TSFL, the Tom Smith Football League.")
    print("")
    if input("Are you ready for some foootballlll? (y/n) ") != 'y':
        print("Alright, restart when you're ready!")
        sys.exit()
    print("")
    print("Awesome! I made a change that should make things easier.")
    time.sleep(1)
    print("")
    print("Now you can safely put your answer file in the SAME folder as the other sheets. Should be more convenient!")
    time.sleep(3)
    print("")
    input('Let\'s get your answer sheet! Cool? Press enter to continue. ')
    file_path = filedialog.askopenfilename()

    dad_sheet_split = file_path.split('/')
    dad_sheet_name = dad_sheet_split[len(dad_sheet_split)-1]

    xl = pd.ExcelFile(file_path)

    master_df = xl.parse("Schedule", header=None, usecols='B:F')

    master_df.columns = ["v", "visitors", "name", "h", "home"]

    dads_name = master_df.at[0, 'name']
    master_df['visitor_potential_game'] = pd.notna(master_df['visitors'])
    master_df['home_potential_game'] = pd.notna(master_df['home'])
    master_df['not_visitor_home'] = np.logical_not(master_df['home'].str.contains("TEAM"))
    master_df['is_a_game'] = master_df['visitor_potential_game'] & master_df['home_potential_game'] & master_df['not_visitor_home']

    master_df['dad_marked_visitor'] = pd.notna(master_df['v'])
    master_df['dad_marked_home'] = pd.notna(master_df['h'])
    master_df['dad_marked_something'] = master_df['dad_marked_visitor'] | master_df['dad_marked_home']
    master_df['dad_marked_nothing'] = np.logical_not(master_df['dad_marked_something'])

    master_df['visitor_won'] = master_df['is_a_game'] & master_df['dad_marked_visitor']
    master_df['home_won'] = master_df['is_a_game'] & master_df['dad_marked_home']

    master_df['complete_game'] = master_df['is_a_game'] & master_df['dad_marked_something']
    master_df['incomplete_game'] = master_df['is_a_game'] & master_df['dad_marked_nothing']

    games_completed = len(master_df.query('complete_game == True').index)
    games_not_completed = len(master_df.query('incomplete_game == True').index)

    total_points_correct = 0

    print("")
    print("Okay, I think the winners are:")
    time.sleep(1)

    print("")

    if games_not_completed > 0:
        print("It looks like I have", games_not_completed, "unfinished games.")
        if input("Does that seem right? (y/n) ") != 'y':
            print("Please exit the program and correct the file.")
            print("")
            sys.exit()

    master_df['is_tie_breaker'] = master_df['v'].str.contains("Total Combined Points")

    for row in master_df.itertuples(index=True):

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

        if row.is_tie_breaker == True:
            try:
                capture_string = master_df.at[row[0], 'v']
                capture_string = str(capture_string)
                total_points_correct = int("".join(filter(str.isdigit, capture_string)))
                print("")
                print('Total Points Combined:', total_points_correct)
                print("")

                time.sleep(1)
                # I now have which games are correct, as well as the correct point total.
                if input("Is this what you have? Enter to continue, type anything if not. ") != '':
                    print('')
                    print("Please correct the file and restart the program.")
                    time.sleep(1)
                    sys.exit()

            except ValueError:
                time.sleep(1)
                if input("\nHmmm.. I don't see any points for Monday. Does that sound right? (y/n) ") != 'y':
                    print("")
                    print("Please correct the file and restart the program.")
                    print("")
                    time.sleep(1)
                    sys.exit()

                else:
                    total_points_correct = 0
                    # I should have some correct games, and now correct points is zero
                    if input("\nAlright, so continuing like normal. Is the above what you have? Enter to continue, type anything if not. ") != '':
                        print('')
                        print("Please correct the file and restart the program.")
                        sys.exit()

    # Now I need to compare another file's answers and point total to my dad's

    ## i build this dataframe one at a time, and append a full row of information after I iterate through a single file
    ## in alphabetical order:
    ## Name | Total correct-Total Incorrect | Points guessed (+/-X)
    ## then i export this file to a excel/csv, and it should be ready for

    print("\nGreat! Now let's go to this week's folder!")
    time.sleep(1)
    print("")
    input("Press enter when you're ready. ")
    print("")

    # path = '/Users/spencer.smith/Documents/Self/Python_Football/Sheets'
    path = filedialog.askdirectory()
    directory = os.fsencode(path)

    print("Awesome... here goes nothing!")
    time.sleep(3)
    print("")

    print("Ready?")
    print("")
    time.sleep(1)
    print("3..")
    print("")
    time.sleep(1)
    print("2..")
    print("")
    time.sleep(1)
    print("1..")
    print("")
    time.sleep(1)

    for file in os.listdir(directory):
        filename = os.fsdecode(file)

        if filename.endswith(".xlsx") and filename != dad_sheet_name:
            filename_w_o_xlsx = filename.split('.xlsx')[0]

            participant_xl = pd.ExcelFile(path +'/' + filename)
            participant_df = participant_xl.parse("Schedule", header=None, usecols='B:F')
            participant_df.columns = ["v", "visitors", "name", "h", "home"]
            participant_df['is_a_game'] = master_df['is_a_game']
            participant_df['is_a_game'] = master_df['is_a_game']
            participant_df['dad_marked_something'] = master_df['dad_marked_something']
            participant_df['visitor_won'] = master_df['visitor_won']
            participant_df['home_won'] = master_df['home_won']
            participant_df['is_tie_breaker'] = master_df['is_tie_breaker']
            participant_df['complete_game'] = master_df['complete_game']
            participant_df['incomplete_game'] = master_df['incomplete_game']

            participant_df['p_marked_visitor'] = pd.notna(participant_df['v'])
            participant_df['p_marked_home'] = pd.notna(participant_df['h'])

            participant_df['p_visitor_chosen'] = participant_df['is_a_game'] & participant_df['p_marked_visitor']
            participant_df['p_home_chosen'] = participant_df['is_a_game'] & participant_df['p_marked_home']
            participant_name = participant_df.at[0, 'name']

            participant_df['is_correct'] = ''

            total_points_guessed = -100
            sign = 'error'
            points_off = -100
            points_off_sort = -101

            for row_i in participant_df.itertuples(index=True):

                visitor_team = row_i.visitors
                home_team = row_i.home

                dad_marked_something = row_i.dad_marked_something

                visitor_won = row_i.visitor_won
                home_won = row_i.home_won

                complete_game = row_i.complete_game

                is_correct = row_i.is_correct

                if complete_game:

                    # determine outcome

                    if visitor_won and home_won:
                        outcome = 'Tie'
                    elif visitor_won and np.logical_not(home_won):
                        outcome = 'Visitor'
                    elif np.logical_not(visitor_won) and home_won:
                        outcome = 'Home'
                    else:
                        outcome = 'No game chosen yet'

                    # determine choice

                    picked_visitor = row_i.p_marked_visitor
                    picked_home = row_i.p_marked_home

                    if picked_visitor and picked_home:
                        choice = 'Tie'
                    elif picked_visitor and np.logical_not(picked_home):
                        choice = 'Visitor'
                    elif np.logical_not(picked_visitor) and picked_home:
                        choice = 'Home'
                    else:
                        choice = 'No choice made'

                    is_correct = (outcome == choice)
                    participant_df.at[row_i[0], 'is_correct'] = is_correct

                if row_i.is_tie_breaker == True:
                    guess_cell = participant_df.at[row_i[0], 'v']
                    guess_cell = str(guess_cell)
                    try:
                        total_points_guessed = int("".join(filter(str.isdigit, guess_cell)))
                    except ValueError:
                        total_points_guessed = -100

                    if total_points_guessed != -100:
                        try:
                            points_off = int(total_points_guessed - total_points_correct)
                        except ValueError:
                            points_off = -100

                    if points_off == -100:
                        points_off_sort = np.inf
                    else:
                        points_off_sort = int(abs(points_off))

                    if points_off == 0:
                        sign = ''
                    elif points_off > 0:
                        sign = '+'
                    elif points_off < 0:
                        sign = '-'

            p_games_correct = len(participant_df.query('complete_game == True & is_correct == True').index)
            p_games_incorrect = len(participant_df.query('complete_game == True & is_correct == False').index)
            games_not_completed = len(master_df.query('incomplete_game == True').index)

            entry_df = pd.DataFrame({
                'Sorting Name' : filename_w_o_xlsx,
                'Name on Sheet': participant_name,
                'Correct': p_games_correct,
                'Incorrect': p_games_incorrect,
                'Points Guessed': total_points_guessed,
                'Points off Sort': points_off_sort,
                'Points off': points_off},
                index=[0])

            report_df = report_df.append(entry_df)

    report_df['Points Guessed'] = report_df['Points Guessed'].astype(int)
    report_df['Points Guessed'] = report_df['Points Guessed'].replace(-100, "Error")
    report_df['Points off'] = report_df['Points off'].astype(int)
    report_df['Points off'] = report_df['Points off'].replace(-100, "Pending")

    report_df = report_df.sort_values(['Correct', 'Points off Sort'], ascending=[False, True])

    report_df = report_df.drop('Points off Sort', axis=1)

    winners_df = report_df[report_df['Correct'] == report_df['Correct'].max()]

    print("Congratulations to... ")
    print("")
    time.sleep(2)
    print(winners_df.to_string())
    print("")
    time.sleep(3)

    print("Here are your full results: ")
    time.sleep(1)
    print("")
    print(report_df.to_string())
    print("")
    print("")

    report_writer = pd.ExcelWriter('logic.xlsx', engine='xlsxwriter')
    report_df.to_excel(report_writer, sheet_name='Results', index=False)
    worksheet = report_writer.sheets['Results']

    writer = pd.ExcelWriter('answers.xlsx', engine='xlsxwriter')
    master_df.to_excel(writer, sheet_name='Schedule')

    columns_to_hide = ['visitor_potential_game', 'home_potential_game', 'not_visitor_home', 'dad_marked_visitor',
                       'dad_marked_home', 'visitor_won', 'home_won', 'dad_marked_something', 'dad_marked_nothing',
                       'is_tie_breaker']

    for i, col in enumerate(report_df.columns):
        iterate_length = report_df[col].astype(str).str.len().max()
        header_length = len(col)
        max_size = max(iterate_length, header_length) + 1
        worksheet.set_column(i, i, max_size)

    report_writer.close()
    writer.close()
    print("")

    time.sleep(2)
    print("I also printed this to an excel file called Results.xlsx. Check it out!")
    print("")
    time.sleep(3)
    print("Love you Dad, Spence.")
    time.sleep(10)
    print("")
    input("Whenever you are ready, your next input will close out the program.")
    print("")

except Exception:
    print("Ohp! Wait a second... an error happened.\n")
    time.sleep(2)
    print("A likely answer is that you left an excel tab open that you want me to use.")
    time.sleep(1)
    print("Please fix the problem and restart the program. We'll be right back.")
    time.sleep(1)