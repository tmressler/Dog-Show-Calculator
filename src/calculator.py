import os
import argparse
import json
import re
import pandas
from pandas.core.frame import DataFrame
from pandas.core.series import Series

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.worksheet import Worksheet

'''
Obedience 

Dog number
Dog name
Breed
Group 
Score
Class 

Placement in class first second third and forth highest scores

Highest score in B classes

Highest score combined  utility B and open B

Highest score in each group (herding, sporting, hound, toy, terrier, non-sporting, working)


Rally

Number 
Dog name
Score 
class 

First second third and forth place

High combined (rally excellent b and rally advanced b)

High triple ( rally advanced b, rally excellent b, and rally master)
'''


def calculate(input_file: str, output_file: str, type: str, break_tie: bool = False) -> tuple[int, str]:
    def find_variable_names(dataframe: DataFrame) -> tuple[str, str, str, str, str, str, str]:
        # determine number column
        r = re.compile(r"[Nn]umber(s)?")
        matches = list(filter(r.search, dataframe))
        if len(matches) == 1:
            number_name = matches[0]
        elif len(matches) > 1:
            raise KeyError("multiple 'number' columns detected")
        else:
            raise KeyError("no 'number' column detected")

        # determine handler column
        r = re.compile(r"[Hh]andler(s)?")
        matches = list(filter(r.search, dataframe))
        if len(matches) == 1:
            handler_name = matches[0]
        elif len(matches) > 1:
            raise KeyError("multiple 'handler' columns detected")
        else:
            raise KeyError("no 'handler' column detected")

        # determine call name column
        r = re.compile(r"[Cc]all [Nn]ame(s)?")
        matches = list(filter(r.search, dataframe))
        if len(matches) == 1:
            call_name = matches[0]
        elif len(matches) > 1:
            raise KeyError("multiple 'call name' columns detected")
        else:
            raise KeyError("no 'call name' column detected")

        # determine class column
        r = re.compile(r"[Cc]lass(es)?")
        matches = list(filter(r.search, dataframe))
        if len(matches) == 1:
            class_name = matches[0]
        elif len(matches) > 1:
            raise KeyError("multiple 'class' columns detected")
        else:
            raise KeyError("no 'class' column detected")

        # determine score column
        r = re.compile(r"[Ss]core(s)?")
        matches = list(filter(r.search, dataframe))
        if len(matches) == 1:
            score_name = matches[0]
        elif len(matches) > 1:
            raise KeyError("multiple 'score' columns detected")
        else:
            raise KeyError("no 'score' column detected")

        if type == 'obedience':
            # determine champion column
            r = re.compile(r"[Cc]hamp(ion)?(s)?")
            matches = list(filter(r.search, dataframe))
            if len(matches) == 1:
                champion_name = matches[0]
            elif len(matches) > 1:
                raise KeyError("multiple 'champion' columns detected")
            else:
                raise KeyError("no 'champion' column detected")

            # determine group column
            r = re.compile(r"[Gg]roup(s)?")
            matches = list(filter(r.search, dataframe))
            if len(matches) == 1:
                group_name = matches[0]
            elif len(matches) > 1:
                raise KeyError("multiple 'group' columns detected")
            else:
                raise KeyError("no 'group' column detected")
        else:
            champion_name, group_name = None, None

        return number_name, handler_name, call_name, champion_name, group_name, class_name, score_name

    def augment_data(contestants: DataFrame) -> DataFrame:
        """
        Removes nan rows and removes contestants with 'AB' or 'NQ' scores.
        :param contestants: a dataframe of contestants
        :return: a dataframe of contestants
        """

        # remove NaN entries and remove contestants who did not qualify or were absent
        contestants = contestants.loc[(~contestants[score_name].isna()) &
                                      (~contestants[score_name].str.contains(r'\D', na=False))].copy()

        # convert scores with + values to their base floats, with the count of +s in another column
        contestants.loc[:, 'Pluses'] = contestants[score_name].astype(str).str.count(r'\+')
        contestants.loc[:, score_name] = contestants[score_name].astype(str).replace(r'\++', '', regex=True).astype(
            float)

        rank_dict = {class_.lower(): len(settings["award hierarchy"]) - rank for rank, class_ in
                     enumerate(settings["award hierarchy"])}
        rank_dict[r".+"] = 0  # all other classes are ranked last
        contestants.loc[:, 'Class Rank'] = contestants[class_name].str.lower().replace(rank_dict, regex=True)

        return contestants

    def find_winners(contestants: DataFrame, condition: Series, is_award: bool, combine_scores=False,
                     break_tie: bool = False) -> DataFrame:

        # separate out contestants meeting condition and sort them based on numerical score
        contestants = contestants.loc[condition].sort_values(score_name, ascending=False).copy()  # copy dataframe

        '''# if no contestants match criteria, return
        if contestants.shape[0] == 0:
            return contestants'''

        if is_award:
            # exclude beginner and preferred classes
            contestants = contestants.loc[(~contestants[class_name].str.contains(r'begin|preferred', case=False))].copy()

            # in case the contestants only had excluded classes
            if contestants.shape[0] == 0:
                return contestants

            # only used for specific awards
            if combine_scores:

                # pull out contestants that appear in all classes
                contestant_class_counts = contestants[call_name].value_counts()
                contestants_in_all_classes = contestants[contestants[call_name].isin(
                    contestant_class_counts.index[contestant_class_counts.ge(combine_scores)])]
                contestant_scores = contestants_in_all_classes.groupby(call_name)[score_name].sum().sort_values(
                    ascending=False)
                winning_contestants = contestant_scores.loc[contestant_scores == contestant_scores.max()]
                winning_info = contestants[contestants[call_name].isin(winning_contestants.keys())]

                combined_info = winning_info[~winning_info[call_name].duplicated()]
                combined_info.loc[:, score_name] = winning_contestants.max()

                if break_tie and combined_info.shape[0] > 1:
                    # sort based on class and scores
                    sorted_winning_contestants = contestants[
                        contestants[call_name].isin(winning_contestants.keys())].sort_values(
                        ['Class Rank', score_name, 'Pluses'], ascending=False)

                    return combined_info[combined_info[call_name] == sorted_winning_contestants.iloc[0][call_name]]

                else:
                    return combined_info

            else:
                # the highest scores amongst contestants
                contestants = contestants[contestants[score_name] == contestants[score_name].max()]
                if break_tie:  # break tie by class
                    score, class_rank, pluses = \
                        contestants.sort_values([score_name, 'Class Rank', 'Pluses'], ascending=False).iloc[0][
                            [score_name, 'Class Rank', 'Pluses']]
                    return contestants.loc[(contestants[score_name] == score) &
                                           (contestants['Class Rank'] == class_rank) &
                                           (contestants['Pluses'] == pluses)]
                else:
                    # if there is more than one class, don't break by pluses
                    if len(contestants[class_name].unique()) > 1:
                        return contestants
                    else:
                        score, pluses = contestants.sort_values([score_name, 'Pluses'], ascending=False).iloc[0][
                            [score_name, 'Pluses']]
                        return contestants.loc[(contestants[score_name] == score) &
                                               (contestants['Pluses'] == pluses)]

        # all the same class, return top 4, tie breaking done by pluses
        else:
            return contestants.sort_values([score_name, 'Pluses'], ascending=False)[:4]

    def write_winners(workbook: Workbook, contestants: DataFrame, competition_name: str, is_award: bool) -> None:
        if is_award:
            for c, contestant in enumerate(
                    dataframe_to_rows(contestants[[number_name, handler_name, call_name]], index=False, header=False)):
                if c == 0:
                    workbook.append([competition_name] + contestant)

                    # format competition name
                    workbook[f"A{workbook.max_row}"].font = Font(bold=True)
                else:
                    workbook.append([''] + contestant)

        else:
            # write competition name and format cells
            workbook.append(['', competition_name])
            workbook[f"B{workbook.max_row}"].font = Font(bold=True)
            workbook[f"B{workbook.max_row}"].alignment = Alignment(horizontal='center')
            workbook.merge_cells(start_row=workbook.max_row, start_column=2, end_row=workbook.max_row, end_column=4)

            # write subheaders
            workbook.append(['', number_name, handler_name, call_name])
            workbook[f"B{workbook.max_row}"].font = Font(bold=True)  # TODO: probably a better way to do more than one
            workbook[f"C{workbook.max_row}"].font = Font(bold=True)
            workbook[f"D{workbook.max_row}"].font = Font(bold=True)

            color_key = {0: '000000FF', 1: '00FF0000', 2: '00FFFF00', 3: '00FFFFFF'}
            text_key = {0: '00FFFFFF', 1: '00FFFFFF', 2: '00000000', 3: '00000000'}

            # write winners
            if contestants.shape[0] < 1:
                workbook.append([''] + ['-'] * 3)
            else:
                position = contestants[score_name].rank(method='min', ascending=False).astype(int)
                for c, contestant in enumerate(
                        dataframe_to_rows(contestants[[number_name, handler_name, call_name]], index=False,
                                          header=False)):
                    workbook.append([position.iloc[c]] + contestant)

                    workbook[f"B{workbook.max_row}"].font = Font(color=text_key[c])
                    workbook[f"B{workbook.max_row}"].fill = PatternFill(start_color=color_key[c],
                                                                        end_color=color_key[c], fill_type='solid')
                    workbook[f"C{workbook.max_row}"].font = Font(color=text_key[c])
                    workbook[f"C{workbook.max_row}"].fill = PatternFill(start_color=color_key[c],
                                                                        end_color=color_key[c], fill_type='solid')
                    workbook[f"D{workbook.max_row}"].font = Font(color=text_key[c])
                    workbook[f"D{workbook.max_row}"].fill = PatternFill(start_color=color_key[c],
                                                                        end_color=color_key[c], fill_type='solid')

            workbook.append([''])

    def optimal_width(sheet: Worksheet) -> None:
        """
        Formats the worksheet so each column is the width of the largest cell value.
        :param sheet: the worksheet to resize
        """
        for column in sheet.columns:
            new_column_length = max(len(str(cell.value)) for cell in column)  # find max cell width in column
            column_letter = openpyxl.utils.get_column_letter(column[0].column)
            sheet.column_dimensions[column_letter].width = new_column_length

    # make sure input file is valid
    if input_file[-5:] != '.xlsx':
        return 0, "only .xlsx is supported"

    # check to see if file exists
    if not os.path.exists(input_file):
        return 0, "file not found"

    # import settings
    if os.path.exists("settings.json"):
        with open("settings.json", 'r') as f:
            settings = json.load(f)
    else:
        with open("settings.json", 'w') as f:
            json.dump({{"award hierarchy": ["utility", "open", "novice"],
                        "defaults": {"multiple award winners": True}}}, f)

    data = pandas.read_excel(input_file)

    # find column names
    number_name, handler_name, call_name, champion_name, group_name, class_name, score_name = find_variable_names(data)

    classes = data[class_name].dropna().unique()  # find set of classes

    data = augment_data(data)  # remove nan rows and contestants with 'AB' or 'NQ'

    if input_file == output_file:
        output_workbook = openpyxl.load_workbook(input_file)
        output_workbook.create_sheet("Winners")
        output_worksheet = output_workbook["Winners"]
    else:
        output_workbook = Workbook()
        output_worksheet = output_workbook.active

    for class_ in classes:
        winners = find_winners(data, Series(data[class_name] == class_), False, break_tie=break_tie)
        write_winners(output_worksheet, winners, class_, False)

    # add headers for award winners
    output_worksheet.append(['', number_name, handler_name, call_name])
    output_worksheet[f'B{output_worksheet.max_row}'].font = Font(bold=True)
    output_worksheet[f'C{output_worksheet.max_row}'].font = Font(bold=True)
    output_worksheet[f'D{output_worksheet.max_row}'].font = Font(bold=True)

    if type == "obedience":
        award_conditions = (
            ("High in Trial", data[class_name].str.contains(r"\sb$", case=False), False),
            ("High Combined", ((data[class_name].str.contains("open b", case=False)) |
                               (data[class_name].str.contains("utility b", case=False))), 2),
            ("High Combined Preferred",
             ((data[class_name].str.contains("[Pp]referred [Oo]pen")) |
              (data[class_name].str.contains("[Pp]referred [Uu]tility"))), 2),
            ("High Scoring Champion of Record", data[champion_name].str.contains('Ch', na=False), False)
        )
    elif type == "rally":
        award_conditions = (
            ("High Combined", ((data[class_name].str.contains(r"rally excellent b", case=False)) |
                               (data[class_name].str.contains(r"rally advanced? b", case=False))), 2),
            ("High Triple", ((data[class_name].str.contains(r"rally excellent b", case=False)) |
                             (data[class_name].str.contains(r"rally advanced? b", case=False)) |
                             (data[class_name].str.contains(r"rally master", case=False))), 3)
        )
    else:
        raise ValueError("type must be 'obedience' or 'rally'")

    # find award winners
    for award, award_condition, combine_scores in award_conditions:
        winners = find_winners(data, award_condition, True, combine_scores=combine_scores, break_tie=break_tie)
        write_winners(output_worksheet, winners, award, True)

    if type == "obedience":
        groups = [group for group in data[group_name].unique() if str(group) != 'nan']  # find set of groups

        # find group winners
        for group in groups:
            winners = find_winners(data, Series(data[group_name] == group), True, break_tie=break_tie)
            write_winners(output_worksheet, winners, group, True)

    optimal_width(output_worksheet)

    output_workbook.save(output_file)

    return 1, ""


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Determine the winners of a dog show.")
    parser.add_argument('input_file', type=open, help="the input .xlsx file")
    parser.add_argument('output_file', help="the output file")
    parser.add_argument('competition_type', choices=['obedience', 'rally'], help="the type of competition")
    parser.add_argument('-n', '--no_ties', action='store_true', help="break ties by class if possible")
    args = parser.parse_args()

    status, message = calculate(args.input_file.name, args.output_file, args.competition_type, args.no_ties)

    if status:
        print("Complete")
    else:
        print(f"Error: {message}")
