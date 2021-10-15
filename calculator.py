import os
import sys
import re
import pandas
from pandas.core.frame import DataFrame

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


def calculate(input_file, output_file):
    def find_variable_names(dataframe: pandas.core.frame.DataFrame) -> tuple[str, str]:
        # determine class column
        r = re.compile(r"[Cc]lass(es)?")
        matches = list(filter(r.match, dataframe))
        if len(matches) == 1:
            class_name = matches[0]
        elif len(matches) > 1:
            raise KeyError("multiple 'class' columns detected")
        else:
            raise KeyError("no 'class' column detected")

        # determine score column
        r = re.compile(r"[Ss]core(s)?")
        matches = list(filter(r.match, dataframe))
        if len(matches) == 1:
            score_name = matches[0]
        elif len(matches) > 1:
            raise KeyError("multiple 'score' columns detected")
        else:
            raise KeyError("no 'score' column detected")

        return class_name, score_name

    def plus_to_dec(contestants: DataFrame) -> DataFrame:
        """
        Converts the scores in a dataframe with pluses to a decimal value. For example '192+' -> '192.001'
        :param contestants: a dataframe of contestants
        :return: a dataframe of contestants
        """
        # TODO: move method here

        return contestants

    def dec_to_plus(contestants: DataFrame) -> DataFrame:
        # TODO: move method here
        return

    def find_class_winners(contestants: DataFrame) -> DataFrame:
        class_winners = contestants.sort_values('Score', ascending=False).iloc[:4]

        # replace numeric value with + signs for display  // TODO: could be written more pythonic
        for i, score in enumerate(class_winners[score_name]):
            str_score = str(score)
            if '.' in str_score:
                if str_score[::-1].index('.') == 3:
                    if str_score[-1] == '3':
                        plus_value = '+++'
                    elif str_score[-1] == '2':
                        plus_value = '++'
                    else:
                        plus_value = '+'
                    if int(str_score[4]) == 5:
                        class_winners[score_name].iloc[i] = str_score[:5] + plus_value
                    else:
                        class_winners[score_name].iloc[i] = str_score[:3] + plus_value

        return class_winners

    def find_award_winners(contestants: pandas.core.frame.DataFrame, break_tie: bool) -> pandas.core.frame.DataFrame:
        award_winners = contestants.sort_values('Score', ascending=False).iloc[:4]  # TODO: implement break tie
        # TODO: import award hierarchy to determine tiebreaking

        # replace numeric value with + signs for display  # TODO: could be written more pythonic
        for i, score in enumerate(award_winners[score_name]):
            str_score = str(score)
            if '.' in str_score:
                if str_score[::-1].index('.') == 3:
                    if str_score[-1] == '3':
                        plus_value = '+++'
                    elif str_score[-1] == '2':
                        plus_value = '++'
                    else:
                        plus_value = '+'
                    if int(str_score[4]) == 5:
                        award_winners[score_name].iloc[i] = str_score[:5] + plus_value
                    else:
                        award_winners[score_name].iloc[i] = str_score[:3] + plus_value

        return award_winners

    def write_winners(winners: pandas.core.frame.DataFrame) -> None:
        print(winners)

    # make sure input file is valid
    if input_file[-5:] == '.xlsx':
        data = pandas.read_excel(input_file)
    else:
        raise ValueError("only Excel files (.xlsx) are currently supported")

    class_name, score_name = find_variable_names(data)

    classes = data[class_name].unique()  # find set of classes

    # replace + signs with numeric value for sorting  // TODO: could be written more pythonic
    for i, score in enumerate(data[score_name]):
        plus_value = 0
        if type(score) is str and score[-1] == '+':
            while score[-1] == '+':
                plus_value += 0.001  # scores shouldn't be in increments of more than 0.5
                score = score[:-1]
            data[score_name][i] = float(score) + plus_value

    for class_ in classes:
        # skip nan class
        if pandas.isna(class_):
            continue

        # pull out contestants by class and who have scores
        contestants = data.loc[(data[class_name] == class_) &
                               ((data[score_name].apply(lambda x: isinstance(x, int))) |
                                (data[score_name].apply(lambda x: isinstance(x, float))))]
        winners = find_class_winners(contestants)

        write_winners(winners)

    # find High in Trial
    contestants = data.loc[(~data[class_name].isna()) & (data[class_name].str.contains(' B')) &
                               ((data[score_name].apply(lambda x: isinstance(x, int))) |
                                (data[score_name].apply(lambda x: isinstance(x, float))))]

    winners = find_class_winners(contestants)
    write_winners(winners)

    return True


if __name__ == '__main__':
    if len(sys.argv) > 3 or len(sys.argv) < 2:
        print("Usage: calculator.py [input_file] [output_file]")
        sys.exit()
    elif len(sys.argv) == 3:
        input_file = sys.argv[1]
        output_file = sys.argv[2]
    else:
        input_file = sys.argv[1]
        input_file_split = input_file.rsplit('.', 1)
        output_file = f"{input_file_split[0]}_scores.{input_file_split[1]}"

    calculate(input_file, output_file)

    print("Complete")
