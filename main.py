from Classes import Mortgage
# from Read_write_csv import load_mortgage
# from Read_write_csv import print_mortgage_results

import pandas as pd


class Mortgages:
    def __init__(self, mortgage_term, loan_id, balance, note_rate):
        self.mortgage_term = mortgage_term
        self.loan_id = loan_id
        self.balance = balance
        self.note_rate = note_rate

        print("Balance = ", self.balance)
        print("loan id = ", self.loan_id)

        return


def main():

    csv_file_mortgage_attributes = '/Users/michaelaneiro/Downloads/mortgage_attributesv2.csv'  # MacBook
    df = pd.read_csv(csv_file_mortgage_attributes)
    # rows = len(df)
    print("Print list:")
    mortgage_list = df.iloc[1].tolist()
    for i in range(0, len(mortgage_list)):
        print(mortgage_list[i])
    print("Create m")
    m = Mortgages(mortgage_list[0], mortgage_list[1], mortgage_list[2], mortgage_list[3])
    # m = Mortgages('30-year', 'ABC123', 300000.00, 0.045)
    # print('mortgage_list: ',mortgage_list[0], " ", mortgage_list[1], " ", mortgage_list[2], " ", mortgage_list[3])
    print("Print mortgage list: ", mortgage_list)

    # my_objects = []
    # obj = pd.read_csv(csv_file_mortgage_attributes)
    # my_objects.append(obj)
    # print(my_objects)

    print("printing the dataframe:")
    for row in range(len(df)):
        print("df.loc[row].atloan_id = ", df.loc[row].at["loan_id"])  # or using df.
        print("df.iloc[row] = ", df.iloc[row])

main()
