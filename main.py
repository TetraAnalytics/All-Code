import pandas as pd
from PVandYIELDS import pv_bond
from BondCouponCalculations import Bond, calculate_bond_coupons
from Classes import Mortgage
from MortgageCFengine import calculate_single_mtge_cash
from PVandYIELDS import pv_mtg_cash_flows, calculate_mtg_yield
from Read_write_csv import print_mortgage_cashflows_csv
from Read_write_xlsx import print_mortgage_cashflows_xlsx, print_mortgage_results
from RiskMeasures import mortgage_greeks, bond_greeks


def main():
    # MORTGAGES:
    # Define a mortgage cashflow class object to aggregate all loans' cashflows
    # aggregate = MortgageCashflows()
    # define file name and location of the loan data to import
    csv_file_mortgage_attributes = '/Users/michaelaneiro/Downloads/mortgage_attributes.csv'  # MacBook
    # use pandas function to read csv into a dataframe
    df = pd.read_csv(csv_file_mortgage_attributes)

    # Iterate through the rows of the dataframe into a list variable to run mortgage analytics
    rows = len(df)

    for r in range(rows):
        ml = df.iloc[r].tolist()  # mortgage_list row number count starts at 0
        # create an object of a class 'Mortgage', passing in the variables defining the mortgage
        m = Mortgage(ml[0], ml[1], ml[2], ml[3], ml[4], ml[5], ml[6], ml[7], ml[8], ml[9], ml[10], ml[11], ml[12],
                     ml[13], ml[14], ml[15], ml[16], ml[17], ml[18], ml[19], ml[20], ml[21], ml[22], ml[23], ml[24])
        calculate_single_mtge_cash(m)
        m.accrued_interest = (m.note_rate - m.servicing_fee) / 12 * (m.settle_day - 1) / 30 * 100
        m.price = pv_mtg_cash_flows(m, m.BEY) / m.balance * 100
        print("Mortgage price: ", m.price)
        m.BEY = calculate_mtg_yield(m)
        print("Mortgage yield: ", m.BEY)
        print_mortgage_cashflows_xlsx(m)
        # print_mortgage_cashflows_csv(m)
        print_mortgage_results(m)
        mortgage_greeks(m)

        b = Bond('5/15/2015', '5/15/2015', '5/15/2025', 2, "ACT/ACT", "Fixed", 0.00, 0.05, 100.00, '11/15/2015', 0.0,
                 0.0, 0.0, 100, 0.05)
        calculate_bond_coupons(b)
        print("Bond PV at 5% yield = ", round(pv_bond(b, 0.05), 3))
        bond_greeks(b)

    return


main()
