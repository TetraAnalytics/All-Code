import os
from os.path import exists
import csv


def print_mortgage_cashflows_csv(m):
    # Open a new CSV file in write mode
    mortgage_results = '/Users/michaelaneiro/Downloads/mortgage_results_out.csv'  # MacBook

    # Delete the previous output file
    if exists(mortgage_results):
        os.remove(mortgage_results)

    with open(mortgage_results, 'w', newline='') as csvfile:
        # Create a CSV writer object
        csvwriter = csv.writer(csvfile)
        # Write the header row
        csvwriter.writerow(['Balance', 'Scheduled Pay', 'Scheduled Prin', 'Scheduled Interest', 'Prepay', 'Default', 'Loss'])
        # Write each MortgageCashflows object to a row in the CSV file
        for i in range(m.amortization_term):
            csvwriter.writerow([m.cashflows[i].balance, m.cashflows[i].scheduled_payment, m.cashflows[i].scheduled_principal,  m.cashflows[i].scheduled_interest, m.cashflows[i].prepayment_principal, m.cashflows[i].default_principal, m.cashflows[i].default_loss])

    return
