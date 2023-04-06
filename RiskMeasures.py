import math


def bond_greeks(b):
    yr_fraction = year_fraction(b["valueDate"], b["startDate"], b["nextCouponDate"], b["dayCnt"], b["freq"])
    month = 1
    while b["cashflow"][month]["couponCash"] != 0:
        b["macaulay_duration"] += (month + yr_fraction) * b["cashflow"][month]["couponCash"] / (
                1 + b["yld"] / b["freq"]) ** month
        b["convexity"] += (1 / (1 + b["yld"] / b["freq"]) ** 2) * b["cashflow"][month]["couponCash"] / (
                1 + b["yld"] / b["freq"]) ** month * ((month + yr_fraction) ** 2 + month)
        month += 1

    # maturity cash flows
    month -= 1
    b["macaulay_duration"] += (month + yr_fraction) * b["notional"] / (1 + b["yld"] / b["freq"]) ** month
    b["macaulay_duration"] /= (b["price"] * b["freq"])

    b["convexity"] += (1 / (1 + b["yld"] / b["freq"]) ** 2) * b["notional"] / (1 + b["yld"] / b["freq"]) ** month * (
            (month + yr_fraction) ** 2 + month + yr_fraction)
    b["convexity"] = b["convexity"] / b["price"] / b["freq"] ** 2 if b["notional"] > 10 else 0

    b["modified_duration"] = b["macaulay_duration"] / (1 + b["yld"] / b["freq"])


def mortgage_greeks(m):
    month = 1
    wal = 0

    while m["cashflows"][int(month)]["scheduled_interest"] > 0 and month < 360:
        cashflow_term = m["cashflows"][month]["scheduled_interest"] + m["cashflows"][month]["scheduled_principal"] + \
                        m["cashflows"][month]["prepay_principal"] + m["cashflows"][month]["default_principal"] - \
                        m["cashflows"][month]["default_loss"]
        m["macaulay_duration"] += month * cashflow_term / (1 + m["BEY"] / 12) ** (
                month + m["delay"] / 30 + (30 - m["settle_day"]) / 30)
        principal_cash = m["cashflows"][month]["scheduled_principal"] + m["cashflows"][month]["prepay_principal"] + \
                         m["cashflows"][month]["default_principal"] - m["cashflows"][month]["default_loss"]
        wal += principal_cash * (month + (m["delay"] - 1) / 30 + (30 - min(30, m["settle_day"]) - 1) / 30 - 1) / (
                12 * m["balance"])
        month += 1

    m["macaulay_duration"] = m["macaulay_duration"] / (m["price"] * m["balance"]) * 100 / 12 + (
            ((m["delay"] + 1) - (m["settle_day"])) / 31) / 12
    m["modified_duration"] = m["macaulay_duration"] / (1 + m["BEY"] / 2)

    m["wal"] = wal

# After converting the Excel VBA code to Python, you can use the `bond_greeks` and `mortgage_greeks` functions
# by passing the respective bond and mortgage dictionaries as arguments.
# Please note that the `year_fraction` function
# and the required data structures for bond and mortgage are not provided in your code snippet.
# Make sure to define them properly before using the functions.
