from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from BondDayCount import year_fraction, dc_30_360, act_360, act_365, act_act

class Cashflow:
    def __init__(self):
        self.date = None
        self.coupon_cash = 0
        self.principal_cash = 0
        self.total_cash = 0


class Bond:
    def __init__(self, start_date, value_date, maturity_date, freq, day_count, floater_index, floater_spread, coupon, notional_balance, next_coupon_date, macaulay_duration, modified_duration, convexity, price, yld):
        self.start_date = datetime.strptime(start_date, '%m/%d/%Y')
        self.value_date = datetime.strptime(value_date, '%m/%d/%Y')
        self.maturity_date = datetime.strptime(maturity_date, '%m/%d/%Y')
        self.freq = freq
        self.day_count = day_count
        self.floater_index = floater_index
        self.floater_spread = floater_spread
        self.coupon = coupon
        self.notional_balance = notional_balance
        self.next_coupon_date = datetime.strptime(next_coupon_date, '%m/%d/%Y')
        self.macaulay_duration = macaulay_duration
        self.modified_duration = modified_duration
        self.convexity = convexity
        self.price = price
        self.yld = yld

        self.cashflow = [Cashflow() for _ in range(100)]


def calculate_bond_coupons(b):

    next_pay = b.start_date - relativedelta(months=12/b.freq)
    while next_pay <= b.value_date:
        next_pay += relativedelta(months=12/b.freq)

    last_pay = next_pay - relativedelta(months=12/b.freq)
    next_next_pay = next_pay + relativedelta(months=12/b.freq)
    b.next_coupon_date = next_pay

    i = 1
    b.cashflow[i].date = next_pay
    if b.day_count == "ACT/ACT":
        if b.floater_index == "Fixed":
            b.cashflow[i].coupon_cash = b.coupon * b.notional_balance * (1 / b.freq)
        else:
            b.cashflow[i].coupon_cash = (b.floater_index + b.floater_spread) * b.notional_balance * (1 / b.freq)
        b.cashflow[i].total_cash = b.cashflow[i].total_cash + b.cashflow[i].coupon_cash
    else:
        if b.floater_index == "Fixed":
            b.cashflow[i].coupon_cash = b.coupon * b.notional_balance * year_fraction(last_pay, next_pay, next_next_pay, b.day_count,
                                                                                  b.freq) / b.freq
        else:
            b.cashflow[i].coupon_cash = (b.floater_index + b.floater_spread) * b.notional_balance * year_fraction(last_pay, next_pay, next_next_pay,
                                                                                      b.day_count,
                                                                                      b.freq) / b.freq
        b.cashflow[i].total_cash = b.cashflow[i].total_cash + b.cashflow[i].coupon_cash

    while next_pay <= b.maturity_date - relativedelta(months=12/b.freq):
        i += 1
        last_pay = next_pay
        next_pay += relativedelta(months=12/b.freq)
        next_next_pay += relativedelta(months=12/b.freq)
        b.cashflow[i].date = next_pay
        if b.day_count == "ACT/ACT":
            b.cashflow[i].coupon_cash = b.coupon * b.notional_balance * (1 / b.freq)
            b.cashflow[i].total_cash = b.cashflow[i].total_cash + b.cashflow[i].coupon_cash
        else:
            b.cashflow[i].coupon_cash = b.coupon * b.notional_balance * year_fraction(last_pay, next_pay, next_next_pay, b.day_count,
                                                                                      b.freq) / b.freq
            b.cashflow[i].total_cash = b.cashflow[i].total_cash + b.cashflow[i].coupon_cash

    b.cashflow[i].principal_cash = b.notional_balance
    b.cashflow[i].total_cash = b.cashflow[i].total_cash + b.notional_balance

    return
