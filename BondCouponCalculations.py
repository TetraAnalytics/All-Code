from datetime import timedelta


class Cashflow:
    def __init__(self):
        self.date = None
        self.coupon_cash = 0
        self.principal_cash = 0
        self.total_cash = 0


class Bond:
    def __init__(self, start_date, value_date, maturity_date, freq, day_cnt, coupon, notional):
        self.start_date = start_date
        self.value_date = value_date
        self.maturity_date = maturity_date
        self.freq = freq
        self.day_count = day_count
        self.coupon = coupon
        self.notional_balance = notional_balance
        self.next_coupon_date = None
        self.cashflow = [Cashflow() for _ in range(1000)]


def calc_bond_fixed_cpns(b):
    freq = 1 // (b.freq / 12)

    i = 1
    next_pay = b.start_date - timedelta(days=freq * 30)
    while next_pay <= b.value_date:
        next_pay += timedelta(days=freq * 30)

    last_pay = next_pay - timedelta(days=freq * 30)
    next_next_pay = next_pay + timedelta(days=freq * 30)
    b.next_coupon_date = next_pay

    i = 1
    b.cashflow[i].date = next_pay
    if b.day_cnt == "ACT/ACT":
        b.cashflow[i].coupon_cash = b.coupon * b.notional_balance * (1 / b.freq)
        b.cashflow[i].total_cash = b.cashflow[i].total_cash + b.cashflow[i].coupon_cash
    else:
        b.cashflow[i].coupon_cash = b.coupon * b.notional_balance * year_fraction(last_pay, next_pay, next_next_pay, b.day_cnt,
                                                                                  b.freq) / b.freq
        b.cashflow[i].total_cash = b.cashflow[i].total_cash + b.cashflow[i].coupon_cash

    while next_pay <= b.maturity_date - timedelta(days=freq * 30):
        i += 1
        last_pay = next_pay
        next_pay += timedelta(days=freq * 30)
        next_next_pay += timedelta(days=freq * 30)
        b.cashflow[i].date = next_pay
        if b.day_cnt == "ACT/ACT":
            b.cashflow[i].coupon_cash = b.coupon * b.notional_balance * (1 / b.freq)
            b.cashflow[i].total_cash = b.cashflow[i].total_cash + b.cashflow[i].coupon_cash
        else:
            b.cashflow[i].coupon_cash = b.coupon * b.notional_balance * year_fraction(last_pay, next_pay, next_next_pay, b.day_cnt,
                                                                                      b.freq) / b.freq
            b.cashflow[i].total_cash = b.cashflow[i].total_cash + b.cashflow[i].coupon_cash

    b.cashflow[i].principal_cash = b.notional_balance
    b.cashflow[i].total_cash = b.cashflow[i].total_cash + b.notional_balance
