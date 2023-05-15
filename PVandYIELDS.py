from BondDayCount import year_fraction
from dateutil.relativedelta import relativedelta


def convert_to_mey(bey):
    return 12 * ((1 + bey / 2) ** (1 / 6) - 1)


def convert_freq(freq):
    freq = freq.lower()
    if freq == "monthly":
        return 12
    elif freq == "quarterly":
        return 4
    elif freq == "semi-annual":
        return 2
    elif freq == "annual":
        return 1


def accrued_interest(dated, settle_date, next_coupon, day_count, coupon, frequency):
    if day_count == "ACT/ACT":
        return -(settle_date - dated) / (frequency * (next_coupon - dated)) * coupon
    else:
        return -year_fraction(dated, settle_date, next_coupon, day_count, frequency) * coupon / frequency


def bond_present_value(cash, bey, month, delay, days_in_month, sd_day):
    return cash / (1 + (bey / 12)) ** (month + (delay - 1) / 30 + (days_in_month - sd_day + 1) / 30 - 1)


def pv_mtg_cash_flows(m, yld):
    pv_cf = 0
    mey = convert_to_mey(yld)

    for i in range(0, min(m.amortization_term, m.balloon_term)):
        cumulative_cash = m.cashflows[i].scheduled_principal + m.cashflows[i].scheduled_interest + m.cashflows[i].prepayment_principal + (
                    m.cashflows[i].default_principal - m.cashflows[i].default_loss)
        pv_cf += bond_present_value(cumulative_cash, mey, i + 1, m.payment_delay, m.days_in_month, m.settle_day)

    return pv_cf


def calculate_mtg_yield(m):
    i = 1
    target_price = m.price
    yld1 = 0.05
    pv1 = pv_mtg_cash_flows(m, 0.05) / m.balance * 100
    yld2 = yld1 - 0.005
    pv2 = pv_mtg_cash_flows(m, yld2) / m.balance * 100
    slope = (pv1 - pv2) / (yld1 - yld2) / -1
    price_difference = (target_price - pv1)

    while abs(price_difference) > 1e-08:
        yld1 = yld1 - (target_price - pv1) / slope
        pv1 = pv_mtg_cash_flows(m, yld1) / m.balance * 100

        yld2 = yld1 - 0.0001
        pv2 = pv_mtg_cash_flows(m, yld2) / m.balance * 100

        slope = (pv1 - pv2) / (yld1 - yld2) / -1
        price_difference = (target_price - pv1)

        i += 1
        if i > 10:
            price_difference = 0

    return yld1


# You'll need to replace the `Bond` class with the Python version you've already converted.
def pv_bond(b, yld):
    bond_pv = 0.0
    nextpay = b.start_date
    next_next_pay = b.start_date

    i = 1
    while nextpay < b.maturity_date:
        nextpay = nextpay + relativedelta(months=12/b.freq)
        if b.day_count == "ACT/ACT":
            discount_maturity_months = i - 1 + year_fraction(b.value_date, b.start_date, next_next_pay, b.day_count, b.freq)
        else:
            discount_maturity_months = year_fraction(b.value_date, nextpay, next_next_pay, b.day_count, b.freq)

        bond_pv += present_value(b.cashflow[i].coupon_cash + b.cashflow[i].principal_cash, yld, b.freq, discount_maturity_months)
        i += 1
    return bond_pv


def present_value(cash, yld, freq, discount_maturity_months):
    return cash / (1 + yld / freq) ** discount_maturity_months


# You'll need to replace the `Bond` class with the Python version you've already converted.
def calc_bnd_yield(b):
    i = 1
    target_price = b.price + b.accrued

    yld1 = 0.035
    px1 = pv_bond(b, yld1)

    yld2 = yld1 - 0.001
    px2 = pv_bond(b, yld2)

    slope = (px1 - px2) / (yld1 - yld2) / -1
    price_diff = (target_price - px1)

    while abs(price_diff) > 1e-05:
        yld1 = yld1 - (target_price - px1) / slope
        px1 = pv_bond(b, yld1)

        yld2 = yld1 - 0.001
        px2 = pv_bond(b, yld2)

        slope = (px1 - px2) / (yld1 - yld2) / -1
        price_diff = (target_price - px1)

        i += 1
        if i > 10:
            price_diff = 0

    return yld1
