from BondDayCount import year_fraction


def bond_greeks(b):
    yr_fraction = year_fraction(b.value_date, b.start_date, b.next_coupon_date, b.day_count, b.freq)
    i = 1
    while b.cashflow[i].coupon_cash != 0:
        b.macaulay_duration += (i + yr_fraction) * b.cashflow[i].coupon_cash / (
                1 + b.yld / b.freq) ** i
        b.convexity += (1 / (1 + b.yld / b.freq) ** 2) * b.cashflow[i].coupon_cash / (
                1 + b.yld / b.freq) ** i * ((i + yr_fraction) ** 2 + i)
        i += 1

    # maturity cash flows
    i -= 1
    b.macaulay_duration += (i + yr_fraction) * b.notional_balance / (1 + b.yld / b.freq) ** i
    b.macaulay_duration /= (b.price * b.freq)

    b.convexity += (1 / (1 + b.yld / b.freq) ** 2) * b.notional_balance / (1 + b.yld / b.freq) ** i * (
            (i + yr_fraction) ** 2 + i + yr_fraction)
    b.convexity = b.convexity / b.price / b.freq ** 2 if b.notional_balance > 10 else 0

    b.modified_duration = b.macaulay_duration / (1 + b.yld / b.freq)

    print("Modified Duration = ", round(b.modified_duration, 2), " ", "Convexity = ", round(b.convexity,2))


def mortgage_greeks(m):
    w = 0.0

    for i in range(1, m.amortization_term):
        m.macaulay_duration += i * (
                m.cashflows[i].scheduled_interest +
                m.cashflows[i].scheduled_principal +
                m.cashflows[i].prepayment_principal +
                m.cashflows[i].default_principal -
                m.cashflows[i].default_loss
        ) / (1 + m.BEY / 12) ** (i + m.payment_delay / 30 + (30 - m.settle_day) / 30)

        p = m.cashflows[i].scheduled_principal + m.cashflows[i].prepayment_principal + m.cashflows[
            i].default_principal - \
            m.cashflows[i].default_loss

        w += p * (i + (m.payment_delay - 1) / 30 + (30 - min(30, m.settle_day) - 1) / 30 - 1) / (12 * m.balance)

        i += 1

    m.macaulay_duration = m.macaulay_duration / (m.price * m.balance) * 100 / 12
    m.modified_duration = m.macaulay_duration / (1 + m.BEY / 2)

    m.weighted_average_life = w

    return
