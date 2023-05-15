def calculate_single_mtge_cash(m):
    balance = m.balance
    note_rate = m.note_rate

    cdr = m.staticCDR
    cpr = m.staticCPR

    print("Amort Term", m.amortization_term)

    for i in range(0, m.amortization_term):
        m.calculate_arm_wac(m, 0.05, i)
        m.cashflows[i].balance = balance
        m.cashflows[i].scheduled_interest = balance * (note_rate - m.servicing_fee) / 12
        m.cashflows[i].scheduled_payment = m.scheduled_payment(balance, note_rate, m.amortization_term - i, 0)

        m.cashflows[i].scheduled_principal = m.cashflows[i].scheduled_payment - balance * note_rate / 12
        balance -= m.cashflows[i].scheduled_principal

        m.cashflows[i].default_principal = m.default(balance, cdr)
        if m.LTV == 0:
            m.cashflows[i].default_loss = 0
        else:
            m.cashflows[i].default_loss = m.cashflows[i].default_principal * m.staticSeverity
        balance -= m.cashflows[i].default_principal

        if i == m.balloon_term:
            m.cashflows[i].prepayment_principal = balance
        else:
            m.cashflows[i].prepayment_principal = m.prepayment(balance, cpr)
        balance -= m.cashflows[i].prepayment_principal

    return


def calc_arm_wac(m, index_rate, current_month):
    if (current_month + m.loan_age - 1) == m.initialResetPeriod:
        arm_wac = min(max(index_rate + m.margin, m.LifeFloor), m.LifeCap)
        arm_wac = min(max(arm_wac, m.note_rate - m.InitialRateCap), m.note_rate + m.InitialRateCap)
        m.note_rate = arm_wac
    else:
        if (current_month + m.loan_age - 1) > m.initialResetPeriod and (current_month + m.loan_age - 1) % m.pReset == 0:
            arm_wac = min(max(index_rate + m.margin, m.LifeFloor), m.LifeCap)
            arm_wac = min(max(arm_wac, m.note_rate - m.PeriodicCap), m.note_rate + m.PeriodicCap)
            m.note_rate = arm_wac
        else:
            arm_wac = m.note_rate

    return arm_wac
