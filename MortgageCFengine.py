class MortgageCashflows:
    def __int__(self):
        self.scheduled_payment = 0
        self.scheduled_principal = 0
        self.scheduled_interest = 0
        self.prepayment_principal = 0
        self.default_principal = 0
        self.default_loss = 0


class Mortgage:
    def __init__(self):
        self.mtgeType = ''
        self.id = ''
        self.balance = 0.0
        self.note_rate = 0.0
        self.Net = 0.0
        self.servicing = 0.0
        self.amortization_term = 0
        self.balloon_term = 0
        self.loan_age = 0
        self.initialResetPeriod = 0
        self.margin = 0
        self.LifeFloor = 0
        self.LifeCap = 0
        self.Gross_Rate = 0
        self.InitialRateCap = 0
        self.pReset = 0
        self.PeriodicCap = 0
        self.settle_day = 0
        self.days_in_month = 0
        self.payment_delay = 0
        self.BEY = 0.0
        self.mey = 0.0
        self.price = 0.0
        self.accrued_interest = 0.0
        self.cashflows = [MortgageCashflows() for _ in range(360)]
        self.macDur = 0.0
        self.modDur = 0.0
        self.effDur = 0.0
        self.WAL = 0.0
        self.convexity = 0.0
        self.staticCPR = 0.0
        self.staticCDR = 0.0
        self.staticSeverity = 0.0
        self.LTV = 0.0


def scheduled_payment(balance, note_rate, amortization_term, residual):
    if amortization_term == 0:
        return 0.0
    else:
        return (balance * note_rate / 12 * (1 + note_rate / 12) ** amortization_term - residual * note_rate / 12) / (
                    (1 + note_rate / 12) ** amortization_term - 1)


def prepayment(balance, scheduled_principal, default_principal, cpr):
    return (balance - scheduled_principal - default_principal) * (1 - (1 - cpr) ** (1 / 12))


def default(balance, scheduled_principal, cdr):
    return (balance - scheduled_principal) * (1 - (1 - cdr) ** (1 / 12))


def severity_calculation(ltv):
    if ltv >= 0.7:
        return 0.6
    elif ltv >= 0.6:
        return 0.4
    else:
        return 0.2


def calc_arm_wac(m, index_rate, month):
    if (month + m.loan_age - 1) == m.initialResetPeriod:
        arm_wac = min(max(index_rate + m.margin, m.LifeFloor), m.LifeCap)
        arm_wac = min(max(arm_wac, m.note_rate - m.InitialRateCap), m.note_rate + m.InitialRateCap)
        m.note_rate = arm_wac
    else:
        if (month + m.loan_age - 1) > m.initialResetPeriod and (month + m.loan_age - 1) % m.pReset == 0:
            arm_wac = min(max(index_rate + m.margin, m.LifeFloor), m.LifeCap)
            arm_wac = min(max(arm_wac, m.note_rate - m.PeriodicCap), m.note_rate + m.PeriodicCap)
            m.note_rate = arm_wac
        else:
            arm_wac = m.note_rate

    return arm_wac


def calculate_single_mtge_cash(m):
    balance = m.balance
    note_rate = m.note_rate

    cdr = m.staticCDR
    cpr = m.staticCPR

    for i in range(1, m.amortization_term + 1):
        m.cashflows[i].interest = balance * (note_rate - m.servicing) / 12
        m.cashflows[i].scheduled_payment = scheduled_payment(balance, note_rate, m.amortization_term - i + 1, 0)
        if (m.loan_age + i) < m.IOPeriod:
            m.cashflows[i].scheduled_payment = m.cashflows[i].interest

        m.cashflows[i].scheduled_principal = m.cashflows[i].scheduled_payment - balance * note_rate / 12
        balance -= m.cashflows[i].scheduled_principal

        m.cashflows[i].default_principal = default(balance, 0, cdr)
        if m.LTV == 0:
            m.cashflows[i].default_loss = 0
        else:
            m.cashflows[i].default_loss = m.cashflows[i].default_principal * severity_calculation(m.LTV)
        balance -= m.cashflows[i].default_principal

        if i == m.balloon_term:
            m.cashflows[i].prepayment_principal = balance
        else:
            m.cashflows[i].prepayment_principal = prepayment(balance, 0, 0, cpr)
        balance -= m.cashflows[i].prepayment_principal
        m.cashflows(i).default_principal = default(balance, 0, cdr)

        if m.LTV == 0:
            m.cashflows(i).default_loss = 0
        else:
            m.cashflows(i).default_loss = m.cashflows(i).default_principal * severity_calculation(m.LTV)
        balance = balance - m.cashflows(i).default_principal

        if i == m.balloon_term:
            m.cashflows(i).prepayment_principal = balance

        m.cashflows(i).prepayment_principal = prepayment(balance, 0, 0, cpr)
        balance = balance - m.cashflows(i).prepayment_principal

    return
