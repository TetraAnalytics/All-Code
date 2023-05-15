
class MortgageCashflows:
    def __int__(self):
        self.id = ''
        self.balance = 0.0
        self.scheduled_payment = 0.0
        self.scheduled_principal = 0.0
        self.scheduled_interest = 0.0
        self.prepayment_principal = 0.0
        self.default_principal = 0.0
        self.default_loss = 0.0


class Mortgage:
    def __init__(self, mortgage_term, loan_id, balance, note_rate, servicing_fee, amortization_term, balloon_term, loan_age, initialResetPeriod, margin, LifeFloor, LifeCap, InitialRateCap, pReset, PeriodicCap, settle_day, days_in_month, payment_delay, BEY, price, accrued_interest, staticCPR, staticCDR, staticSeverity, LTV):
        self.mortgage_term = mortgage_term
        self.loan_id = loan_id
        self.balance = balance
        self.note_rate = note_rate
        self.servicing_fee = servicing_fee
        self.amortization_term = amortization_term
        self.balloon_term = balloon_term
        self.loan_age = loan_age
        self.initialResetPeriod = initialResetPeriod
        self.margin = margin
        self.LifeFloor = LifeFloor
        self.LifeCap = LifeCap
        self.InitialRateCap = InitialRateCap
        self.pReset = pReset
        self.PeriodicCap = PeriodicCap
        self.settle_day = settle_day
        self.days_in_month = days_in_month  # Remove. calculate from settle_day
        self.payment_delay = payment_delay
        self.BEY = BEY
        self.price = price
        self.accrued_interest = accrued_interest
        self.macaulay_duration = 0.0  # Remove.  Output only
        self.modified_duration = 0.0  # Remove.  Output only
        self.weighted_average_life = 0.0  # Remove.  Output only
        self.convexity = 0.0  # Remove
        self.staticCPR = staticCPR
        self.staticCDR = staticCDR
        self.staticSeverity = staticSeverity
        self.LTV = LTV

        self.cashflows = [MortgageCashflows() for _ in range(360)]

    @staticmethod
    def scheduled_payment(balance, note_rate, amortization_term, residual):
        if amortization_term == 0:
            return 0.0
        else:
            return (balance * note_rate / 12 * (1 + note_rate / 12) ** amortization_term - residual * note_rate / 12) / (
                        (1 + note_rate / 12) ** amortization_term - 1)

    @staticmethod
    def prepayment(balance, staticCPR):
        return balance * (1 - (1 - staticCPR) ** (1 / 12))

    @staticmethod
    def default(balance, staticCDR):
        return balance * (1 - (1 - staticCDR) ** (1 / 12))

    @staticmethod
    def calculate_arm_wac(m, index_rate, month):
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
        # print(month, ": ARM WAC: ", m.note_rate - m.servicing_fee)

        return arm_wac
