Attribute VB_Name = "Classes"
Type BondCashflows
    prinCash        As Double
    couponCash      As Double
    totalCash       As Double
    Date            As Date
End Type

Type Bond
    type            As String
    id              As String
    notional        As Double
    coupon          As Double
    valueDate       As Date
    startDate       As Date
    nextCouponDate  As Date
    maturityDate    As Date
    maturity_months            As Integer
    busDay          As String
    dayCnt          As String
    freq            As Integer
  '  cashflow(1 To 360) As cashflows
    price           As Double
    bondMargin      As Double
    bondIndex       As String
    accrued         As Double
    yld             As Double
    spread          As Double
    curvePoint      As Double
    macDur          As Double
    modDur          As Double
    effDur          As Double
    convexity       As Double
    group           As String
    CostBasis       As Double
    baselType       As String
End Type

Type mcashFlows
    schedPay       As Double         'amortized payment due
    schedPrin      As Double        'Scheduled principal Payments
    interest       As Double           'Net_Rate interest payment
    prepayPrin     As Double      'principal paid form voluntary prepayments
    default        As Double            'principal amount that defaulted
    loss           As Double                'amount of loss per defaulted principal (default * severity)
End Type

Type mortgage
    mtgeType       As String
    id             As String
    balance        As Double
    leaseResidual  As Double
    Gross_Rate          As Double
    Net_Rate            As Double
    servicing_Fee      As Double
    maturity_months      As Integer
    balloon_term_months  As Integer
    loan_age       As Integer
    Interest_Only_Period As Integer
    index          As String
    margin         As Double
    InitialRateCap As Double
    PeriodicCap    As Double
    LifeCap        As Double
    LifeFloor      As Double
    initialResetPeriod  As Integer  ' Initial Months-to-Roll
    pReset         As Integer  ' Periodic Reset
    periodicity    As Integer
    SDday          As Integer
    daysInMonth    As Integer
    daycount       As String
    Delay          As Integer
    BEY            As Double
    mey            As Double
    price          As Double
    accrued        As Double
    cashflows(1 To 360) As mcashFlows
    macDur         As Double
    modDur         As Double
    WAL            As Double
    staticCPR      As Double
    staticCDR      As Double
    staticSeverity As Double
End Type
