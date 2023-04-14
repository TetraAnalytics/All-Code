Attribute VB_Name = "Greeks"
'Greeks
Option Explicit
Option Compare Text

Sub bondGreeks(ByRef b As Bond)
Dim month, n As Integer
Dim yr_fraction As Double

    yr_fraction = year_fraction(b.valueDate, b.startDate, b.nextCouponDate, b.dayCnt, b.freq)
    month = 1
    While b.cashflow(t).couponCash <> 0
        b.macauley_duration = b.macauley_duration + (month + yr_fraction) * b.cashflow(t).couponCash / (1 + b.yld / b.freq) ^ month
        b.convexity = b.convexity + (1 / (1 + b.yld / b.freq) ^ 2) * b.cashflow(t).couponCash / (1 + b.yld / b.freq) ^ t * ((month + yr_fraction) ^ 2 + month)
        month = month + 1
    Wend

    ' maturity cash flows
     month = month - 1
     b.macauley_duration = b.macauley_duration + (month + yr_fraction) * b.notional / (1 + b.yld / b.freq) ^ month
     b.macauley_duration = b.macauley_duration / b.price / b.freq
    
     b.convexity = b.convexity + (1 / (1 + b.yld / b.freq) ^ 2) * b.notional / (1 + b.yld / b.freq) ^ month * ((month + yr_fraction) ^ 2 + month + yr_fraction)
     If b.notional > 10 Then
        b.convexity = b.convexity / b.price / b.freq ^ 2  ' convexity adjustment
        Else: b.convexity = 0
    End If
  
    b.modDur = b.macauley_duration / (1 + b.yld / b.freq) ' modified duration
    
End Sub

Sub mortgageGreeks(ByRef m As mortgage)
    Dim t, n As Integer
    Dim w, p As Double

    m.macauley_duration = 0

    t = 1
    While m.cashflows(Int(t)).interest > 0# And t < 360
        m.macauley_duration = m.macauley_duration + t * (m.cashflows(t).interest + m.cashflows(t).scheduled_Principal + m.cashflows(t).prepayPrin + m.cashflows(t).default - m.cashflows(t).loss) / (1 + m.BEY / 12) ^ (t + m.Delay / 30 + (30 - m.SDday) / 30)
        p = m.cashflows(t).scheduled_Principal + m.cashflows(t).prepayPrin + m.cashflows(t).default - m.cashflows(t).loss   ' for WAL
        w = w + p * (t + (m.Delay - 1) / 30 + (30 - min(30, m.SDday) - 1) / 30 - 1) / (12 * m.balance)  ' for WAL
        t = t + 1
    Wend

    m.macauley_duration = m.macauley_duration / (m.price * m.balance) * 100 / 12 + (((m.Delay + 1) - (m.SDday)) / 31) / 12
    m.modDur = m.macauley_duration / (1 + m.BEY / 2)
    
    m.WAL = w

End Sub
