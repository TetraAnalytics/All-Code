Attribute VB_Name = "CouponCalculations"
Sub calcBondFixedCpns(ByRef b As Bond)
    Dim i As Integer
    Dim freq As Integer
    Dim nextpay, lastPay, nextnextPay As Date
    
    freq = 1 / (b.freq / 12)

    i = 1
    nextpay = DateAdd("m", -freq, b.startDate)
    While nextpay <= b.valueDate
        nextpay = DateAdd("m", freq, nextpay)
    Wend
    lastPay = DateAdd("m", -freq, nextpay)
    nextnextPay = DateAdd("m", freq, nextpay)
    b.nextCouponDate = nextpay
 
    i = 1
    b.cashflow(i).Date = nextpay
    If b.dayCnt = "ACT/ACT" Then
        b.cashflow(i).couponCash = b.coupon * b.notional * (1 / b.freq)
        b.cashflow(i).totalCash = b.cashflow(i).totalCash + b.cashflow(i).couponCash
    Else
        b.cashflow(i).couponCash = b.coupon * b.notional * yearFrak(lastPay, nextpay, nextnextPay, b.dayCnt, b.freq) / b.freq 'first coupon accural
        b.cashflow(i).totalCash = b.cashflow(i).totalCash + b.cashflow(i).couponCash
    End If
    
    While nextpay <= DateAdd("m", -freq, b.maturityDate)
        i = i + 1
        lastPay = nextpay
        nextpay = DateAdd("m", freq, nextpay)
        nextnextPay = DateAdd("m", freq, nextpay)
        b.cashflow(i).Date = nextpay
        If b.dayCnt = "ACT/ACT" Then
            b.cashflow(i).couponCash = b.coupon * b.notional * (1 / b.freq)
            b.cashflow(i).totalCash = b.cashflow(i).totalCash + b.cashflow(i).couponCash
        Else
            b.cashflow(i).couponCash = b.coupon * b.notional * yearFrak(lastPay, nextpay, nextnextPay, b.dayCnt, b.freq) / b.freq
            b.cashflow(i).totalCash = b.cashflow(i).totalCash + b.cashflow(i).couponCash
        End If
    Wend

    b.cashflow(i).prinCash = b.notional
    b.cashflow(i).totalCash = b.cashflow(i).totalCash + b.notional

End Sub

