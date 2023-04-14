Attribute VB_Name = "PVandYields"
'PV and Yields
Option Explicit
Option Compare Text

Function mey(BEY) As Double   ' convert bond equivalent yield to a monthly yield to discount cash flows
  mey = 12# * ((1 + BEY / 2) ^ (1# / 6#) - 1#) '
End Function

Function convertFreq(freq As String) As Integer
   Select Case LCase(freq)
        Case "monthly":  convertFreq = 12
        Case "quarterly": convertFreq = 4
        Case "semi-annual": convertFreq = 2
        Case "annual": convertFreq = 1
    End Select
End Function

Function accrued(ByRef dated As Date, ByRef sd As Date, ByRef nxtCpn As Date, ByRef daycount As String, ByRef cpn As Double, ByVal freq As Double) As Double
    If daycount = "ACT/ACT" Then
            accrued = -(sd - dated) / (freq * (nxtCpn - dated)) * cpn
    Else: accrued = -yearFrak(dated, sd, nxtCpn, daycount, freq) * cpn / freq
    End If
End Function

Function PresVal(cash, mey_, month, Delay, daysInMonth, SD_day) As Double
    PresVal = cash / (1 + (mey_ / 12)) ^ (month + (Delay - 1) / 30 + (daysInMonth - SD_day + 1) / 30 - 1)
End Function

Function PVMtgCashFlows(ByRef m As mortgage, ByVal yld As Double) As Double
    Dim i As Integer
    Dim cumCash, mey_, PVcf As Double
    PVcf = 0

    mey_ = 12 * ((1 + yld / 2) ^ (2 / 12) - 1)
    
    For i = 1 To min(m.maturity_months, m.balloon_term_months)
        Select Case m.mtgeType
            Case "IO":  cumCash = m.cashflows(i).interest
            Case "PO": cumCash = m.cashflows(i).schedPrin + m.cashflows(i).prepayPrin + (m.cashflows(i).default - m.cashflows(i).loss)
            Case Else:  cumCash = m.cashflows(i).schedPrin + m.cashflows(i).interest + m.cashflows(i).prepayPrin + (m.cashflows(i).default - m.cashflows(i).loss)
        End Select
        PVcf = PVcf + PresVal(cumCash, mey_, i, m.Delay, m.daysInMonth, m.SDday)
    Next i
    
    PVMtgCashFlows = PVcf
    
End Function

Function CalcMtgYield(ByRef m As mortgage) As Double
' Iterate to calculate a yield based on a price and an array of cashflows
    Dim yld1, Px1, yld2, Px2, Slope, cash, price_diff, target_price As Double
    Dim i As Integer
    i = 1
 
    target_price = m.price '+ m.accrued
    yld1 = 0.05
    Px1 = PVMtgCashFlows(m, 0.05) / m.balance * 100
    yld2 = yld1 - 0.005
    Px2 = PVMtgCashFlows(m, yld2) / m.balance * 100
    Slope = (Px1 - Px2) / (yld1 - yld2) / -1
    price_diff = (target_price - Px1)
  
    While Abs(price_diff) > 1E-08
        yld1 = yld1 - (target_price - Px1) / Slope
        Px1 = PVMtgCashFlows(m, yld1) / m.balance * 100
        
        yld2 = yld1 - 0.0001
        Px2 = PVMtgCashFlows(m, yld2) / m.balance * 100
        
        Slope = (Px1 - Px2) / (yld1 - yld2) / -1
        price_diff = (target_price - Px1)
        
        i = i + 1
        If i > 10 Then
            price_diff = 0
        End If
    Wend
 
    CalcMtgYield = yld1
 
End Function

Function pvBond(ByRef b As Bond, ByVal yld As Double) As Double
'Present value an array of cashflows on individual bonds
 
    Dim presValue, discountmaturity_months, cash As Double
    Dim i As Integer
    Dim freq As Integer
    Dim nextpay, nextnextPay As Date
    
    freq = 1 / (b.freq / 12)
    presValue = 0
    nextpay = b.startDate
    
    i = 1
    While nextpay < b.maturityDate
        nextpay = DateAdd("m", freq, nextpay)
        If b.dayCnt = "ACT/ACT" Then
                discountmaturity_months = i - 1 + yearFrak(b.valueDate, b.startDate, nextnextPay, b.dayCnt, b.freq)
        Else: discountmaturity_months = yearFrak(b.valueDate, nextpay, nextnextPay, b.dayCnt, b.freq)
        End If
        presValue = presValue + presentValue(b.cashflow(i).couponCash + b.cashflow(i).prinCash, yld, b.freq, discountmaturity_months)
        i = i + 1
    Wend
    
    pvBond = presValue
 
End Function

Function presentValue(ByVal cash As Double, ByVal yld As Double, ByVal freq As Integer, ByVal discountmaturity_months As Double) As Double
    presentValue = (cash) / (1 + yld / freq) ^ discountmaturity_months
End Function

Function CalcBndYield(ByRef b As Bond) As Double  ' Iterate to calculate a yield based on a price and an array of cashflows
    'For actual portfolio bonds
    Dim yld1, Px1, yld2, Px2, Slope, cash, price_diff, target_price As Double
    Dim i As Integer
    i = 1
    
    target_price = b.price + b.accrued

    yld1 = 0.035
    Px1 = pvBond(b, yld1)
    
    yld2 = yld1 - 0.001
    Px2 = pvBond(b, yld2)
    
    Slope = (Px1 - Px2) / (yld1 - yld2) / -1
    price_diff = (target_price - Px1)
  
    While Abs(price_diff) > 1E-05
        yld1 = yld1 - (target_price - Px1) / Slope
        Px1 = pvBond(b, yld1)
        
        yld2 = yld1 - 0.001
        Px2 = pvBond(b, yld2)

        Slope = (Px1 - Px2) / (yld1 - yld2) / -1
        price_diff = (target_price - Px1)
        
        i = i + 1
        If i > 10 Then
            price_diff = 0
        End If
    Wend
 
 CalcBndYield = yld1
 
End Function
