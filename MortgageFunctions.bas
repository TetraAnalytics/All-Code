Attribute VB_Name = "MortgageFunctions"
Function schedPay(bal, rate, maturity_months, residual)
    If maturity_months = 0 Then
        schedPay = 0#
    Else
        schedPay = (bal * rate / 12 * (1 + rate / 12) ^ maturity_months - residual * rate / 12) / ((1 + rate / 12) ^ maturity_months - 1)
    End If
End Function

Function prepayment(ByVal bal As Double, ByVal CPR As Double)
    prepayment = bal * (1 - (1 - CPR) ^ (1 / 12))
End Function

Function default(ByVal bal, ByVal CDR)
    default = bal * (1 - (1 - CDR) ^ (1 / 12))
End Function

Sub calcSingleMtgeCash(ByRef m As mortgage)
    Dim i As Integer
    Dim bal, Gross_Rate, startGross_Rate As Double
    Dim CPR, CDR, severity    As Double
    
    bal = m.balance
    Gross_Rate = m.Gross_Rate
    
    CDR = m.staticCDR
    CPR = m.staticCPR
    severity = 0.4
        
    'BEGIN LOOP
    For i = 1 To m.maturity_months
        m.cashflows(i).interest = bal * (Gross_Rate - m.servicing_Fee) / 12
        m.cashflows(i).schedPay = schedPay(bal, Gross_Rate, m.maturity_months - i + 1, 0)
        If (m.loan_age + i) < m.Interest_Only_Period Then
            m.cashflows(i).schedPay = m.cashflows(i).interest
        End If
        
        m.cashflows(i).schedPrin = m.cashflows(i).schedPay - bal * Gross_Rate / 12
        bal = bal - m.cashflows(i).schedPrin
        
        If i = m.balloon_term_months Then
            m.cashflows(i).prepayPrin = bal
        Else
            m.cashflows(i).prepayPrin = prepayment(bal, CPR)
        End If
        
        bal = bal - m.cashflows(i).prepayPrin
        
        m.cashflows(i).default = default(bal, CDR)
        m.cashflows(i).loss = m.cashflows(i).default * severity
        bal = bal - m.cashflows(i).default

    Next i

End Sub

Function CalcARMWAC(ByRef m As mortgage, ByVal indexRate As Double, ByVal mnth As Integer) As Double

    If (mnth + m.loan_age - 1) = m.initialResetPeriod Then
        CalcARMWAC = min(max(indexRate + m.margin, m.LifeFloor), m.LifeCap)
        CalcARMWAC = min(max(CalcARMWAC, m.Gross_Rate - m.InitialRateCap), m.Gross_Rate + m.InitialRateCap)
        m.Gross_Rate = CalcARMWAC
    Else
        If (mnth + m.loan_age - 1) > m.initialResetPeriod And modulus(mnth + m.loan_age - 1, m.pReset) = 0 Then
            CalcARMWAC = min(max(indexRate + m.margin, m.LifeFloor), m.LifeCap)
            CalcARMWAC = min(max(CalcARMWAC, m.Gross_Rate - m.PeriodicCap), m.Gross_Rate + m.PeriodicCap)
            m.Gross_Rate = CalcARMWAC
        Else
            CalcARMWAC = m.Gross_Rate
        End If
        
    End If
    
End Function


