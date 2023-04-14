Attribute VB_Name = "ReadWrite"
'Read and Write
Option Explicit
Option Compare Text

Sub LoadBondData(ByRef b As Bond)
    Dim i As Integer
    Dim freq As Double
    Dim firstPay As Date
    
    b.notional = 100 'Sheets("BondCalc").Cells(3, 3).value
    b.bondIndex = Sheets("BondCalc").Range("D9").Value
    b.bondMargin = Sheets("BondCalc").Range("D10").Value
    b.coupon = max(Sheets("BondCalc").Range("G6").Value, 1E-12)
    b.valueDate = Sheets("BondCalc").Range("D5").Value  'settlement date
    b.maturityDate = min(DateAdd("yyyy", 30, b.valueDate), Sheets("BondCalc").Range("D6").Value)
    b.dayCnt = Sheets("BondCalc").Range("G8").Value
    b.freq = convertFreq(Sheets("BondCalc").Range("G7").Value)
    b.yld = Sheets("BondCalc").Range("D15").Value
    b.spread = Sheets("BondCalc").Range("D14").Value / 100 / 100
    b.curvePoint = NumberofMonths(b.maturityDate, b.valueDate)
    freq = 1 / (b.freq / 12)  'converts to number of months between resets
        firstPay = b.maturityDate
        While firstPay > b.valueDate
            firstPay = DateAdd("m", -freq, firstPay)
        Wend
    b.startDate = firstPay
    b.price = Sheets("BondCalc").Range("D16").Value

End Sub

Sub printBondResults(ByRef b As Bond)
    'prints cash flows and prices for the fixed bonds
    Dim accruedInt As Double
    Dim presValue, discount As Double
    Dim i As Integer
    Dim freq As Integer
    Dim nextpay, nextnextPay As Date
            
    Sheets("BondCashFlows").Cells(1, 1).Value = "Bond Payment"
    Sheets("BondCashFlows").Cells(1, 2).Value = "Payment Date"
    Sheets("BondCashFlows").Cells(1, 3).Value = "Discount Period"
    Sheets("BondCashFlows").Cells(1, 4).Value = "Coupon"
    Sheets("BondCashFlows").Cells(1, 5).Value = "Discount Rate"
    Sheets("BondCashFlows").Cells(2, 6).Value = "Discount CF"
    
    freq = 1 / (b.freq / 12)
    nextpay = b.startDate
    
    i = 1
    While nextpay < b.maturityDate And i <= 360
        nextpay = DateAdd("m", freq, nextpay)
        If b.dayCnt = "ACT/ACT" Then
                discount = i - 1 + yearFrak(b.valueDate, b.startDate, nextnextPay, b.dayCnt, b.freq)
        Else: discount = yearFrak(b.valueDate, nextpay, nextnextPay, b.dayCnt, b.freq)
        End If
     '   presValue = presValue + b.cashflow(i).couponCash / (1 + b.yld / b.freq) ^ discount
            Sheets("BondCashFlows").Cells(i + 1, 1).Value = i
            Sheets("BondCashFlows").Cells(i + 1, 2).Value = nextpay
            Sheets("BondCashFlows").Cells(i + 1, 3).Value = discount
     '       If b.cashflow(i).couponCash > 0.01 Then
  '              Sheets("BondCashFlows").Cells(i + 1, 4).Value = b.cashflow(i).couponCash
            Else
                Sheets("BondCashFlows").Cells(i + 1, 4).Value = 0
            End If
            Sheets("BondCashFlows").Cells(i + 1, 5).Value = b.yld
  '          Sheets("BondCashFlows").Cells(i + 1, 6).Value = b.cashflow(i).couponCash / (1 + b.yld / b.freq) ^ discount
        i = i + 1
    Wend
    
    i = i - 1
    If b.cashflow(i).couponCash > 0.01 Then
        Sheets("BondCashFlows").Cells(i + 1, 4).Value = b.cashflow(i).couponCash + b.notional
    Else
        Sheets("BondCashFlows").Cells(i + 1, 4).Value = b.notional
    End If
    Sheets("BondCashFlows").Cells(i + 1, 6).Value = Sheets("BondCashFlows").Cells(i + 1, 6).Value + b.notional / (1 + b.yld / b.freq) ^ discount
    
    Sheets("bondcalc").Range("D8").Value = b.nextCouponDate
    Sheets("bondCalc").Range("G15").Value = b.macDur
    Sheets("bondCalc").Range("G16").Value = b.modDur
    Sheets("bondCalc").Range("G17").Value = b.convexity
    Sheets("bondCalc").Range("G18").Value = 0
    
End Sub

Sub loadMortgage(ByRef m As mortgage)
Dim d As Date

    m.mtgeType = Sheets("MortgageCalc").Range("C4").Value
    m.balance = Sheets("MortgageCalc").Range("C6").Value
    m.Gross_Rate = Sheets("MortgageCalc").Range("c7").Value
    m.Net_Rate = Sheets("MortgageCalc").Range("c8").Value
    m.servicing_Fee = m.Gross_Rate - m.Net_Rate
    m.maturity_months = Sheets("MortgageCalc").Range("c9").Value
    m.balloon_term_months = Sheets("MortgageCalc").Range("c10").Value
    If m.balloon_term_months = 0 Then
        m.balloon_term_months = m.maturity_months
    End If
    m.loan_age = Sheets("MortgageCalc").Range("c11").Value
    m.Delay = Sheets("MortgageCalc").Range("c12").Value
    m.Interest_Only_Period = Sheets("MortgageCalc").Range("c13").Value
    d = Sheets("MortgageCalc").Range("L5").Value  ' pull in settlement date for later conversion
    m.daysInMonth = Day(DateSerial(Year(d), month(d) + 1, 1) - 1)
    m.SDday = Day(d)
    m.BEY = Sheets("MortgageCalc").Range("i10").Value
    m.mey = 12 * ((1 + m.BEY / 2) ^ (2 / 12) - 1)
    m.price = Sheets("MortgageCalc").Range("i6").Value + m.Net_Rate / 12 * (m.SDday - 1) / 30 * 100
    m.staticCPR = Sheets("MortgageCalc").Range("f6").Value
    m.staticCDR = Sheets("MortgageCalc").Range("f7").Value

End Sub

Sub printMortgageCashFlows(ByRef m As mortgage)
    Dim i As Integer
    Dim balance As Double

    Worksheets("MortgageCashFlows").Range("a3:k370").ClearContents
    balance = m.balance

    Sheets("MortgageCashFlows").Cells(2, 1).Value = "Month"
    Sheets("MortgageCashFlows").Cells(2, 2).Value = "Balance"
    Sheets("MortgageCashFlows").Cells(2, 3).Value = "Sched Pay"
    Sheets("MortgageCashFlows").Cells(2, 4).Value = "PT Interest"
    Sheets("MortgageCashFlows").Cells(2, 5).Value = "Sched Prin"
    Sheets("MortgageCashFlows").Cells(2, 6).Value = "Prepay Prin"
    Sheets("MortgageCashFlows").Cells(2, 7).Value = "Default"
    Sheets("MortgageCashFlows").Cells(2, 8).Value = "Loss"
    Sheets("MortgageCashFlows").Cells(2, 9).Value = "Month CPR"
    Sheets("MortgageCashFlows").Cells(2, 10).Value = "Month CDR"
    Sheets("MortgageCashFlows").Cells(2, 11).Value = "Net_Rate Rate"
    Sheets("MortgageCashFlows").Cells(2, 12).Value = "Total Cash"

    For i = 1 To min(m.balloon_term_months, m.maturity_months)
        Sheets("MortgageCashFlows").Cells(i + 2, 1).Value = i
        Sheets("MortgageCashFlows").Cells(i + 2, 2).Value = balance
        Sheets("MortgageCashFlows").Cells(i + 2, 3).Value = m.cashflows(i).schedPay
        Sheets("MortgageCashFlows").Cells(i + 2, 4).Value = m.cashflows(i).interest
        Sheets("MortgageCashFlows").Cells(i + 2, 5).Value = m.cashflows(i).schedPrin
        Sheets("MortgageCashFlows").Cells(i + 2, 6).Value = m.cashflows(i).prepayPrin
        Sheets("MortgageCashFlows").Cells(i + 2, 7).Value = m.cashflows(i).default
        Sheets("MortgageCashFlows").Cells(i + 2, 8).Value = m.cashflows(i).loss
        If (balance - m.cashflows(i).schedPrin - m.cashflows(i).prepayPrin) > 0 Then
            Sheets("MortgageCashFlows").Cells(i + 2, 9).Value = 1 - Application.WorksheetFunction.Power(1 - m.cashflows(i).prepayPrin / (balance - m.cashflows(i).schedPrin - m.cashflows(i).default), 12)
            Sheets("MortgageCashFlows").Cells(i + 2, 10).Value = 1 - Application.WorksheetFunction.Power(1 - m.cashflows(i).default / (balance - m.cashflows(i).schedPrin), 12)
        End If
        Sheets("MortgageCashFlows").Cells(i + 2, 11).Value = m.cashflows(i).interest * 12 / balance
        Sheets("MortgageCashFlows").Cells(i + 2, 12).Value = m.cashflows(i).interest + m.cashflows(i).schedPrin + m.cashflows(i).prepayPrin + m.cashflows(i).default - m.cashflows(i).loss
        balance = balance - m.cashflows(i).default - m.cashflows(i).prepayPrin - m.cashflows(i).schedPrin
    Next i
    
End Sub

Sub printMortgageResults(ByRef m As mortgage)

    Sheets("MortgageCalc").Range("i6").Value = m.price - m.accrued
    Sheets("MortgageCalc").Range("i7").Value = m.accrued
    Sheets("MortgageCalc").Range("i8").Value = m.price
    Sheets("MortgageCalc").Range("i10").Value = m.BEY
    Sheets("MortgageCalc").Range("i11").Value = mey(m.BEY)
    Sheets("MortgageCalc").Range("i15").Value = m.modDur
    Sheets("MortgageCalc").Range("i17").Value = m.WAL

End Sub
