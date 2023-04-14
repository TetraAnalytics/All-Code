Attribute VB_Name = "calcMortgage"
Sub CalcMortgage()
    Dim m As mortgage
    
    Call loadMortgage(m)
    Call calcSingleMtgeCash(m)
    m.accrued = m.Net_Rate / 12 * (m.SDday - 1) / 30 * 100
    m.price = PVMtgCashFlows(m, m.BEY) / m.balance * 100
    m.BEY = CalcMtgYield(m)
    Call mtgeGreeks(m)
    Call printMortgageResults(m)
    Call printMortgageCashFlows(m)
End Sub
