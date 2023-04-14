Attribute VB_Name = "BondDayCounts"
Function yearFrak(ByVal dt1 As Date, ByVal dt2 As Date, ByVal dt3 As Date, ByVal cnt As String, ByVal freq As Double) As Double

    Select Case cnt
        Case "30/360"
            yearFrak = DC30360(dt1, dt2, freq)
         Case "ACT/ACT"
            yearFrak = ACTACT(dt1, dt2, dt3, freq)
        Case "ACT/360"
            yearFrak = ACT360(dt1, dt2, freq)
        Case "ACT/365"
            yearFrak = ACT365(dt1, dt2, freq)
        Case Else
            MsgBox "Not correct day count input"
    End Select

End Function

Function DC30360(ByVal dt2 As Date, ByVal dt1 As Date, ByVal freq As Double) As Double
    Dim d1, d2, m1, m2, Y1, Y2 As Integer
    'for fixed leg of swap
    d1 = Day(dt1)
    d2 = Day(dt2)
    m1 = month(dt1)
    m2 = month(dt2)
    Y1 = Year(dt1)
    Y2 = Year(dt2)
    
    If d1 = 31 Then d1 = 30
    If d2 = 31 And d1 >= 30 Then d2 = 30
    If m1 = 2 And modulus(Year(dt1), 4) = 0 And d1 = 29 Then d1 = 30  ' test for leap year with modulus
    If m1 = 2 And modulus(Year(dt1), 4) <> 0 And d1 = 28 Then d1 = 30 ' not a leap year
        
    DC30360 = -(((360 * (Y2 - Y1)) + (30 * (m2 - m1)) + (d2 - d1)) / 360) * freq

End Function

Function ACT360(ByVal dt1 As Date, ByVal dt2 As Date, ByVal freq As Double) As Double
    ACT360 = -(dt1 - dt2) / 360 * freq * 360 / 366
End Function

Function ACT365(ByVal dt1 As Date, ByVal dt2 As Date, ByVal freq As Double) As Double
    ACT365 = -(dt1 - dt2) / 366 * freq
End Function

Function ACTACT(ByVal valueDate As Date, ByVal startDate As Date, ByVal nxtDate As Date, ByVal freq As Integer) As Double
'ISMA is for Treasuries
   nxtDate = DateAdd("m", 12 / freq, startDate) 'fix the 3.  was 6 for semi
   
   If (nxtDate - startDate) = 0 Then
     ACTACT = 0
     Else: ACTACT = (nxtDate - valueDate) / ((nxtDate - startDate))
   End If

End Function
