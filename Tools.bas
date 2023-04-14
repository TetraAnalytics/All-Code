Attribute VB_Name = "Tools"
Function modulus(ByVal n, ByVal d) As Double
    modulus = n - d * Int(n / d)
End Function

Function min(a, b)
    If a < b Then
        min = a
        Else: min = b
    End If
End Function

Function max(x, Y)
    If x > Y Then
        max = x
        Else: max = Y
    End If
End Function

Function NumberofMonths(ByVal d1 As Date, ByVal d2 As Date)
    NumberofMonths = (Year(d1) - Year(d2)) * 12 + month(d1) - month(d2)
End Function

'Cubic Spline
Option Base 1

Function cubic_spline(input_column As Range, output_column As Range, x As Double)

Dim input_count As Integer
Dim output_count As Integer

input_count = input_column.Rows.Count
output_count = output_column.Rows.Count

' Next check to be sure that "input" # points = "output" # points
If input_count <> output_count Then
    cubic_spline = "The number of indices and the number of output_columns don't match"
    GoTo out
End If
 
ReDim xin(input_count) As Single
ReDim yin(input_count) As Single

Dim c As Integer

For c = 1 To input_count
xin(c) = input_column(c)
yin(c) = output_column(c)
Next c

'''''''''''''''''''''''''''''''''''''''
' values are populated
'''''''''''''''''''''''''''''''''''''''
Dim n As Integer 'n=input_count
Dim i, k As Integer 'these are loop counting integers
Dim p, qn, sig, un As Single
ReDim u(input_count - 1) As Single
ReDim yt(input_count) As Single 'these are the 2nd deriv values

n = input_count
yt(1) = 0
u(1) = 0

For i = 2 To n - 1
    sig = (xin(i) - xin(i - 1)) / (xin(i + 1) - xin(i - 1))
    p = sig * yt(i - 1) + 2
    yt(i) = (sig - 1) / p
    u(i) = (yin(i + 1) - yin(i)) / (xin(i + 1) - xin(i)) - (yin(i) - yin(i - 1)) / (xin(i) - xin(i - 1))
    u(i) = (6 * u(i) / (xin(i + 1) - xin(i - 1)) - sig * u(i - 1)) / p
    
    Next i
    
qn = 0
un = 0

yt(n) = (un - qn * u(n - 1)) / (qn * yt(n - 1) + 1)

For k = n - 1 To 1 Step -1
    yt(k) = yt(k) * yt(k + 1) + u(k)
Next k


''''''''''''''''''''
'now eval spline at one point
'''''''''''''''''''''
Dim klo, khi As Integer
Dim h, b, a As Single

' first find correct interval
klo = 1
khi = n
Do
k = khi - klo
If xin(k) > x Then
khi = k
Else
klo = k
End If
k = khi - klo
Loop While k > 1
h = xin(khi) - xin(klo)
a = (xin(khi) - x) / h
b = (x - xin(klo)) / h
Y = a * yin(klo) + b * yin(khi) + ((a ^ 3 - a) * yt(klo) + (b ^ 3 - b) * yt(khi)) * (h ^ 2) / 6


cubic_spline = Y

out:
End Function

Public Function linearInterp(xArr As Variant, yArr As Variant, ByVal x As Double) As Double

    If ((x < xArr(LBound(xArr))) Or (x > xArr(UBound(xArr)))) Then
        MsgBox x & " " & xArr(LBound(xArr)) & " " & xArr(UBound(xArr))
        MsgBox ("linear Interpolation: x is out of bound.  Lower bound= " & xArr(LBound(xArr)) & " and Upperbound= " & xArr(UBound(xArr)))
        MsgBox ("X = " & x)
        Exit Function
    End If
    
    If xArr(LBound(xArr)) = x Then
      linearInterp = yArr(LBound(yArr))
      Exit Function
    End If
    
    Dim i As Single
    
    For i = LBound(xArr) To UBound(xArr)
      If xArr(i) >= x Then
        linearInterp = yArr(i - 1) + (x - xArr(i - 1)) / (xArr(i) - xArr(i - 1)) * (yArr(i) - yArr(i - 1))
        Exit Function
      End If
    Next i

End Function

