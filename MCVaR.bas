Option Base 1
Option Explicit

'====================================================================================================
Function decom(matrix As Range)
'====================================================================================================

    Dim n As Integer, m As Integer
    Dim i As Integer, j As Integer, k As Integer

    'set the columns and rows
    Dim A()
    A = matrix
    n = matrix.Rows.Count
    m = matrix.Columns.Count
    'start calculation if colomns = rows
    If m <> n Then
        Exit Function
    Else
    End If
    
    'cholesky decomposition C.C'=A
    Dim S As Double
    ReDim C(1 To n, 1 To n)
    For j = 1 To n
        S = 0
        For k = 1 To j - 1
            S = S + C(j, k) ^ 2
        Next k
        C(j, j) = A(j, j) - S
        C(j, j) = Sqr(C(j, j))

        For i = j + 1 To n
            S = 0
            For k = 1 To j - 1
                S = S + C(i, k) * C(j, k)
            Next k
            C(i, j) = (A(i, j) - S) / C(j, j)
        Next i
    Next j
    'output the result
    decom = C

End Function

'====================================================================================================
 Function multinorm(A) As Variant
'====================================================================================================

    Dim n As Integer, m As Integer
    Dim i As Integer, j As Integer, k As Integer
    n = A.Rows.Count
    m = A.Columns.Count

    'calculate vector of correlated normal variables, X = A.Z
    ReDim Z(n) As Variant      'independent random normal variable
    ReDim x(n) As Variant
    For i = 1 To m
        x(i) = 0
        For j = 1 To n
            Z(j) = -WorksheetFunction.NormSInv(Rnd)
            x(j) = x(i) + A(i, j) * Z(j)        'generate correlated variable
        Next j
    Next i

    multinorm = x

End Function

'====================================================================================================
Function MC_VaR(price As Range, vol As Range, shares As Range, A As Range, dt As Double, nscen As Long, alpha As Double) As Long
'====================================================================================================
    Dim m As Long
    Dim i As Long, j As Long, k As Long
    m = A.Columns.Count

    'import mean and std, and shares as array
    ReDim mu(m) As Variant
    ReDim sigma(m) As Variant
    ReDim quan(m) As Variant
    For i = 1 To m
        mu(i) = price(i)
        sigma(i) = vol(i) * Sqr(dt)     'convert annual vol to daily std
        quan(i) = shares(i)
    Next i

    'generate multivariate normal variables of prices
    Dim dS() As Variant
    ReDim dS(1 To nscen, 1 To m)
    Dim x() As Variant       'correlated standard normal vector
    ReDim dw(nscen)      'total wealth in portfolio
    For i = 1 To nscen
        x = multinorm(A)
        For j = 1 To m
        dw(i) = 0
            dS(i, j) = -mu(j) + sigma(j) * x(j)
            dw(i) = dw(i) + dS(i, j) * quan(j)
        Next j
    Next i
    
    quicksort dw, 1, nscen
    
    'quantile for VaR at significance level 1%
    Dim quantile As Double
    quantile = nscen * alpha

   MC_VaR = dw(quantile)
End Function

'====================================================================================================
Function An_VaR(price As Range, vol As Range, shares As Range, dt As Double, matrix As Range, Z As Double) As Double
'====================================================================================================
    Dim tmpvar As Double
    tmpvar = 0
    Dim i As Integer, j As Integer, n As Integer
    n = matrix.Columns.Count
    
    ReDim sigma(n) As Variant
    For i = 1 To n
        sigma(i) = vol(i) * Sqr(dt)     'convert annual vol to daily std
    Next i
    
    For i = 1 To n
        For j = 1 To n
            tmpvar = tmpvar + price(i) * shares(i) * price(j) * shares(j) * matrix(i, j) * sigma(i) * sigma(j)
        Next j
    Next i
    An_VaR = Z * Sqr(tmpvar)

End Function


Public Sub quicksort(varry As Variant, inlow As Long, inhi As Long)
Dim pivot As Variant, tmpswap As Variant
Dim tmplow As Long, tmphi As Long

tmplow = inlow
tmphi = inhi
pivot = varry((inlow + inhi) \ 2)

While (tmplow <= tmphi)
    While (varry(tmplow) < pivot And tmplow < inhi)
        tmplow = tmplow + 1
    Wend
    While (pivot < varry(tmphi) And tmphi > inlow)
        tmphi = tmphi - 1
    Wend
    
    If (tmplow <= tmphi) Then
        tmpswap = varry(tmplow)
        varry(tmplow) = varry(tmphi)
        varry(tmphi) = tmpswap
        tmplow = tmplow + 1
        tmphi = tmphi - 1
    End If
Wend

    If (inlow < tmphi) Then quicksort varry, inlow, tmphi
    If (tmplow < inhi) Then quicksort varry, tmplow, inhi
    
End Sub
