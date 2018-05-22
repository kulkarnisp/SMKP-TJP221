Attribute VB_Name = "TJPzfunc"
Option Explicit
'Option Base 1
Const Pi As Double = 3.14159265358979
Function InPolygon(XX, YY, Xa, Ya) As Boolean
InPolygon = False

Dim m, x1, x2, y1, y2, Xtest As Double
Dim i, j, ix, iy As Integer
Dim ybound As Boolean

ix = UBound(XX, 1)
iy = UBound(YY, 1)
If Not (ix = iy) Then: GoTo ErrHdl

For i = 0 To ix
j = i + 1
If i = ix Then: j = 0

x1 = XX(i): x2 = XX(j)
y1 = YY(i): y2 = YY(j)
ybound = False
    Select Case Ya
    Case y2 To y1
    ybound = True
    Case y1 To y2
    ybound = True
    End Select

If ybound Then
m = (x2 - x1) / (y2 - y1)
Xtest = (Ya - y1) * m + x1
If Xtest < Xa Then: InPolygon = Not (InPolygon)
End If
Next i

Exit Function

ErrHdl:
MsgBox "Error in JordonCurveTheorem"
End

End Function

Function CSplineA(Xa As Variant, Ya As Variant, Xint As Variant, Optional Out As Long = 1, Optional EType As Long = 1, _
                  Optional End1 As Double = 0, Optional End2 As Double = 0, Optional TransposeH As Boolean = True)


' ETypes: 1 = Specified 2nd derivative, 2 = Specified slope
' End1 and End2 = specified curvature or slope

Dim i As Long, n As Long, Yint As Variant, RevX As Boolean

Dim y2() As Double, XiTemp(1 To 1, 1 To 1) As Double
Dim a() As Double, b() As Double, c() As Double, r() As Double

'
    Xa = GetArray(Xa)
    Ya = GetArray(Ya)

    n = UBound(Xa)
    RevX = CheckAscX(Xa, Ya, n)

    ReDim L(1 To n - 1)
    ReDim m(1 To n - 1)

    For i = 1 To n - 1
        L(i) = Xa(i + 1, 1) - Xa(i, 1)
        m(i) = (Ya(i + 1, 1) - Ya(i, 1)) / L(i)
    Next i


    ReDim a(1 To n, 1 To 1)
    ReDim b(1 To n, 1 To 1)
    ReDim c(1 To n, 1 To 1)
    ReDim r(1 To n, 1 To 1)

    ' Form arrays a, b, c, and r for Tridiag
    Select Case EType
    Case 1

        r(1, 1) = End1
        r(n, 1) = End2

        a(1, 1) = 0
        b(1, 1) = 1
        c(1, 1) = 0

        a(n, 1) = 0
        b(n, 1) = 1
        c(n, 1) = 0


    Case 2
        r(1, 1) = 6 * (m(1) - End1)
        r(n, 1) = 6 * (End2 - m(n - 1))
        b(1, 1) = 2 * L(1)
        c(1, 1) = L(1)
        a(1, 1) = 0

        c(n, 1) = 0
        a(n, 1) = L(n - 1)
        b(n, 1) = 2 * L(n - 1)

    End Select

    For i = 2 To n - 1
        a(i, 1) = L(i - 1)
        b(i, 1) = 2 * (L(i - 1) + L(i))
        c(i, 1) = L(i)
        r(i, 1) = 6 * (m(i) - m(i - 1))
    Next i

    y2 = TriSolve(a, b, c, r)

    For i = 1 To n - 1
        y2(i, 2) = Ya(i, 1)
        y2(i, 3) = m(i) - (L(i) / 6) * (2 * y2(i, 1) + y2(i + 1, 1))
        y2(i, 4) = y2(i, 1) / 2
        y2(i, 5) = (y2(i + 1, 1) - y2(i, 1)) / (6 * L(i))
    Next i

    If Out = 2 Then
        CSplineA = y2

    ElseIf Out = 1 Then

        Yint = Splineint(Xa, Ya, y2, Xint, TransposeH)
        If TransposeH = True Then Yint = WorksheetFunction.Transpose(Yint)
        CSplineA = Yint
    Else: CSplineA = "Invalid ""Out"" value"
    End If
End Function

Function CheckAscX(Xa As Variant, Ya As Variant, n As Long) As Boolean
Dim Temp As Variant, i As Long
' If last x < first x, reverse Xa and Ya and return True, else return False
    If Xa(n, 1) < Xa(1, 1) Then
        Temp = Xa
        For i = 1 To n
            Xa(i, 1) = Temp(n - i + 1, 1)
        Next i
        Temp = Ya
        For i = 1 To n
            Ya(i, 1) = Temp(n - i + 1, 1)
        Next i
        CheckAscX = True
        Exit Function
    End If
    CheckAscX = False
End Function

Function Splineint(Xa As Variant, Ya As Variant, y2a As Variant, Xint As Variant, Optional TransposeH As Boolean) As Variant

Dim i As Long, j As Long, n As Long, nInt As Long, XintTemp(1 To 1, 1 To 1) As Double
Dim x As Double, xd1 As Double, xd2 As Double, b As Double, a As Double, y() As Double
Dim x_1 As Double, x_2 As Double, xi As Double
Dim L As Double, m As Double, c As Double, d As Double
Dim Chord As Double, ArcAng As Double, AvR As Double

    'If TypeName(Xint) = "Range" Then Xint = Xint.Value2
    Xint = GetArray(Xint, TransposeH)

    n = UBound(Xa)
    If IsArray(Xint) Then
        nInt = UBound(Xint)
    Else
        XintTemp(1, 1) = Xint
        Xint = XintTemp
        nInt = 1
    End If

    ReDim y(1 To nInt, 1 To 6)


    j = 1
    For i = 1 To nInt
        ' Find segment
        x = Xint(i, 1)

        If x <= Xa(1, 1) Then
            j = 1
        Else
            xd1 = -1
            xd2 = -1
            If x > Xa(j, 1) Then
                Do While ((xd1 < 0 And xd2 < 0) Or (xd1 > 0 And xd2 > 0)) And (j < n)
                    x_1 = Xa(j, 1)
                    x_2 = Xa(j + 1, 1)
                    xd1 = x_1 - x
                    xd2 = x_2 - x
                    j = j + 1
                Loop
                j = j - 1
            Else
                Do While ((xd1 < 0 And xd2 < 0) Or (xd1 > 0 And xd2 > 0)) And (j > 0)
                    x_1 = Xa(j - 1, 1)
                    x_2 = Xa(j, 1)
                    xd1 = x_1 - x
                    xd2 = x_2 - x
                    j = j - 1
                Loop
            End If
        End If

        xi = x - Xa(j, 1)
        L = Xa(j + 1, 1) - Xa(j, 1)
        m = (Ya(j + 1, 1) - Ya(j, 1)) / L
        a = y2a(j, 2)
        b = y2a(j, 3)
        c = y2a(j, 4)
        d = y2a(j, 5)

        y(i, 1) = a + b * xi + c * xi ^ 2 + d * xi ^ 3
        y(i, 2) = b + 2 * c * xi + 3 * d * xi ^ 2
        y(i, 3) = 2 * c + 6 * d * xi

        If Abs(y(i, 3)) > 0.0000000001 Then
            y(i, 4) = (1 + y(i, 2) ^ 2) ^ (3 / 2) / y(i, 3)
        End If
        If i > 1 Then
            Chord = ((Xint(i, 1) - Xint(i - 1, 1)) ^ 2 + (y(i, 1) - y(i - 1, 1)) ^ 2) ^ 0.5
            ArcAng = Atn(y(i, 2)) - Atn(y(i - 1, 2))
            If ArcAng <> 0 Then AvR = Chord / 2 / Sin(ArcAng / 2)
            y(i, 5) = ArcAng * AvR
            y(i, 6) = Chord
        Else
        End If

    Next i
    Splineint = y
End Function


Function TriSolve(a As Variant, b As Variant, c As Variant, r As Variant) As Variant
Dim U() As Double, n As Long

Dim j As Long, bet As Double, gam() As Double
    If TypeName(a) = "Range" Then a = a.Value2
    If TypeName(b) = "Range" Then b = b.Value2
    If TypeName(c) = "Range" Then c = c.Value2
    If TypeName(r) = "Range" Then r = r.Value2

    n = UBound(a)
    ReDim U(1 To n, 1 To 5)
    ReDim gam(1 To n)

    If (b(1, 1) <> 0#) Then

        bet = b(1, 1)
        U(1, 1) = r(1, 1) / bet
    End If
    For j = 2 To n
        gam(j) = c(j - 1, 1) / bet
        bet = b(j, 1) - a(j, 1) * gam(j)
        If (bet <> 0#) Then
            U(j, 1) = (r(j, 1) - a(j, 1) * U(j - 1, 1)) / bet
        End If
    Next j
    For j = (n - 1) To 1 Step -1
        U(j, 1) = U(j, 1) - gam(j + 1) * U(j + 1, 1)

    Next j
    TriSolve = U
End Function


Public Function GetArray(Arrayname As Variant, Optional TransposeH As Boolean = True) As Variant
Dim TempA() As Variant, LBound1 As Long
Dim UBound1 As Long, UBound2 As Long
Dim i As Long, j As Long

    
        If TypeName(Arrayname) = "Range" Then
            If Arrayname.Rows.count < Arrayname.Columns.count Then
                GetArray = WorksheetFunction.Transpose(Arrayname)
            Else
                TransposeH = False
                GetArray = Arrayname.Value2
            End If
            Exit Function
        End If
 


    ' If Arrayname is not an array, convert it into one.
    ' IsArray is true for multi-cell ranges, but not for a single cell range
    If Not IsArray(Arrayname) Then
        Arrayname = Array(Arrayname)
    End If    ' If Arrayname is a range, convert it into an array containing the range cell values
    If TypeName(Arrayname) = "Range" Then
        GetArray = Arrayname.Value2
        'Otherwise simply allocate the array to GetArray
    Else
        GetArray = Arrayname
    End If

    ' Check for a 1D array, or a base 0 array
    On Error Resume Next
    UBound2 = UBound(GetArray, 2)
    ' Convert to base 1
    If UBound2 = 0 Then
        LBound1 = LBound(GetArray)
        UBound1 = UBound(GetArray)
        ReDim TempA(1 To 1, 1 To UBound1 - LBound1 + 1)
        j = 1
        For i = LBound1 To UBound1
            TempA(1, j) = Arrayname(i)
            j = j + 1
        Next i

        GetArray = TempA
    End If


End Function
Function CSArea(Xa As Variant, Ya As Variant, Optional XLimits As Variant, Optional EType As Long = 1, Optional End1 As Double = 0, Optional End2 As Double = 0) As Variant


' ETypes: 1 = Specified 2nd derivative, 2 = Specified slope
' End1 and End2 = specified curvature or slope

Dim i As Long, j As Long, k As Long, n As Long, SCoeffA As Variant, XLimTemp(1 To 2, 1 To 1)

Dim x_1 As Double, x_2 As Double, x As Double, xd1 As Double, xd2 As Double, L As Double, Area As Double

    '    If TypeName(Xa) = "Range" Then Xa = Xa.Value2
    '    If TypeName(Ya) = "Range" Then Ya = Ya.Value2
    Xa = GetArray(Xa)
    Ya = GetArray(Ya)

    n = UBound(Xa)

    If IsMissing(XLimits) = True Then
        ReDim XLimits(1 To 2, 1 To 1)
        XLimits(1, 1) = Xa(1, 1)
        XLimits(2, 1) = Xa(n, 1)
    End If


    If TypeName(XLimits) = "Range" Then XLimits = XLimits.Value2
    If IsArray(XLimits) = False Then
        CSArea = "Xlimits must be a range with at least two rows or columns"
        Exit Function
    End If

    If UBound(XLimits) = 1 Then
        If UBound(XLimits, 2) < 2 Then
            CSArea = "Xlimits must be a range with at least two rows or columns"
            Exit Function
        Else
            XLimTemp(1, 1) = XLimits(1, 1)
            XLimTemp(2, 1) = XLimits(1, 2)
            ReDim XLimits(1 To 2, 1 To 1)
            XLimits(1, 1) = XLimTemp(1, 1)
            XLimits(2, 1) = XLimTemp(2, 1)
        End If
    End If

    SCoeffA = CSplineA(Xa, Ya, XLimits, 2, EType, End1, End2)



    ' Find segments containing limits of integration
    x = XLimits(1, 1)

    xd1 = -1
    xd2 = -1
    j = 1
    If x > Xa(1, 1) Then
        Do While ((xd1 < 0 And xd2 < 0) Or (xd1 > 0 And xd2 > 0)) And (j < n)
            x_1 = Xa(j, 1)
            x_2 = Xa(j + 1, 1)
            xd1 = x_1 - x
            xd2 = x_2 - x
            j = j + 1
        Loop
        j = j - 1
    Else
        j = 1
    End If

    x = XLimits(2, 1)

    xd1 = -1
    xd2 = -1
    k = 1
    If x > Xa(1, 1) Then
        Do While ((xd1 < 0 And xd2 < 0) Or (xd1 > 0 And xd2 > 0)) And (k < n)
            x_1 = Xa(k, 1)
            x_2 = Xa(k + 1, 1)
            xd1 = x_1 - x
            xd2 = x_2 - x
            k = k + 1
        Loop
        k = k - 1
    Else
        k = 1
    End If


    Area = 0

    For i = j To k

        If i = j Then
            x_1 = XLimits(1, 1)
        Else
            x_1 = Xa(i, 1)
        End If

        If i = k Then
            x_2 = XLimits(2, 1)
        Else
            x_2 = Xa(i + 1, 1)
        End If

        L = x_2 - Xa(i, 1)

        Area = Area + L * SCoeffA(i, 2) + L ^ 2 / 2 * SCoeffA(i, 3) + L ^ 3 / 3 * SCoeffA(i, 4) + L ^ 4 / 4 * SCoeffA(i, 5)
        If x_1 <> Xa(i, 1) Then
            L = x_1 - Xa(i, 1)
            Area = Area - (L * SCoeffA(i, 2) + L ^ 2 / 2 * SCoeffA(i, 3) + L ^ 3 / 3 * SCoeffA(i, 4) + L ^ 4 / 4 * SCoeffA(i, 5))

        End If
    Next i
    CSArea = Area

End Function



Function ChartSplineA(PositionA As Variant, Optional Series, Optional ChartObj)


'Returns x,y values at a given position along a smoothed chart line
' From http://groups.google.com/group/microsoft.public.excel.charting/browse_thread/thread/2406846f5b6c9d29/09417169ec10d29b

' Converted to array function Doug Jenkins 13 May 2010

Dim Chrt As Chart, ChrtS As Series, a As Variant, i As Long, j As Long, _
    S As Double, t As Double, L(0 To 1) As Double, P(0 To 1, 0 To 3) As Double, _
    d(0 To 1, 0 To 2) As Double, U(0 To 2) As Double, QA() As Double, z As Double, n As Long
Dim NumPoints As Long, Position As Double


    Application.Volatile

    '    If TypeName(PositionA) = "Range" Then PositionA = PositionA.Value2
    PositionA = GetArray(PositionA)
    NumPoints = UBound(PositionA)

    Set Chrt = Application.Caller.Worksheet _
               .ChartObjects(IIf(IsMissing(ChartObj), 1, ChartObj)).Chart
    Set ChrtS = Chrt.SeriesCollection(IIf(IsMissing(Series), 1, Series))


    L(0) = (Chrt.Axes(xlCategory).MaximumScale - _
            Chrt.Axes(xlCategory).MinimumScale) / Chrt.PlotArea.InsideWidth
    L(1) = (Chrt.Axes(xlValue).MaximumScale - _
            Chrt.Axes(xlValue).MinimumScale) / Chrt.PlotArea.InsideHeight


    a = Array(ChrtS.XValues, ChrtS.Values)
    n = UBound(a(1)) - 2
    ReDim QA(1 To NumPoints, 1 To 2)

    For j = 1 To NumPoints
        Position = PositionA(j, 1)
        S = Int(Position) + (Position = n + 1)
        t = Position - S


        For i = 0 To 1
            P(i, 1) = a(i + 1)(S + 1)
            P(i, 2) = a(i + 1)(S + 2)
            P(i, 0) = a(i + 1)(S - (S = 0)) - (S = 0) * (P(i, 1) - P(i, 2))
            P(i, 3) = a(i + 1)(S + 3 + (S = n)) + (S = n) * (P(i, 1) - P(i, 2))
            d(i, 0) = (P(i, 2) - P(i, 1)) / L(i)
            d(i, 1) = (P(i, 2) - P(i, 0)) / L(i) / 3
            d(i, 2) = (P(i, 3) - P(i, 1)) / L(i) / 3
        Next i


        For i = 0 To 2
            U(i) = d(0, i) ^ 2 + d(1, i) ^ 2
        Next i
        z = (U(0) / WorksheetFunction.Max(U)) ^ 0.5 / 2


        For i = 0 To 1
            QA(j, i + 1) = t ^ 2 * (3 - 2 * t) * P(i, 2) + _
                           (1 - t) ^ 2 * (1 + 2 * t) * P(i, 1) + _
                           z * t * (1 - t) * (t * (P(i, 1) - P(i, 3)) + _
                                              (1 - t) * (P(i, 2) - P(i, 0)))
        Next i
    Next j

    ChartSplineA = QA


End Function



Function SolveSplineA(Xa As Variant, Ya As Variant, Yint As Variant, Optional Xrange As Variant, Optional NumOut As Long = 0, _
                      Optional EType As Long = 1, Optional End1 As Double = 0, Optional End2 As Double = 0)


' ETypes: 1 = Specified 2nd derivative, 2 = Specified slope
' End1 and End2 = specified curvature or slope

Dim i As Long, j As Long, k As Long, m As Long, NumX As Long, NumY As Long, ResA() As Variant, CubRes As Variant, NumReal As Long

Dim y2() As Double
Dim a As Double, b As Double, c As Double, d As Double
Dim Startx As Long, Startx0 As Long

    y2 = CSplineA(Xa, Ya, Xa, 2, EType, End1, End2)

    If TypeName(Yint) = "Range" Then Yint = Yint.Value2
    If IsArray(Yint) = False Then
        Yint = Array(Yint, 0)
        Yint = WorksheetFunction.Transpose(Yint)
        NumY = 1
    Else
        If UBound(Yint, 2) > UBound(Yint, 1) Then Yint = WorksheetFunction.Transpose(Yint)
        NumY = UBound(Yint)
    End If
    If NumOut = 0 Then NumOut = NumY

    Startx0 = 1
    If IsMissing(Xrange) = True Then
        Xrange = Xa
        NumX = UBound(Xrange)
    Else
        If TypeName(Xrange) = "Range" Then Xrange = Xrange.Value2
        If UBound(Xrange, 2) > UBound(Xrange, 1) Then Xrange = WorksheetFunction.Transpose(Xrange)
        Do While Xa(Startx0, 1) < Xrange(1, 1)
            Startx0 = Startx0 + 1
        Loop
        NumX = Startx0 + 1
        Do While Xa(NumX, 1) < Xrange(2, 1)
            NumX = NumX + 1
        Loop
    End If

    ReDim ResA(1 To NumOut, 1 To 2)

    m = 1
    For i = 1 To NumY
        Startx = Startx0
        If i > 1 Then
            ' If Yint(i, 1) = Yint(i - 1, 1) Then Startx = j + 1
        End If
        For j = Startx To NumX - 1
            a = y2(j, 5)
            b = y2(j, 4)
            c = y2(j, 3)
            d = y2(j, 2) - Yint(i, 1)

            CubRes = cubic(a, b, c, d)
            NumReal = CubRes(4, 1)
            For k = 1 To NumReal
                If (CubRes(k, 1) - (Xa(j + 1, 1) - Xa(j, 1)) <= 0) And (CubRes(k, 1) >= 0) Then
                    ResA(m, 2) = Yint(i, 1)
                    ResA(m, 1) = CubRes(k, 1) + Xa(j, 1)
                    If m < NumOut Then m = m + 1 Else Exit For
                End If
            Next k
            If ResA(m, 1) <> Empty Then Exit For
        Next j


    Next i
    If m = 1 Then
        If i <= NumOut Then ResA(i, 1) = "No solution found within the specified x range"
    ElseIf ResA(m, 2) = Empty Then
        ResA(m, 1) = m - 1 & " solutions found within the specified x range"
    End If


    SolveSplineA = ResA

End Function

Function CHSplineA(Xa As Variant, Ya As Variant, Xint As Variant, Optional Mono As Boolean = False, _
                Optional Out As Double = 1, Optional TransposeH As Boolean = True)


' ETypes: 1 = Specified 2nd derivative, 2 = Specified slope
' End1 and End2 = specified curvature or slope

Dim i As Long, n As Long, nInt As Long, Yint As Variant, j As Long, t As Double
Dim L() As Double, S() As Double, m() As Double, h() As Double, Cubica() As Double
Dim Alpha As Double, Beta As Double, Tau As Double, RevX As Boolean

    '    If TypeName(Xa) = "Range" Then Xa = Xa.Value2
    '    If TypeName(Ya) = "Range" Then Ya = Ya.Value2
    '    If TypeName(Xint) = "Range" Then Xint = Xint.Value2

    Xa = GetArray(Xa)
    Ya = GetArray(Ya)
    Xint = GetArray(Xint, TransposeH)

    n = UBound(Xa)
    nInt = UBound(Xint)
    RevX = CheckAscX(Xa, Ya, n)

    ReDim L(1 To n - 1)
    ReDim S(1 To n)
    ReDim m(1 To n)
    ReDim h(1 To n)
    If Out = 1 Then
        ReDim Yint(1 To nInt, 1 To 1)
    Else
        ReDim Yint(1 To nInt, 1 To 2)
    End If
    ReDim Cubica(1 To n - 1, 1 To 4)

    i = 1
    L(i) = Xa(i + 1, 1) - Xa(i, 1)
    h(i) = Ya(i + 1, 1) - Ya(i, 1)
    S(i) = (h(i) / L(i))
    m(i) = S(i)
    For i = 2 To n - 1
        L(i) = Xa(i + 1, 1) - Xa(i, 1)
        h(i) = Ya(i + 1, 1) - Ya(i, 1)
        S(i) = (h(i) / L(i))
        m(i) = (S(i - 1) + S(i)) / 2
    Next i
    h(i) = h(i - 1)
    S(i) = S(i - 1)
    m(i) = S(i)

    If Mono = True Then
        For i = 1 To n - 1
            If m(i) >= 0 Then
                If h(i) = 0 Then
                    m(i) = 0
                    m(i + 1) = 0
                Else
                    Alpha = m(i) / h(i)
                    Beta = m(i + 1) / h(i)
                    If (Alpha = 0) Or (Beta = 0) Then
                        m(i) = 0
                        m(i + 1) = 0
                    ElseIf Alpha ^ 2 + Beta ^ 2 > 9 Then
                        Tau = 3 / (Alpha ^ 2 + Beta ^ 2) ^ 0.5
                        m(i) = Tau * Alpha * h(i)
                        m(i + 1) = Tau * Beta * h(i)
                    End If
                End If
            End If
        Next i
    End If

    For i = 1 To n - 1
        Cubica(i, 1) = 2 * (Ya(i, 1) - Ya(i + 1, 1)) + (m(i) + m(i + 1)) * L(i)
        Cubica(i, 2) = 3 * (Ya(i + 1, 1) - Ya(i, 1)) - (2 * m(i) + m(i + 1)) * L(i)
        Cubica(i, 3) = m(i) * L(i)
        Cubica(i, 4) = Ya(i, 1)
    Next i

    For i = 1 To nInt
        j = 0
        Do
            j = j + 1
        Loop While (Xint(i, 1) > Xa(j + 1, 1) And j < n - 1)
        t = (Xint(i, 1) - Xa(j, 1)) / L(j)
        If Out = 1 Then
            Yint(i, 1) = Cubica(j, 1) * t ^ 3 + Cubica(j, 2) * t ^ 2 + Cubica(j, 3) * t + Cubica(j, 4)
        Else
            Yint(i, 1) = (3 * Cubica(j, 1) * t ^ 2 + 2 * Cubica(j, 2) * t + Cubica(j, 3)) / L(j)
            Yint(i, 2) = (6 * Cubica(j, 1) * t + 2 * Cubica(j, 2)) / L(j) ^ 2
        End If
    Next i
    If TransposeH = True Then Yint = WorksheetFunction.Transpose(Yint)
    CHSplineA = Yint
End Function

Function CardSpline(Xa As Variant, Ya As Variant, Pint As Variant, Tens As Double, Optional PType As String = "L") As Variant
Dim S As Double, U As Double, UA(1 To 1, 1 To 4) As Double, CardA(1 To 4, 1 To 4) As Double, Card2A As Variant
Dim XA2() As Double, YA2() As Double, XYA As Variant, n As Long, nInt As Long, i As Long, j As Long, k As Long
Dim XA3(1 To 4, 1 To 1) As Double, YA3(1 To 4, 1 To 1) As Double, TransposeH As Boolean

    '    If TypeName(Xa) = "Range" Then Xa = Xa.Value2
    '    If TypeName(Ya) = "Range" Then Ya = Ya.Value2
    '    If TypeName(Pint) = "Range" Then Pint = Pint.Value2
    Xa = GetArray(Xa)
    Ya = GetArray(Ya)
    TransposeH = True
    Pint = GetArray(Pint, TransposeH)

    n = UBound(Xa)
    nInt = UBound(Pint)
    ReDim XA2(1 To 4, 1 To n)
    ReDim YA2(1 To 4, 1 To n)

    S = (1 - Tens) / 2

    CardA(1, 1) = (-S)
    CardA(1, 2) = (2 - S)
    CardA(1, 3) = (S - 2)
    CardA(1, 4) = (S)
    CardA(2, 1) = (2 * S)
    CardA(2, 2) = (S - 3)
    CardA(2, 3) = (3 - (2 * S))
    CardA(2, 4) = (-1 * S)
    CardA(3, 1) = (-1 * S)
    CardA(3, 2) = (0)
    CardA(3, 3) = (S)
    CardA(3, 4) = (0)
    CardA(4, 1) = (0)
    CardA(4, 2) = (1)
    CardA(4, 3) = (0)
    CardA(4, 4) = (0)


    For i = 2 To n - 2
        XA2(1, i) = Xa(i - 1, 1)
        XA2(2, i) = Xa(i, 1)
        XA2(3, i) = Xa(i + 1, 1)
        XA2(4, i) = Xa(i + 2, 1)

        YA2(1, i) = Ya(i - 1, 1)
        YA2(2, i) = Ya(i, 1)
        YA2(3, i) = Ya(i + 1, 1)
        YA2(4, i) = Ya(i + 2, 1)
    Next i

    PType = UCase(PType)

    If PType = "L" Then
        ReDim XYA(1 To nInt, 1 To 2)
        For j = 1 To nInt


            i = Pint(j, 1)
            If Abs(Pint(j, 1) - i) > 0.00000000000001 Then
                i = Int(Pint(j, 1))
                U = Pint(j, 1) - i
            Else
                If i > 1 Then
                    i = Pint(j, 1) - 1
                    U = 1
                Else
                    i = 1
                    U = 0
                End If
            End If
            i = i + 1
            UA(1, 1) = U ^ 3
            UA(1, 2) = U ^ 2
            UA(1, 3) = U
            UA(1, 4) = 1

            For k = 1 To 4
                XA3(k, 1) = XA2(k, i)
                YA3(k, 1) = YA2(k, i)
            Next k

            Card2A = MMult(UA, CardA)
            XYA(j, 1) = MMult(Card2A, XA3)
            XYA(j, 2) = MMult(Card2A, YA3)
        Next j

    ElseIf PType = "X" Or PType = "Y" Then
        XYA = SolveCardSpline(Pint, XA2, YA2, CardA, n, PType)
    Else
        XYA(1, 1) = "Invalid PType"
    End If
    If TransposeH = True Then XYA = WorksheetFunction.Transpose(XYA)
    CardSpline = XYA    'CardA
End Function


Function SolveCardSpline(Pint As Variant, XA2() As Double, YA2() As Double, CardA() As Double, n As Long, PType As String) As Variant
Dim UA(1 To 1, 1 To 4) As Double, i As Long, j As Long, k As Long, Card2A As Variant, nInt As Long
Dim XA3(1 To 4, 1 To 1) As Double, YA3(1 To 4, 1 To 1) As Double, ResA() As Double, U As Double
Dim a As Double, b As Double, c As Double, d As Double, CubRes As Variant, NumReal As Long, LastI As Long, LastP As Double

    ' Based on public domain Perl script at http://en.wikipedia.org/wiki/File:Cardinal_Spline_Example.png
    ' Adapted to VBA by Doug Jenkins 13 May 2010

    nInt = UBound(Pint)
    ReDim ResA(1 To nInt, 1 To 2)
    LastI = 2
    LastP = -0.00000000000001
    For j = 1 To nInt

        For i = LastI To n - 2

            If PType = "X" Then
                For k = 1 To 4
                    XA3(k, 1) = XA2(k, i)
                Next k
            ElseIf PType = "Y" Then
                For k = 1 To 4
                    XA3(k, 1) = YA2(k, i)
                Next k
            End If

            Card2A = MMult(CardA, XA3)
            a = Card2A(1, 1)
            b = Card2A(2, 1)
            c = Card2A(3, 1)
            d = Card2A(4, 1) - Pint(j, 1)

            CubRes = cubic(a, b, c, d)
            NumReal = CubRes(4, 1)

            For k = 1 To NumReal
                If (CubRes(k, 1) <= 1) And (CubRes(k, 1) >= LastP) Then
                    U = CubRes(k, 1)
                    ResA(j, 2) = i - 1 + U
                    LastI = i
                    LastP = U
                    Exit For
                End If
            Next k
            If ResA(j, 2) <> Empty Then Exit For
            LastP = -0.00000000000001
        Next i

        UA(1, 1) = U ^ 3
        UA(1, 2) = U ^ 2
        UA(1, 3) = U
        UA(1, 4) = 1

        If PType = "X" Then
            For k = 1 To 4
                YA3(k, 1) = YA2(k, i)
            Next k

            Card2A = MMult(UA, CardA)
            ResA(j, 1) = MMult(Card2A, YA3)

        ElseIf PType = "Y" Then
            For k = 1 To 4
                XA3(k, 1) = XA2(k, i)
            Next k

            Card2A = MMult(UA, CardA)
            ResA(j, 1) = MMult(Card2A, XA3)
        End If
    Next j
    SolveCardSpline = ResA
End Function


Function MMult(Mat1 As Variant, Mat2 As Variant) As Variant
Dim NumRows1 As Long, NumRows2 As Long, NumCols1 As Long, NumCols2 As Long
Dim i As Long, j As Long, k As Long, ResA() As Double


    NumRows1 = UBound(Mat1)
    NumRows2 = UBound(Mat2)
    NumCols1 = UBound(Mat1, 2)
    NumCols2 = UBound(Mat2, 2)

    ReDim ResA(1 To NumRows1, 1 To NumCols2)

    If NumCols1 <> NumRows2 Then
        MMult = "Invalid array sizes"
        Exit Function
    End If

    For i = 1 To NumRows1
        For j = 1 To NumCols2
            For k = 1 To NumRows2
                ResA(i, j) = ResA(i, j) + Mat1(i, k) * Mat2(k, j)
            Next k
        Next j
    Next i

    If NumRows1 + NumCols2 = 2 Then
        MMult = ResA(1, 1)
    Else
        MMult = ResA
    End If
End Function


Function cubic(z As Double, a As Double, b As Double, c As Double, Optional Out As Long) As Variant

' Copyright (C) 2006 Interactive Design Services

' This program is free software; you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation; either version 2 of the License, or (at
' your option) any later version.

' This program is distributed in the hope that it will be useful, but
' WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
' General Public License for more details.

' You should have received a copy of the GNU General Public License
' along with this program; if not, write to the Free Software
' Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA 02110-1301, USA.


' cubic - finds the real roots of z x^3 + a x^2 + b x + c = 0

Dim q As Double, r As Double, Q_ As Double, R_ As Double, Q3 As Double, R2 As Double, CR2 As Double, CQ3 As Double
Dim sqrtQ As Double, sqrtQ3 As Double, Theta As Double, Norm As Double
Dim sgnR As Long, A_ As Double, B_ As Double    ', Pi As Double
Dim Cubica(1 To 4, 1 To 1) As Variant, QuadA As Variant

    '   Pi = 4 * Atn(1)

    If Out < 0 Then
        cubic = CVErr(xlErrNA)
        Exit Function
    End If

    If z = 0 Then
        ' QuadA = Quadratic(a, b, c)
        Cubica(1, 1) = QuadA(1, 1)
        Cubica(2, 1) = QuadA(2, 1)
        Cubica(3, 1) = 0
        Cubica(4, 1) = QuadA(3, 1)


    Else

        If z <> 1 Then
            a = a / z
            b = b / z
            c = c / z
        End If

        q = (a * a - 3 * b)
        r = (2 * a * a * a - 9 * a * b + 27 * c)

        Q_ = q / 9
        R_ = r / 54

        Q3 = Q_ * Q_ * Q_
        R2 = R_ * R_

        CR2 = 729 * r * r
        CQ3 = 2916 * q * q * q

        If (R_ = 0 And Q_ = 0) Then
            Cubica(4, 1) = 3

            Cubica(1, 1) = -a / 3
            Cubica(2, 1) = -a / 3
            Cubica(3, 1) = -a / 3

            GoTo select_out

        ElseIf (CR2 = CQ3) Then


            Cubica(4, 1) = 3
            sqrtQ = (Q_) ^ 0.5

            If (r > 0) Then

                Cubica(1, 1) = -2 * sqrtQ - a / 3
                Cubica(2, 1) = sqrtQ - a / 3
                Cubica(3, 1) = sqrtQ - a / 3

            Else

                Cubica(1, 1) = -sqrtQ - a / 3
                Cubica(2, 1) = -sqrtQ - a / 3
                Cubica(3, 1) = 2 * sqrtQ - a / 3
                GoTo select_out

            End If


        ElseIf (CR2 < CQ3) Then


            Cubica(4, 1) = 3
            sqrtQ = (Q_) ^ 0.5
            sqrtQ3 = sqrtQ * sqrtQ * sqrtQ
            Theta = WorksheetFunction.Acos(R_ / sqrtQ3)
            Norm = -2 * sqrtQ
            Cubica(1, 1) = Norm * Cos(Theta / 3) - a / 3
            Cubica(2, 1) = Norm * Cos((Theta + 2# * Pi) / 3) - a / 3
            Cubica(3, 1) = Norm * Cos((Theta - 2# * Pi) / 3) - a / 3

            ' Sort CubicA into increasing order
            BubbleSort Cubica, 1, 3

            GoTo select_out

        Else
            Cubica(4, 1) = 1
            If R_ >= 0 Then sgnR = 1 Else sgnR = -1
            A_ = -sgnR * (Abs(R_) + (R2 - Q3) ^ 0.5) ^ (1 / 3)
            B_ = Q_ / A_
            Cubica(1, 1) = A_ + B_ - a / 3
            Cubica(2, 1) = 0
            Cubica(3, 1) = 0

        End If
    End If

select_out:

    If Out > 0 Then
        cubic = Cubica(Out, 1)
    Else
        cubic = Cubica
    End If

End Function

Function Quadratic(a As Double, b As Double, c As Double, Optional Out As Long) As Variant

' * Copyright (C) 2006 Interactive Design Services Pty Ltd

'  This program is free software; you can redistribute it and/or modify
'  it under the terms of the GNU General Public License as published by
'  the Free Software Foundation; either version 2 of the License, or (at
'  your option) any later version.

'  This program is distributed in the hope that it will be useful, but
'  WITHOUT ANY WARRANTY; without even the implied warranty of
'  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
'  General Public License for more details.

'  You should have received a copy of the GNU General Public License
'  along with this program; if not, write to the Free Software
'  Foundation, Inc., 675 Mass Ave, Cambridge, MA 02139, USA.


'quadratic - finds the roots of  a x^2 + b x + c = 0

Dim Disc As Double, SqrtDisc As Double, QuadA(1 To 3, 1 To 1) As Variant, QuadAI(1 To 3, 1 To 1) As Variant

    If Out < -3 Then
        Quadratic = CVErr(xlErrNA)
        Exit Function
    End If

    If a = 0 Then

        QuadA(1, 1) = -c / b
        QuadA(2, 1) = 0
        QuadA(3, 1) = 1

    Else

        Disc = b ^ 2 - 4 * a * c
        If Disc >= 0 Then SqrtDisc = Disc ^ 0.5 Else SqrtDisc = (-Disc) ^ 0.5

        If Disc >= 0 Then
            QuadA(1, 1) = (-b - SqrtDisc) / (2 * a)
            QuadA(2, 1) = (-b + SqrtDisc) / (2 * a)
            QuadA(3, 1) = 2
        Else
            QuadA(1, 1) = (-b) / (2 * a)
            QuadA(2, 1) = (-b) / (2 * a)
            QuadA(3, 1) = 0
            QuadAI(1, 1) = (-SqrtDisc) / (2 * a)
            QuadAI(2, 1) = (SqrtDisc) / (2 * a)
            QuadAI(3, 1) = 2
        End If
    End If

    If Out < -1 Then
        Quadratic = QuadAI(-Out, 1)
    ElseIf Out = -1 Then
        Quadratic = QuadAI
    ElseIf Out = 0 Then
        Quadratic = QuadA
    Else
        Quadratic = QuadA(Out, 1)
    End If

End Function



Sub BubbleSort(ToSort As Variant, LB As Long, UB As Long, Optional SortAscending As Boolean = True)

' Chris Rae's VBA Code Archive - http://chrisrae.com/vba
' By Chris Rae, 19/5/99.

' Amended for Quartic, Doug Jenkins 27/11/2006

Dim AnyChanges As Boolean
Dim BubbleSort As Long
Dim SwapFH As Variant
    Do
        AnyChanges = False
        For BubbleSort = LB To UB - 1
            If (ToSort(BubbleSort, 1) > ToSort(BubbleSort + 1, 1) And SortAscending) _
               Or (ToSort(BubbleSort, 1) < ToSort(BubbleSort + 1, 1) And Not SortAscending) Then
                ' These two need to be swapped
                SwapFH = ToSort(BubbleSort, 1)
                ToSort(BubbleSort, 1) = ToSort(BubbleSort + 1, 1)
                ToSort(BubbleSort + 1, 1) = SwapFH
                AnyChanges = True
            End If
        Next BubbleSort
    Loop Until Not AnyChanges
End Sub

Function InterpA(TableRange As Variant, RowValA As Variant, ColOffset As Long) As Variant
Dim NoRows As Long, i As Long, NoInts As Long, j As Long, RowVal As Double
Dim ROffset As Long
Dim XN As Double, XP As Double, YN As Double, YP As Double
Dim inta() As Double

    '    If TypeName(TableRange) = "Range" Then TableRange = TableRange.Value2
    '    If TypeName(RowValA) = "Range" Then RowValA = RowValA.Value2
    TableRange = GetArray(TableRange)
    RowValA = GetArray(RowValA)
    ' Find table size
    NoRows = UBound(TableRange)
    NoInts = UBound(RowValA)
    ReDim inta(1 To NoInts, 1 To 1)
    XP = TableRange(2, 1)
    XN = TableRange(3, 1)
    For j = 1 To NoInts
        RowVal = RowValA(j, 1)
        If XN > XP Then
            ' Find row offset
            For i = 2 To NoRows
                If RowVal < TableRange(i, 1) Then
                    ROffset = i
                    Exit For
                End If
            Next i
        ElseIf XN < XP Then
            For i = 2 To NoRows
                If RowVal > TableRange(i, 1) Then
                    ROffset = i
                    Exit For
                End If
            Next i
        End If

        XN = TableRange(ROffset, 1)
        XP = TableRange(ROffset - 1, 1)

        YN = TableRange(ROffset, ColOffset + 1)

        YP = TableRange(ROffset - 1, ColOffset + 1)


        inta(j, 1) = YP + ((RowVal) - (XP)) / ((XN) - (XP)) * (YN - YP)
    Next j
    InterpA = inta
End Function





