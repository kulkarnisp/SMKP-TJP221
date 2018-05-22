Attribute VB_Name = "TJPamod"
    Dim Nd%, Nr%
    Public Qr#, Hr#
    Dim dmn#, dmx#
    
    Public solver_options
    
    Public Mdat(), Fdat() As Double
    Dim Edat(), nEdat() As Double
    Dim Vdat(), Wdat() As Double
    Dim Bdat(), Ddat() As Double
    
    Public cutr, modr As Double
    Dim ViscReq As Boolean
    Dim Cdat(), Ndat() As Double
    Dim Rdat() As Double
    
    Public Countr%
    Public Sdat() As String
    Const g As Double = 9.8065
    Const Pi As Double = 3.14159265358979
    

Sub VPF_Search()
'    On Error GoTo errhdl
    Starttime = Timer
    g = 40
    Dim j%
    Countr = 0
    Dim Vmodel As String
    Sheets("Input").Range("A14:Q80").ClearContents
    
    
    For j = g + 1 To g + 27
    '   Range("3:3").Select
        Vmodel = "W" & j
        Call VPF_data(Vmodel)
        Call VPF_Calc(Vmodel, Sdat, Countr)
    'Call Copy_AnyDat(Sdat, "A10", "Input")
    Next j
              
    If Countr = 0 Then: GoTo NotFound
    Call Copy_AnyDat(Sdat, "A11", "Input")
    Call SortRange("Input", "A14:Q80", "A14")
    
    SecondElasped = Round(Timer - Starttime, 2)
    Debug.Print "time required is " & SecondElasped
    
    Exit Sub
ErrHdl:
       MsgBox "Jarvis: Critical Error code 1506", vbCritical, "Jarvis - i"
      End

NotFound:
   MsgBox "Nothing found matching the specs; code:1507", vbInformation, "Friday - j"
   SecondElasped = Round(Timer - Starttime, 2)
   Debug.Print "time wasted for VPF is " & SecondElasped
End


End Sub
   
            
Public Sub VPF_Calc(Vmodel, Sdat, Optional Countr, Optional P = 0)
Application.ScreenUpdating = False
'On Error GoTo errhdl

    Dim SHName As String
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("Calc")
    SHName = "Curve"
    Dim Found As Boolean
    'SProtection SHName, False
    
    Inp_UDat Qr, Hr, Hz, Nd, Visc, ViscReq 'Rev-02 Flow is always m3/hr

    ModelRepeat = False
    NPSHRepeat = False
    m = 2
    n = 16
       
    If P <> 0 Then
     m = P
     n = P
     CopyD = True
    End If
    
For P = m To n Step 2

Nr = Round((1 - 6 / P / 100) * 2 * Hz / P, 1)
TestCrit = Nr * (Qr / 60) ^ 0.5 'Qr is always m3/hr
If TestCrit > 6000 And TestCrit < 9500 Then

    If Not ModelRepeat Then: Inp_Data Cdat, "A2", "Calc" 'RevFx
    Cal_modr Cdat, Hr, Qr, Nr, modr, Found
    If Found Then
        Cal_BDat Cdat, 0, Bdat, Nr, modr, 1, 3
        If Not NPSHRepeat Then: Inp_Data Ndat, "A15", "Calc" 'Rev Fx
        Cal_D2dat Ndat, 0, Ddat, Nr, modr, 1
        Prc_IDt Bdat, Ddat, Mdat
        Cal_Edat Mdat, Mdat, Edat, False
        Cal_Rdat Mdat, Rdat, Edat, Qr, Hr
    
        If CopyD Then
          ThisWorkbook.Sheets(SHName).Range("AJ2:BB70").ClearContents
          Call Copy_AnyDat(Edat, "AT22", SHName) 'copy efficiency data
          Call Copy_AnyDat(Mdat, "AJ26", SHName)
          Call Copy_AnyDat(Rdat, "AT3", SHName)
          Inp_Data Cdat, "AORmax", "Calc"
            Cal_R2dat Mdat, Rdat, Edat, Cdat, Qr
            Copy_AnyDat Rdat, "AZ5", SHName
        Else
            Countr = Countr + 1
            Add_Sdat Vmodel, P, Nr, modr, Rdat, Sdat
            NPSHRepeat = True
        End If
    End If
ModelRepeat = True
End If

Next P


Exit Sub
ErrHdl:
   MsgBox "Friday: Critical Error code 1508", vbCritical, "Jarvis - i"
   End

End Sub

            
Public Sub SMKP_Calc(Optional SModel, Optional rpm, Optional CopyD = True, Optional Affinity = False)
Application.ScreenUpdating = False
'On Error GoTo errhdl
    
Starttime = Timer
    Dim SHName As String
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("Calc")

    SHName = "Curve"
'    SProtection SHName, False
    ThisWorkbook.Sheets(SHName).Range("AJ2:BB70").ClearContents
    Nd = 5
    
   
    Inp_UDat Qr, Hr, Nr, Nd, Visc, ViscReq, r, Spgr  'Rev02 R is degree of flow in affinity law
    Inp_Data Cdat, "A2", "Calc"
    Inp_Data Ndat, "A15", "Calc"
            dmx = Cdat(0, 0, 1)
            dmn = Cdat(0, 0, 2)
            
    CopyE = True
    
    If Affinity Then: r = 2 'temp 20Aprl theoretica Q~N^3*d^r
    
    Cal_BDat Cdat, Nd, Bdat, Nr, dmx, dmn, r '&Rev02 R is degree of flow in affinity law
    Cal_DDat Ndat, Nd, Ddat, Nr, dmx, dmn

     Prc_IDt Bdat, Ddat, Mdat, Spgr 'Mod 23 Aprl Specific Gravity
     Cal_cutr Mdat, Qr, Hr, Nr, cutr 'Rated Dia
     Cal_Fdat Mdat, Fdat, cutr   'Calculate the Fdat parameters for cutr
     Cal_Edat Mdat, Fdat, Edat
     
     
      
If CopyE Then
     If sh.Range("isoEff") Then
        sh.Range("R42:T290").Clear 'for iso ifficiency data
        Inp_Data Cdat, "isoEff", "Calc"
        Cal_nEdat Cdat, Mdat, Edat, nEdat
        Copy_AnyDat nEdat, "P31", "Calc" 'copy iso-efficiency data
     End If
     Cal_Visc Fdat, Edat, Vdat, Nr, Visc, ViscReq
     If ViscReq Then
        Cal_Visc Mdat, Edat, Wdat, Nr, Visc, ViscReq, True
        Cal_Edat Wdat, Vdat, Edat
        Cal_cutr Wdat, Qr, Hr, Nr, cutr
        Cal_Fdat Mdat, Fdat, cutr
        Cal_Edat Mdat, Fdat, Edat
        Cal_Visc Fdat, Edat, Vdat, Nr, Visc, True
        Call Copy_AnyDat(Vdat, "AJ40", SHName) 'copy viscosity data
      End If
End If
 If ViscReq Then
 Else
  End If
 Cal_Rdat Fdat, Rdat, Edat, Qr, Hr, Spgr

If Not (CopyD) Then: Add_Sdat SModel, P, Nr, cutr, Rdat, Sdat
 
' In case of calculation without copying data
If CopyD Then
     Call Copy_AnyDat(Edat, "AT22", SHName) 'copy efficiency data
     Call Copy_AnyDat(Fdat, "AJ26", SHName)
     Call Copy_AnyDat(Rdat, "AT3", SHName)
       
     Inp_Data Cdat, "AORmax", "Calc"
     Cal_R2dat Fdat, Rdat, Edat, Cdat, Qr
     Copy_AnyDat Rdat, "AZ5", SHName
'    Cal_sysC Qr, Hr
     Cal_Fdat Mdat, Fdat, dmx 'Max Dia
     If sh.Range("isoCurveBreak") Then Lim_maxpt Fdat, Edat
     Call Copy_AnyDat(Fdat, "AJ1", SHName)
     Cal_Fdat Mdat, Fdat, dmn 'Min Dia
     Call Copy_AnyDat(Fdat, "AJ13", SHName)
End If


'   Range(sh.Cells(1, 1), sh.Cells(UBound(myArray, 1) + 1, UBound(myArray, 2) + 1)) = myArray
SecondElasped = Round(Timer - Starttime, 2)
Debug.Print "time required for SMKP is " & SecondElasped

Exit Sub
ErrHdl:
   MsgBox "Jarvis: Critical Error code 1505", vbCritical, "Friday - s"
   End
 Application.ScreenUpdating = True
 
End Sub


Sub Cal_BDat(Cdat, Nd, Bdat, Nr, dmx, dmn, r)
Dim i%, j%, k%, L%, m%

Yd = UBound(Cdat, 2)
Xd = UBound(Cdat, 3)


cc = 1
If r = 3 Then
Pow = 5
Else
Pow = 4
End If


'
'For i = 1 To Xd: Chk_unt "flow", "m3/min", Cdat(1, i): Next i

ReDim Bdat(Nd, Yd, Xd)
For k = 0 To Nd

XX = ((dmx + 0.01) - (1 - dmn + 0.03) * k / 5)
j = 0 'Diameter
    Bdat(k, j, 0) = XX
   For i = 1 To Xd: Bdat(k, j, i) = i: Next i
j = 1 'flow Rate
    Bdat(k, j, 0) = 30 + j
    For i = 1 To Xd: Bdat(k, j, i) = Cdat(0, 1, i) * XX ^ r * (Nr / Cdat(0, 4, i)) ^ 1: Next i
j = 2 'Head Diff
    Bdat(k, j, 0) = 30 + j
    For i = 1 To Xd: Bdat(k, j, i) = Cdat(0, j, i) * XX ^ 2 * (Nr / Cdat(0, 4, i)) ^ 2: Next i
j = 3 'Motor Power ip
    Bdat(k, j, 0) = 30 + j
    For i = 1 To Xd: Bdat(k, j, i) = Cdat(0, j, i) * XX ^ Pow * (Nr / Cdat(0, 4, i)) ^ 3: Next i
j = 4 'RPM
    Bdat(k, j, 0) = 30 + j
    For i = 1 To Xd: Bdat(k, j, i) = Cdat(0, j, i) * (Nr / Cdat(0, 4, i)) ^ 1: Next i
j = 5 'Hydraulic Eff
    Bdat(k, j, 0) = 30 + j
    For i = 1 To Xd: Bdat(k, j, i) = Cdat(0, j, i) + Cdat(0, j + 1, i) * (1 - XX): Next i
j = 6 'flow for Hydraulic Eff n = 1
    Bdat(k, j, 0) = 30 + j
    For i = 1 To Xd: Bdat(k, j, i) = Cdat(0, 1, i) * XX ^ cc * (Nr / Cdat(0, 4, i)) ^ 1: Next i


Next k
'Debug.Print "BDAT OK"

End Sub
Sub Cal_D2dat(Cdat, Nd, Ddat, Nr, dmx, dmn)
'On Error Got ErrHdl
Dim i%, j%, k%, L%, m%
Yd = UBound(Cdat, 2)
Xd = UBound(Cdat, 3)

ReDim Ddat(Nd, Yd, Xd)

'For i = 1 To Xd: Chk_unt "flow", "m3/min", Cdat(1, i): Next i

For k = 0 To Nd
    
XX = ((dmx + 0.01) - (1 - dmn + 0.03) * k / 5)
    'j = 0 'Diameter
    Ddat(k, 0, 0) = XX
    For i = 1 To Xd: Ddat(k, 0, i) = i: Next i
    For j = 1 To Yd: Ddat(k, j, 0) = 40 + j: Next j
    
    'j = 1 Flow rate; Cdat(3,i) = rpmr
    'j = 2 NPSh and j = 3 Nss values
    For i = 1 To Xd
    If Cdat(0, 1, i) <> 0 Then
        Ddat(k, 1, i) = Cdat(0, 1, i) * XX ^ 3 * (Nr / Cdat(0, 3, i)) 'Flow
        Ddat(k, 3, i) = Cdat(0, 3, i) * Cdat(0, 1, i) ^ 0.5 / Cdat(0, 2, i) ^ 0.75 'NSS
        Ddat(k, 2, i) = (Nr * Ddat(k, 1, i) ^ 0.5 / Ddat(k, 3, i)) ^ (4 / 3) 'NPSH
        Ddat(k, 4, i) = Nr * Ddat(k, 1, i) ^ 0.5 / Ddat(k, 2, i) ^ 0.75 'NSS New
     End If
     Next i
        
Next k
'Debug.Print "Ddat oka"

Exit Sub
ErrHdl:



End Sub

Sub Cal_DDat(Cdat, Nd, Ddat, Nr, dmx, dmn)
Dim i%, j%, k%, L%, m%

Yd = UBound(Cdat, 2)
Xd = UBound(Cdat, 3)

ReDim Ddat(Nd, Yd, Xd)

'For i = 1 To Xd: Chk_unt "flow", "m3/min", Cdat(1, i): Next i

For k = 0 To Nd
    
XX = ((dmx + 0.01) - (1 - dmn + 0.03) * k / 5)
    'j = 0 'Diameter
    Ddat(k, 0, 0) = Cdat(0, 0, 0) * XX
    For i = 1 To Xd: Ddat(k, 0, i) = i: Next i
    For j = 1 To Yd: Ddat(k, j, 0) = 40 + j: Next j
    
    'j = 1 Flow rate; Cdat(3,i) = rpmr
    'j = 2 NPSh and j = 3 Nss values
    For i = 1 To Xd
        Ddat(k, 1, i) = Cdat(0, 1, i) * (Nr / Cdat(0, 3, i))
        Ddat(k, 3, i) = Cdat(0, 3, i) * Cdat(0, 1, i) ^ 0.5 / Cdat(0, 2, i) ^ 0.75
        Ddat(k, 2, i) = (Nr * Ddat(k, 1, i) ^ 0.5 / Ddat(k, 3, i)) ^ (4 / 3)
        Ddat(k, 4, i) = Nr * Ddat(k, 1, i) ^ 0.5 / Ddat(k, 2, i) ^ 0.75
     Next i
        
Next k
'Debug.Print "Ddat oka"

End Sub

Private Sub Prc_IDt(Bdat, Ddat, Mdat, Spgr) 'Rev 23 Aprl 2018 power at sp gravity
    Dim i%, j%, jx%, k%, L%
    Dim q#
    Dim d(3), c(3) As Double
         
    jx = UBound(Bdat, 2)
    ix = UBound(Bdat, 3)
    kx = UBound(Bdat, 1)

ReDim Qex(ix), Eex(ix) As Double
For i = 1 To ix
Qex(i) = Bdat(k, 1, ix)
Next i

    ReDim Mdat(kx, jx + 1, ix)
    
    For k = 0 To kx
       
    For j = 0 To jx: Mdat(k, j, 0) = Bdat(k, j, 0): Next j
    For i = 0 To ix:  Mdat(k, 0, i) = Bdat(k, 0, i): Next i
        
       Qmx = Bdat(k, 1, ix) '11th value is max to be modified
               
        
        
       For L = 1 To ix
                q = Qmx * (L - 0.99999) / 10
                Mdat(k, 1, L) = q
            If L = 0 Or L = ix Then
                    For j = 1 To 5
                        Mdat(k, j, L) = Bdat(k, j, L)
                    Next j
            
            Else
                i = 1
                Do While q > Bdat(k, 1, i)
                    i = i + 1
                Loop
                If i = 1 Then
                ElseIf i > ix - 2 Then 'Max allowed last point in Langrange
                    i = i - 2
                End If
                
                For j = 2 To 5
                d(0) = Bdat(k, 1, i - 1):  c(0) = Bdat(k, j, i - 1)
                d(1) = Bdat(k, 1, i):      c(1) = Bdat(k, j, i)
                d(2) = Bdat(k, 1, i + 1):  c(2) = Bdat(k, j, i + 1)
                d(3) = Bdat(k, 1, i + 2):  c(3) = Bdat(k, j, i + 2)
                Mdat(k, j, L) = Lgrng(q, 3, d, c)
                
                Next j
            End If
            Mdat(k, 6, L) = Spgr * g * Mdat(k, 1, L) / 36 * Mdat(k, 2, L) / (Mdat(k, 5, L) + 0.001) 'zero eff gives div e
           Mdat(k, 6, 1) = Mdat(k, 3, 1) * Spgr
              
        ' For NPSH cal
        i = 1
        valNPSH = 0
        Do While q > Ddat(k, 1, i)
         i = i + 1
         If i > UBound(Ddat, 3) - 1 Then: Exit Do
        Loop
       If i = 1 Then: GoTo skipNPSH
       If i > UBound(Ddat, 3) - 2 Then: i = UBound(Ddat, 3) - 2
                d(0) = Ddat(k, 1, i - 1):  c(0) = Ddat(k, 2, i - 1)
                d(1) = Ddat(k, 1, i):      c(1) = Ddat(k, 2, i)
                d(2) = Ddat(k, 1, i + 1):  c(2) = Ddat(k, 2, i + 1)
                d(3) = Ddat(k, 1, i + 2):  c(3) = Ddat(k, 2, i + 2)
         
         For i = 0 To 2
         If d(i) = 0 Then: GoTo skipNPSH
         Next i
         
         valNPSH = Lgrng(q, 2, d, c)
skipNPSH:
        Mdat(k, 7, L) = valNPSH
                      
        Next L
        y1 = Mdat(k, 6, 2)
        y2 = Mdat(k, 6, 3)
        x1 = Mdat(k, 1, 2)
        x2 = Mdat(k, 1, 3)
        
            Mdat(k, 6, 1) = (y2 * x1 - y1 * x2) / (x1 - x2)
        
    Next k
'Debug.Print "Prc Idt oka"
End Sub


Sub Cal_cutr(Mdat, Qr, Hr, Nr, cutr)

Dim i%, j%, k%
Dim iq(), ix As Integer
Dim nH(), nQ() As Double
Dim x(), y() As Double

kx = UBound(Mdat, 1) 'kx = Nd max no of cut data
 
ReDim iq(kx), nQ(kx), nH(kx)
ReDim x(2), y(2)

'check which Qn() is more than Qr value
i = 1
    For k = 0 To kx
   Do While Qr > Mdat(k, 1, i)
     If i = 11 Then: GoTo ExitLoop2
     i = i + 1
   Loop
    
ExitLoop2:
        iq(k) = i - 2
    Next k
If iq(0) = 9 Then GoTo EndLoop
If iq(0) = 0 Then GoTo EndLoop

'Calculation of cut ratio
    For k = 0 To kx
        For i = 0 To 2
            x(i) = Mdat(k, 1, iq(k) + i)
            y(i) = Mdat(k, 2, iq(k) + i)
        Next i
        nH(k) = Lgrng(Qr, 2, x, y)
    Next k
    
    k = 0
    
    Do While Hr < nH(k)
        k = k + 1
        If k = kx + 1 Then: GoTo EndLoop
    Loop
    
    If k = 0 Then: GoTo EndLoop
    ix = k - 1
    If k = kx Then: ix = k - 2
    For k = 0 To 2
        x(k) = nH(ix + k)
        y(k) = Mdat(ix + k, 0, 0)
    Next k
    cutr = Lgrng(Hr, 2, x, y)

' Debug.Print cutr
    Exit Sub
    
EndLoop:
    MsgBox "Rated point is Outside confidence range of selected model", , "Jarvis"
    Application.ScreenUpdating = True
    goStoD "SMKP_Chart", "Curve"
    frmQuickSelect.Show
    End

End Sub
Sub Cal_Fdat(Mdat, Fdat, cutn)

Dim x(2), y(2) As Double
Dim XX%, i%, j%
    jx = UBound(Mdat, 2)
    ix = UBound(Mdat, 3)
    kx = UBound(Mdat, 1)

If Not cutn = 0# Then
k = 0
Do While cutn < Mdat(k, 0, 0)
k = k + 1
If k = kx + 1 Then GoTo skip
Loop



If Not k = kx + 1 Then
XX = k - 1
If k = kx Then: XX = k - 2
'Debug.Print ix
    ReDim Fdat(0, jx, ix)
    For i = 1 To ix: Fdat(0, 0, i) = i: Next i
    For j = 1 To jx - 1: Fdat(0, j, 0) = 30 + j: Next j
        
    For k = 0 To 2:  x(k) = Mdat(XX + k, 0, 0): Next k
    
    For i = 1 To ix
    For j = 1 To jx
        For k = 0 To 2: y(k) = Mdat(XX + k, j, i): Next k
        If Not (y(1) = 0 And y(2) = 0) Then: Fdat(0, j, i) = Lgrng(cutn, 2, x, y)
    Next j
    Next i
    
    Fdat(0, 0, 0) = cutn

End If

End If
Exit Sub
skip:
Debug.Print "Fdat not Calculated"

End Sub

Sub Cal_modr(Cdat, Hr, Qr, Nr, modr, Found As Boolean)

Dim i%, j%, k%
Dim Qi, Qj, Qn, Nsr, Nsi, Nsj, Nsn As Double
Dim Qbe, Ebe As Double
    ix = UBound(Cdat, 3)
    jx = UBound(Cdat, 2)

accuracy = 0


Nn = Cdat(0, 4, 2) 'Model speed
Qi = Cdat(0, 1, 3) 'Min flow rate** say. at 3
Hi = Cdat(0, 2, 3)
Qj = Cdat(0, 1, ix) 'max flow rate in input function
Hj = Cdat(0, 2, ix) 'Head Corresp to max flow

Nsr = Nr * (Qr / 3600) ^ 0.5 / Hr ^ 0.75 'Rev Qr ip is always m3/hr
Nsi = Nn * (Qi / 3600) ^ 0.5 / Hi ^ 0.75
Nsj = Nn * (Qj / 3600) ^ 0.5 / Hj ^ 0.75


If Nsj < Nsr Then: GoTo NotFound

If Nsi > Nsr Then: GoTo NotFound

ReDim x(ix), y(ix), z(ix) As Double
For i = 0 To ix - 1
x(i) = Cdat(0, 1, i + 1)
y(i) = Cdat(0, 2, i + 1)
z(i) = Cdat(0, 5, i + 1)
Next i


Call Cal_BepLgr(x, z, Qbe, Ebe)

'
'Hj = Lfunction(x, y, Qj)

Nsn = 0

Do While Not (Round(Nsn, accuracy) = Round(Nsr, accuracy))
    Qn = (Qi + Qj) / 2
    Hn = Lfunction(x, y, Qn)
    Nsn = Nn * (Qn / 3600) ^ 0.5 / Hn ^ 0.75
    Select Case Nsn
    Case Is > Nsr
      Qj = Qn
    Case Is < Nsr
      Qi = Qn
    Case Else
    GoTo NotFound
    End Select
Loop
ratio1 = (Hr / Hn) ^ 0.5 / (Nr / Nn)
ration2 = ((Qr * Nn) / (Qn * Nr)) ^ (1 / 3)


Select Case Qn 'Rev03 To reduce no of output results
Case 0.9 * Qbe To 1.1 * Qbe
Case Else
GoTo NotFound
End Select

modr = ratio1
Found = True
Exit Sub

NotFound:
Found = False

End Sub


Sub Cal_Rdat(Fdat, Rdat, Edat, Qr, Hr, Optional Spgr)

Dim XX%, i%, j%, ix%, jx%, kx%
    
    jx = UBound(Fdat, 2)
    ix = UBound(Fdat, 3)
    kx = UBound(Fdat, 1)
    
ReDim x(ix), y(ix) As Double
ReDim Rdat(kx, 3, jx + 2)

it = UBound(Edat, 3)


For k = 0 To kx

For j = 1 To jx
For i = 0 To ix
    x(i) = Fdat(k, 1, i)
    y(i) = Fdat(k, j, i)
Next i
    Rdat(k, 0, j) = 30 + j
    Rdat(k, 1, j) = Lfunction(x, y, Qr)
    Rdat(k, 2, j) = Lfunction(x, y, Edat(0, 1, it))
    
Next j
Next k
    Rdat(0, 1, 2) = Hr
'    Rdat(0, 1, 3) = Rdat(0, 1, 3) * SpGr
    jx = 8
    For i = 0 To 3
    Rdat(0, i, jx + 1) = Qr * (1 + 0.03 * (i) * (i - 1) * (i - 3))
    Rdat(0, i, jx + 2) = Hr * (1 - 0.04 * (i) * (i - 2) * (i - 3))
    Next i

End Sub




Sub Cal_R2dat(Fdat, Rdat, Edat, Cdat, Qr)
On Error GoTo CleanFail
Dim XX%, i%, j%, ix%, jx%, kx%, q#

    jx = UBound(Fdat, 2)
    ix = UBound(Fdat, 3)
    kx = UBound(Fdat, 1)
    
    it = UBound(Edat, 3)
   
ReDim x(ix), y(ix) As Double
ReDim Rdat(kx, 3, jx + 2)
count = 0
CleanRun:
    For i = 0 To ix
    x(i) = Fdat(0, 1, i)
    y(i) = Fdat(0, 2, i)
    Next i
    k = 0
    For i = 1 To 8
    Rdat(k, 0, i) = 30 + i
    XX = Int(i / 2) + 1
    q = Edat(0, 1, it) * Cdat(0, 0, XX) / 100
    arr = Array(q, Qr)
    If XX = 4 Then: q = Application.Min(arr)
    If XX = 3 Then: q = Application.Max(arr)
    If q > Fdat(0, 1, 11) Then: GoTo 10
    Rdat(k, 1, i) = q
    Rdat(k, 2, i) = Lfunction(x, y, q)
    Rdat(k, 1, i + 1) = q
    Rdat(k, 2, i + 1) = 0.001
10:
    i = i + 1
'   MsgBox "TJPmod: " & xx & "th Operating exceeds the drawing limit."
    Next i

Exit Sub
CleanFail:
count = count + 1
MsgBox "TJPmod: Clean Fail in Operating limit calculation."
Resume Next
If count = 3 Then: Exit Sub


End Sub
Sub Add_Sdat(Vmodel, P, Nr, cutm, Rdat, Sdat)
     
    jx = UBound(Rdat, 3)
    ix = UBound(Rdat, 2)
      
i = Countr + 1
k = 0
ReDim Preserve Sdat(k, 18, Countr + 1)

kt = UBound(Sdat, 3)
If i > kt Then: MsgBox "more than enough models available": Exit Sub
Sdat(k, 0, i) = P
Sdat(k, 1, i) = Vmodel
Sdat(k, 2, i) = Nr * ((Rdat(0, 1, 1) / 3600) ^ 0.5 / Rdat(0, 1, 2) ^ 0.75)
Sdat(k, 3, i) = cutm

yj = 4
'For k = 0 To kt
For j = 1 To jx - 3
Sdat(k, yj, i) = Rdat(0, 2, j)
yj = yj + 1
Next j

yj = yj + 1
Sdat(k, yj, i) = Rdat(0, 1, 1) / Rdat(0, 2, 1)
yj = yj + 1

'Next k

For j = 5 To 7
    Sdat(k, yj, i) = Rdat(0, 1, j)
    yj = yj + 1
Next j



End Sub

Sub Lim_maxpt(Fdat, Edat)
Dim i%, j%, ii%, ix%, jx%, kx%, XX#
    
    jx = UBound(Fdat, 2)
    ix = UBound(Fdat, 3)
    kx = UBound(Fdat, 1)
k = 0
XX = 1.2
ii = 1
Do While (Edat(0, 1, 6) * XX > Fdat(k, 1, ii))
    ii = ii + 1
    If ii > ix Then GoTo skip
Loop
    ii = Application.Min(ii + 1, ix)

For i = 1 To ii
For j = 0 To jx: Fdat(k, j, i) = Fdat(k, j, i): Next j
Next i
For i = ii + 1 To ix
For j = 0 To jx: Fdat(k, j, i) = 0: Next j
Next i


skip:
End Sub

Sub Cal_Sdat(Rdat, Qr, Hr, Hs)

k = (Hr - Hs) / (Qr) ^ 2

ReDim Rdat(0, 2, 10)
For i = 0 To 10
Rdat(0, 1, i) = (i + 1)
Rdat(0, 1, i) = (i / 10 * Qr) + 0.000001
Rdat(0, 2, i) = Hs + k * (i / 10 * Qr) ^ 2 + 0.00001

Next i

End Sub

Sub Cal_Edat(Mdat, Fdat, Edat, Optional fiveO = True)
    
    Dim i%, j%, jx%, k%
      
    jx = UBound(Mdat, 2)
    ix = UBound(Mdat, 3)
    kx = 5 ' UBound(Mdat, 1) mod: on 21 May 18
    
   ReDim Qbe(kx + 1), Ebe(kx + 1), Hbe(kx + 1), Nbe(kx + 1) As Double
   ReDim npsh(ix), q(ix), h(ix), Eta(ix) As Double
    
    
If fiveO Then
    For k = 0 To kx
         For i = 1 To ix
         q(i) = Mdat(k, 1, i)
         Eta(i) = Mdat(k, 5, i)
        h(i) = Mdat(k, 2, i)
        npsh(i) = Mdat(k, 7, i)
         Next i
         'Cal_BepPar Q, Eta, Qbe(k), Ebe(k)
         Cal_BepLgr q, Eta, Qbe(k), Ebe(k)
         Hbe(k) = Lfunction(q, h, Qbe(k))
         Nbe(k) = Lfunction(q, npsh, Qbe(k))
    Next k
End If

    k = 0
     For i = 1 To ix
     q(i) = Fdat(k, 1, i)
     Eta(i) = Fdat(k, 5, i)
     h(i) = Fdat(k, 2, i)
     npsh(i) = Fdat(k, 7, i)  ' correct from Mdat to Fdat
     Next i
     'Cal_BepPar Q, Eta, Qbe(kx + 1), Ebe(kx + 1)
     Cal_BepLgr q, Eta, Qbe(kx + 1), Ebe(kx + 1)
     Hbe(kx + 1) = Lfunction(q, h, Qbe(kx + 1))
      Nbe(k) = Lfunction(q, npsh, Qbe(k))
      
      
 ReDim Edat(0, 4, kx + 1)
     For k = 0 To kx + 1
      Edat(0, 0, k) = 40 + k
       Edat(0, 1, k) = Qbe(k)
       Edat(0, 2, k) = Hbe(k)
       Edat(0, 3, k) = Ebe(k)
       Edat(0, 4, k) = Nbe(k)
    Next k
Edat(0, 0, 0) = Fdat(0, 0, 0)

End Sub


Sub Cal_nEdat(Cdat, Mdat, Edat, nEdat)
    Dim a(1 To 3) As Double
    Dim x(2), y(2), z(2) As Double
    Dim q As Double
    Dim i%, j%, kk%, kkx%, kx1%, ii%, zz%, mxE%
      
    jjx = UBound(Mdat, 2)
    iix = UBound(Mdat, 3)
    kkx = UBound(Mdat, 1)
        
mxE = 0  'Max efficiency value
For i = LBound(Edat, 2) To UBound(Edat, 2)
If Edat(0, 3, i) > mxE Then: mxE = Edat(0, 3, i)
Next i

i = 1
If Int(Cdat(0, 1, 1)) <> 0 Then
    iz = UBound(Cdat, 3)
    ReDim nEdat(iz, 2, 2 * kkx + 1)
    For zz = 1 To iz:  'eta non zero
    If Cdat(0, 1, zz) <> 0 Then
    nEdat(i, 0, 0) = Cdat(0, 1, zz)
    i = i + 1
    End If
    Next zz
    iz = i - 1
GoTo LowFlow
End If


'Iso Eff valoues must be less than bep val
    iz = 6 'Six iso lines
    Effr = mxE - mxE Mod 2
    ReDim nEdat(iz, 2, 2 * kkx + 1)
    nEdat(1, 0, 0) = Effr
    For zz = 2 To iz
    nEdat(zz, 0, 0) = Effr - (zz - 1) * 2
    Next zz
    
    
LowFlow:
For zz = 1 To iz 'loop for each eff value
    eff = nEdat(zz, 0, 0)
    kx1 = 0
    
    For kk = 0 To kkx  ' loop along the LHS side of BEP
        ii = 1
        Do While Mdat(kk, 5, ii) < eff
        ii = ii + 1
        If ii > 11 Then: Exit Do
        Loop
        If ii > 11 Then: GoTo MissMe1
        If ii > 2 Then
                x(0) = Mdat(kk, 1, ii - 1): y(0) = Mdat(kk, 5, ii - 1)
                x(1) = Mdat(kk, 1, ii): y(1) = Mdat(kk, 5, ii)
                x(2) = Mdat(kk, 1, ii + 1): y(2) = Mdat(kk, 5, ii + 1)
            Parab x, y, a
            a(3) = a(3) - eff
            b = Sqr(a(2) ^ 2 - 4 * a(1) * a(3))
            q = (-a(2) + b) / (2 * a(1))
            If Not (q > x(0) And q < x(2)) Then: b = -b
            q = (-a(2) + b) / (2 * a(1))

            If q < Edat(0, 1, kk) Then
                nEdat(zz, 1, kx1) = q
                x(0) = Mdat(kk, 1, ii - 1): y(0) = Mdat(kk, 2, ii - 1)
                x(1) = Mdat(kk, 1, ii): y(1) = Mdat(kk, 2, ii)
                x(2) = Mdat(kk, 1, ii + 1): y(2) = Mdat(kk, 2, ii + 1)
                nEdat(zz, 2, kx1) = Lgrng(q, 2, x, y)
                kx1 = kx1 + 1
            End If
        End If
MissMe1:
    Next kk
    
    nEdat(zz, 0, 1) = kx1
    If kx1 < kkx Then ' Exactly on BEP
        kk = kkx
        Do While Edat(0, 3, kk) < eff
            kk = kk - 1
            If kk = 0 Then: GoTo HighFlow
        Loop
        For ii = 0 To 1
            x(ii) = Edat(0, 3, kk + ii)
            y(ii) = Edat(0, 1, kk + ii)
            z(ii) = Edat(0, 2, kk + ii)
        Next ii
        nEdat(zz, 1, kx1) = Lgrng(eff, 1, x, y)
        nEdat(zz, 2, kx1) = Lgrng(eff, 1, x, z)
    End If
    
HighFlow:
    kx1 = kx1 + 1
    For kk = kkx To 1 Step -1 ' loop along RHS side of BEP
        ii = iix
        Do While eff > Mdat(kk, 5, ii)
            ii = ii - 1
            If ii < 1 Then: GoTo MissMe2
        Loop
        
MissMe2:
        Select Case Edat(0, 3, kk)
            Case Is <= eff
            ii = 12
            Case Is > eff
            ii = 9
        End Select
       
        
        If ii < 11 Then
            ii = ii + 1
            x(0) = Mdat(kk, 1, ii - 1): y(0) = Mdat(kk, 5, ii - 1)
            x(1) = Mdat(kk, 1, ii): y(1) = Mdat(kk, 5, ii)
            x(2) = Mdat(kk, 1, ii + 1): y(2) = Mdat(kk, 5, ii + 1)
            Parab x, y, a
            a(3) = a(3) - eff
            b = Sqr(a(2) ^ 2 - 4 * a(1) * a(3))
            q = (-a(2) + b) / (2 * a(1))
            If Not (q > x(0) And q < x(2)) Then: b = -b
            q = (-a(2) + b) / (2 * a(1))

            If (q > Edat(0, 1, kk) And q < Mdat(kk, 1, 11) * 1.01) Then
                nEdat(zz, 1, kx1) = q
                x(0) = Mdat(kk, 1, ii - 1): y(0) = Mdat(kk, 2, ii - 1)
                x(1) = Mdat(kk, 1, ii): y(1) = Mdat(kk, 2, ii)
                x(2) = Mdat(kk, 1, ii + 1): y(2) = Mdat(kk, 2, ii + 1)
                nEdat(zz, 2, kx1) = Lgrng(q, 2, x, y)
                kx = kx + 1

            End If
         End If
      Next kk
  Next zz
                Debug.Print "Here"
        
End Sub

Sub Cal_BepPar(q, E, Qbe, Ebe)
    Dim i%, ix%, rw%, cl%
    Dim x(2) As Double
    Dim y(2) As Double
    Dim a(1 To 3) As Double
    rw = ActiveCell.Row
    cl = ActiveCell.Column
    Ebe = 0#
    For i = 1 To 11
        If E(i) > Ebe Then
            Ebe = E(i)
        Else
            Exit For
        End If
    Next i
    ix = i - 1
    x(0) = q(ix - 1): x(1) = q(ix): x(2) = q(ix + 1)
    y(0) = E(ix - 1): y(1) = E(ix): y(2) = E(ix + 1)
    Parab x, y, a
    Qbe = -a(2) / 2# / a(1)
    Ebe = (a(1) * Qbe + a(2)) * Qbe + a(3)
    'Y(1) = H(ix - 1): Y(2) = H(ix): Y(3) = H(ix + 1)
    'Parab X, Y, a
    'Hbe = (a(1) * Qbe + a(2)) * Qbe + a(3)

End Sub

Sub Cal_BepLgr(x, y, Xbe, Ybe)

Dim i%, ii%, j%
Dim nQ(), a, b As Double
Dim d(3), c(3) As Double

dyr = 20 'Any random number for slope
Ybe = y(2) ' Initial Max eta
iy = UBound(y)
ix = UBound(x)
iix = 20 ' Accuracy criterion
      
      ReDim nQ(iix)
 
 
    For i = 2 To iy
        If Ybe < y(i) Then
            j = i
            Ybe = y(i)
        End If
    Next i
    
    If j = iy Then
    Xbe = x(ix)
    Ybe = y(iy)
    End If
    
i = j
            d(0) = x(j - 1):  c(0) = y(j - 1)
            d(1) = x(j):      c(1) = y(j)
            d(2) = x(j + 1):  c(2) = y(j + 1)
            
     For ii = 0 To iix
        nQ(ii) = x(i - 1) + (x(i + 1) - x(i - 1)) * ii / iix
     If Not ii = 0 Then
            a = nQ(ii)
            b = nQ(ii - 1)
            dydx = (Lgrng(a, 2, d, c) - Lgrng(b, 2, d, c)) / (a - b)
            If Abs(dydx) < Abs(dyr) Then
              dyr = dydx
              Xbe = nQ(ii)
              Ybe = Lgrng(Xbe, 2, d, c)
            End If
        End If
      Next ii
 
     
 End Sub
 
Sub Cal_Visc(Fdat, Edat, Vdat, Nr, Visc, ViscReq As Boolean, Optional MdatAlso As Boolean = False)

Dim i%, j%, ix%, jx%, k%
Dim b, Qb, Hb, Eb As Double
Dim Cq, Ch(), Ce As Double
 
 Eb = Edat(0, 3, 0)
 Qb = Edat(0, 1, 0)
 Hb = Edat(0, 2, 0)
 Nb = Edat(0, 4, 0)
 ix = UBound(Fdat, 3)
 jx = UBound(Fdat, 2)
 kx = UBound(Fdat, 1)
 
string2 = "Viscosity Correction is Required; Be Sure to Enable for report generation"
string2A = "Algorithm will continue with correction factors"

If Round(Visc, 1) < 1.3 Then: GoTo NotReq
Apump = 0.1 'For end suction pumps HI 9.6.1


ReDim Vdat(kx, jx, ix)
ReDim Ch(ix)
For k = 0 To kx
    If MdatAlso Then
     Eb = Edat(0, 3, k)
     Qb = Edat(0, 1, k)
     Hb = Edat(0, 2, k)
     Nb = Edat(0, 4, k)
    End If
     
     
    b = 16.5 * (Visc ^ 0.5 * Hb ^ 0.0625) / (Qb ^ 0.375 * Nr ^ 0.25)
    If Not b > 1 Then: GoTo NotReq  'Viscosity Correction Criterion
    Ce = b ^ (-0.0547 * b ^ 0.69)
    Cq = (2.71) ^ (-0.165 * (Log(b) / Log(10)) ^ 3.15)
    If Ce > 0.99 Then: GoTo NotReq  'Viscosity Correction Criterion
    If Not ViscReq Then: MsgBox string2 & Chr(13) & string2A
    ViscReq = True
    Cnpsh = 1 + (Apump * (1 / Cq - 1) * 274000 * (Nb / Qb ^ 0.667 / Nr ^ 1.33))
    If Cnpsh < 1.1 Then Cnpsh = 1 'NPSH correction added on 16 Mar 2018
 
    Vdat(k, 0, 0) = Fdat(k, 0, 0)
    For j = 0 To jx: Vdat(k, j, 0) = 0.01: Next j
'    Debug.Print Ce & "  " & Cq
    For i = 1 To ix
        
        Vdat(k, 1, i) = Cq * Fdat(k, 1, i)
        Ch(i) = 1 - (1 - Cq) * (Fdat(k, 1, i) / Eb) ^ 0.75
        Vdat(k, 2, i) = Ch(i) * Fdat(k, 2, i)
        Vdat(k, 3, i) = Fdat(k, 3, i) / (Cq * Ch(i))
        Vdat(k, 5, i) = Fdat(k, 5, i) * Ce
         Vdat(k, 4, i) = Fdat(k, 4, i)
        Vdat(k, 6, i) = 16.3 / 60 * Vdat(k, 1, i) * Vdat(k, 2, i) / (Vdat(k, 5, i) + 0.0001)
       Vdat(k, 6, 1) = Vdat(k, 3, 1)
        Vdat(k, 7, i) = Fdat(k, 7, i) * Cnpsh 'NPSH to be added :Pending as of 23 Jan 2018
    Next i
    
    
        Vdat(k, 6, 1) = 0
        Vdat(k, 0, 4) = Ch(8)
        Vdat(k, 0, 1) = Ce
        Vdat(k, 0, 2) = Cq
Next k


Exit Sub
NotReq:
If ViscReq Then: MsgBox "Viscosity Correction Not Required; Algorithm will continue without correction factors"
ViscReq = False

End Sub
 
Function Lfunction(dXf, dYf, XX, Optional sht = False)
Dim x(2), y(2) As Double
i = 0
j = 0
'ReDim Xf(20), Yf(20)
'    If sht Then
'        With ActiveSheet
'        For Each cell In .Range(dXf): Xf(i) = cell.Value: i = i + 1: Next cell
'        For Each cell In .Range(dYf): Yf(j) = cell.Value: j = j + 1: Next cell
'        End With
'    Else
'      For i = 0 To UBound(dXf, 1): Xf(i) = dXf(i): Next i
'      For j = 0 To UBound(dYf, 1): Yf(j) = dYf(j): Next j
'    End If
'i = 0
i = 1
Do While XX > dXf(i)
i = i + 1
If i = 12 Then: GoTo ErrHdl
Loop

If i = 11 Then
ix = i - 2
Else
ix = i - 1
End If

For i = 0 To 2
x(i) = dXf(ix + i)
y(i) = dYf(ix + i)
Next i
Lfunction = Lgrng(XX, 2, x, y)

Exit Function
ErrHdl:
'MsgBox "out of Range"
Lfunction = 0

End Function


Public Function Lgrng(XX, m, x, y)
On Error GoTo errorhld
    Dim i%, j%
    Dim S#, w#

For i = 1 To m
If x(i - 1) - x(i) = 0 Then: MsgBox "LangrangianÅ@zero Error": End
Next i
    For i = 0 To m
        w = 1#
        For j = 0 To m
            If i <> j Then w = w * (XX - x(j)) / (x(i) - x(j))
        Next j
        S = S + w * y(i)
    Next i
    Lgrng = S
Exit Function
errorhld:
MsgBox "Langrangian NonZero Error", , "Jarvis"
End
End Function

Public Sub Parab(x, y, a)
'Fitting a unique second order polunomial between three points
    Dim k#
    Dim x1#, x2#, x3, y1#, y2#, y3#
    x1 = x(0): x2 = x(1): x3 = x(2)
    y1 = y(0): y2 = y(1): y3 = y(2)
    
    k = x1 * x1 * (x2 - x3) - x2 * x2 * (x1 - x3) + x3 * x3 * (x1 - x2)
    a(1) = ((x2 - x3) * y1 - (x1 - x3) * y2 + (x1 - x2) * y3) / k
    a(2) = (-(x2 * x2 - x3 * x3) * y1 + (x1 * x1 - x3 * x3) * y2 - (x1 * x1 - x2 * x2) * y3) / k
    a(3) = ((x2 * x2 * x3 - x2 * x3 * x3) * y1 - (x1 * x1 * x3 - x1 * x3 * x3) * y2 + (x1 * x1 * x2 - x1 * x2 * x2) * y3) / k
End Sub
 


Sub Inp_BDt(Bdat, Nd)
Dim i%, j%, k%
Dim Nmx%, Xd%, Yd%
ReDim Nd(2)
Nd = 5
Xd = 11
Yd = 6
ReDim Bdat(Nd, Yd - 1, Xd)
i = 0

With ThisWorkbook.Sheets("Sheet1")
    For k = 0 To Nd
        
        t = 0
        i = i + 1
        Do While Not IsEmpty(Cells(45 + i, 2))
        For j = 0 To 5
        Bdat(k, j, t) = .Cells(45 + i, 1 + j).Value
        Next j
        i = i + 1
        t = t + 1
        Loop
    Next k
End With
End Sub

