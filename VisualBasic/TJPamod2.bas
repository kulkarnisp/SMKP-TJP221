Attribute VB_Name = "TJPamod2"
Public List() As String
Public model As String
Public Cdat() As Double

Public NCtr%

Sub SMKP_Search()
Application.ScreenUpdating = False

'On Error GoTo ErrHdl
'
'Call Module2.SheetVisible(True, "SMKP_Data")
    Dim val As String
    Dim Qr, Hr
    ReDim List(0, 1, 10)
    val = ""
    SHName = "Selection"
    
    Qr = Sheets(SHName).Range("E26").Value
    Hr = Sheets(SHName).Range("E25").Value
    
    Range("flow") = Qr
    Range("head") = Hr
'   LoadFluid Water
    Range("SpGr") = 1
    Range("viscosity") = 1
    Range("ViscosityCorrection") = False
   
   Call SearchJordan(List, Qr, Hr, 3)
   Debug.Print List(0, 0, 0)
   
    Select Case Range("Hz")
    Case Is = 50
        i = 0
    Case Is = 60
        i = 1
    Case Else
        i = 0
    End Select
    
   Range("Speed") = List(0, 1, i)
   modelid = List(0, 0, i)
   
   If modelid = "" Then: GoTo Nomodel
      
   Call SMKP_Data(modelid)
   Call SMKP_Calc
      
   ' Power selection at end of curve value
    Range("Power").Value = SelectMotor(Sheets("Curve").Range("AP38").Value)
   
    Application.ScreenUpdating = True
    
    M1 = "Total Models found:- " & NCtr
    M2 = "Model Selected:- " & modelid
    M5 = "Data caculated rated dia:- " & Range("cutdia").Value
    M3 = "Data caculated rated dia:- " & Range("cutdia").Value
    M4 = "Proceed to plot the curve?"
    
    BoolYB = MsgBox(M1 & Chr(13) & M2 & Chr(13) & M3 & Chr(13) & M4, vbYesNoCancel, "Friday-info")
    If BoolYB = vbYes Then: goStoD "Curve", "Selection"
  ' Call TJPintact.CreatePDF("Curve")
     


Exit Sub
ErrHdl:
MsgBox "SearchError", vbCritical
End

Nomodel:
MsgBox "No Model Currently available", vbMsgBoxHelpButton
goStoD "SMKP_Chart", "calc"
End

End Sub
'Sub MHI()
'
'
'
'
'End Sub
 
Sub TataONGC(TR As Boolean)
    Application.ScreenUpdating = True
    Application.DisplayAlerts = False
    
    Dim Namer As String
    array2 = Split("P-155 A/B,P-157 A/B,P-228 A/B,P-229 A/B,P-316 A/B,P-317 A/B,", ",")
    
    For i = 0 To UBound(array2) - 1
    
    Call Module1.TrSheettoForm("Calc", "J3", frmInput, True)
    Namer = array2(i)
    Call Module1.TrHistorytoForm("History", Namer, frmInput, True)
    
    
    frmInput.TB14.Value = 5.5
    
    Call Module1.TrHistorytoForm("History", frmInput.TB4.Value, frmInput, False)
    Call Module1.TrSheettoForm("Calc", "J3", frmInput, False)
    Unload frmInput
    Call SMKP_Data
    Call SMKP_Calc
    Call TJPintact.CreatePDF("Curve")
    Call TJPintact.CreatePDF("Details")
    
    Next i
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub
Function SelectMotor(Power)
n = 1.1 'Power Multiplier factor

Call Inp_Data(Cdat, "CD34", "SMKP_Data")
kx = UBound(Cdat, 3)
'Bubble sort PList

ReDim PList(kx) As Double
For i = 0 To kx
PList(i) = Cdat(0, 1, i)
Next i

Power = Power * 1.1
For i = 0 To kx
If PList(i) > Power Then GoTo Found
Next i

i = i - 1

Found:
SelectMotor = PList(i)

End Function
Sub SearchJordan(List, Xb, Yb, Optional Na = 3)
'List is output list containig models
'Xb,Yb are operating paramenters
'
Dim Cdat() As String
Dim XX() As Double
Dim YY() As Double
Dim List1(3, 1, 5) As String
Dim Xa(3), Ya(3) As Double
Dim Nsp(3) As Double
sttime = Timer

NCtr = 0
SHName = "Selection"
    
    a = False
    If Not (Yb <> 0) Then: a = True
    If Not (Xb <> 0) Then: a = True
    
    If a Then
    MsgBox "Zero Input Values", vbCritical, "Jarvis"
    End
    End If

'Nsp = Split("2950#1470#3550#1770", "#")

For j = 0 To Na
    Xa(j) = ((Xb / 3600) / Yb ^ 0.5) ^ 0.5
    'Ya(j) = Yb ^ 0.5 / Nsp(j) * 100
Next j

ScRg = "CM1"
Sh2Name = "SMKP_Data"
For ii = 1 To 45
        
    ScRg = Tr3Data(ScRg, Sh2Name)
    Inp_Data Cdat, ScRg, Sh2Name
    Mxrt = UBound(Cdat, 3)
    
    ReDim XX(Mxrt), YY(Mxrt)
    For i = 1 To Mxrt
    XX(i - 1) = Cdat(0, 1, i)
    YY(i - 1) = Cdat(0, 2, i)
    Next i
    modelRPM Cdat(0, 4, 2), Nsp
    
    For j = 0 To Na
        If Nsp(j) = 0 Then
        Ya(j) = 0
        Else
        Ya(j) = Yb ^ 0.5 / Nsp(j) * 100
        End If
        
        If InPolygon(XX, YY, Xa(j), Ya(j)) Then
        List1(j, 0, k) = Cdat(0, 4, 2)
        List1(j, 1, k) = Nsp(j)
        k = k + 1
        NCtr = k 'Genreal counter of matches
        End If
    Next j

Next ii

    k = 0
    For j = 0 To Na
    For i = 0 To 5
    If List1(j, 0, i) = "" Then
    Else
    List(0, 0, k) = List1(j, 0, i)
    List(0, 1, k) = List1(j, 1, i)
    k = k + 1
    End If
    Next i
    Next j

'    Range("S31").Select
'    ActiveSheet.ListObjects("Table").Name = "TableSMK"
Elasped = Round(Timer - sttime, 2)
Debug.Print "Time Req " & Elasped
Copy_AnyDat List, "AG3", "Curve"

Exit Sub


End Sub
Sub LoadFluid(ScRt)

   Range("Viscosity").Value = 1
   Range("SpGr") = 1
   Range("Fluid").Value = "Water"
   
End Sub

Function Tr3Data(ScRg, SHName)
'On Error GoTo ErrHdl

With ThisWorkbook.Sheets(SHName)
rn = .Range(ScRg).Row
cn = .Range(ScRg).Column
i = 0
If .Cells(rn + i, cn).Value = 0 Then Exit Function

Do While True

    If .Cells(rn + i, cn).Value = .Cells(rn + i + 1, cn).Value Then
            i = i + 1
            If i = 15 Then: GoTo ErrHdl
    Else
    Tr3Data = .Cells(rn + i + 1, cn).Address
    Exit Function
    End If

Loop

End With

Exit Function

ErrHdl:
Debug.Print "Error Tr3"
End


End Function

Sub modelRPM(idmodel, Nsp)

With Sheets("SMKP_Data")
ScRg = "BQ1"

ScRg = Tr2Data(idmodel, ScRg, "SMKP_Data")

For i = 0 To 3
Nsp(i) = .Range(ScRg).Offset(, 5 + i).Value
Next i


End With
End Sub

Function Tr2Data(TrNAme, ScRg, SHName)
'Finds row of TrName in following rows after "ScRg2" in sheets SHName

With Sheets(SHName)
rn = .Range(ScRg).Row
cn = .Range(ScRg).Column
i = 0
If .Cells(rn + i, cn).Value = 0 Then Exit Function

Do While True
    If Not (.Cells(rn + i, cn).Value = TrNAme) Then
    i = i + 1
If i = 50 Then: GoTo ErrHdl
    Else
    GoTo Number
    End If
Loop

Number:
Tr2Data = .Cells(rn + i, cn).Address

End With
Exit Function

ErrHdl:
Debug.Print "Error Tr2"
End

End Function
