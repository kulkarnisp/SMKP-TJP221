Attribute VB_Name = "TJPaction"
 Public rCountr%
 Public rSdat() As String
    


Sub SearchModel(Series)

Select Case Series
Case Is = "SMKP"
MsgBox "SMKP Auto search disabled", vbCritical
End
SMKP_Search
Case Is = "VPF"
VPF_Search
Case Else
MsgBox "Data Unavailable for series " & Series
End
End Select

End Sub
Sub CalSelected(Series, modelid, P)

Select Case Series
Case Is = "SMKP"

MsgBox "SMKP Auto search disabled", vbCritical
End

SModel = modelid
Nr = P
SMKP_Data SModel
SMKP_Calc SModel, Nr, False
Case Is = "VPF"
    Vmodel = modelid
    Call VPF_data(Vmodel)
    Call VPF_Calc(Vmodel, Sdat, , P)
Case Else
MsgBox "Data Unavailable for series " & Series
End
End Select

End Sub

Sub OnSpecs()

Dim WS, OSheet As Worksheet

   Set WS = ThisWorkbook.Sheets("Input")
    
    ActRow = ActiveCell.Row
    ActCol = ActiveCell.Column
    With WS
    If .Cells(ActRow, 1).Value = "" Then
        Exit Sub
    End If
    If .Cells(ActRow, 2).Value = "" Then
        Exit Sub
    End If

    
    modelid = .Cells(ActRow, 2).Value
    P = .Cells(ActRow, 1).Value
    Series = Range("Series").Value
    
    Call CalSelected(Series, modelid, P)
        

    .Range("13:100").Select
    'Range("L35").Activate
    Selection.Interior.ColorIndex = xlNone
    Selection.Font.ColorIndex = 0
    
    j = 1

    Do While j <> 17
    .Cells(ActRow, j).Select
    With Selection.Interior
        .ColorIndex = 10
        .Pattern = xlSolid
    End With
         Selection.Font.ColorIndex = 2
    j = j + 1
    Loop


   .Activate
    .Cells(ActRow, 2).Select
    
   End With
End Sub
Public Sub Chk_unt(Qty, Unit, val)
On Error GoTo ErrHdl
Msgx = "This unit is unavailable select again."
Select Case Qty  'converts any input unit of flow to m3/hr
    Case Is = "flow"
        Select Case Unit
        Case Is = "m3/hr"
        val = val
        Case Is = "m3/min"
        val = val * 60
        Case Is = "m3/sec"
        val = val * 3600
        Case Is = "m3/sec"
        val = val * 3600
        Case Is = "Gpm"
        val = val * 60 / 264.172
        Case Else
        MsgBox Msgx
        End
        End Select
        Unit = "m3/hr"
        
Case Is = "head"
    Select Case Unit
    Case Is = "m"
    val = val
    Case Is = "ft"
    val = val / 3.048
    Case Else
    MsgBox Msgx
    End
    End Select
       Unit = "m"

Case Is = "speed"
    Select Case Unit
    Case Is = "rpm"
    val = val
    Case Is = "Hz"
    val = val * 60
    Case Is = "rad/s"
    val = val * 60 / (2 * Pi)
    Case Else
    MsgBox Msgx
    End
    End Select
       Unit = "rpm"
    
End Select

'Chk_unt = val
Converted:
Exit Sub
ErrHdl:
MsgBox "Unit Conversion Error; Reselect Unit", vbCritical
End
End Sub

Function Cnv_unt(Qty, Unit, val)
On Error GoTo ErrHdl
Msgx = "This unit is unavailable select again."
Select Case Qty  'And Converts m3/hr to AnyUnit
Case Is = "flow"
    Select Case Unit
    Case Is = "m3/hr"
    val = val
    Case Is = "m3/min"
    val = val / 60
    Case Is = "m3/sec"
    val = val / 3600
    Case Is = "m3/sec"
    val = val / 3600
    Case Is = "Gpm"
    val = val / 60 * 264.172
    Case Else
    MsgBox Msgx
    End
    End Select
    
Case Is = "head"
    Select Case Unit
    Case Is = "m"
    val = val
    Case Is = "ft"
    val = val * 3.048
    Case Else
    MsgBox Msgx
    End
    End Select


Case Is = "speed"
    Select Case Unit
    Case Is = "rpm"
    val = val
    Case Is = "Hz"
    val = Int(val / 60)
    Case Is = "rad/s"
    val = val / 60 / (2 * Pi)
    Case Else
    MsgBox Msgx
    End
    End Select
    
    
End Select

Cnv_unt = val
Converted:
Exit Function
ErrHdl:
MsgBox "Unit Conversion Error; Reselect Unit", vbCritical
End
End Function


Function SMKPChecklist(modelid)

ReDim ModelArray(46)
'If not(me.TB21.Value="SMKP") then

For i = 1 To 46
ModelArray(i) = Sheets("Calc").Cells(28 + i, 1).Value
Next i
For i = 1 To 46
If ModelArray(i) = modelid Then: GoTo 10
Next i
SMKPChecklist = False
Exit Function

10:
SMKPChecklist = True

End Function


Sub goStoD(d As String, Leave As String, Optional OSplit As Boolean = True, Optional OLock As Boolean = False)
Application.ScreenUpdating = False
    On Error Resume Next
    If Not (OLock) Then: GoTo 10
    itt = InputBox("enter data-access-code")
    If Not (itt = "noriya") Then GoTo 20
10:
    Source = Leave
    sh = d
    Sheets(sh).Visible = True
    Sheets(sh).Select
    'Sheets(Sh).Cells(5, 11).Select
    Sheets(Source).Visible = xlVeryHidden
    Sheets(sh).Range("C4").Select
    ActiveWindow.Zoom = 70
'    With ActiveWindow
'        .SplitColumn = 0
'        .SplitRow = 0
'    End With
    ActiveCell.Offset(15, -2).Range("A1").Select
    If OSplit Then
    With ActiveWindow
'        .SplitColumn = 31.2127659574468
''        .SplitRow = 35.7142857142857
'        .Panes(1).ScrollRow = 1
'         .Panes(1).ScrollColumn = 1
'       .Panes(4).Activate
'      .Panes(4).ScrollRow = 112
'      .Panes(2).ScrollColumn = 32
     
    End With
    End If
'   Call SProtection(d, True)
Application.ScreenUpdating = True

20:
End Sub

Sub HelpJarvis(Optional ScCrit As String)

MsgBox "Help document under development!", vbInformation, "Jarvis-info"

End Sub

Sub BtColr(Butn As Object, SCrit As String)
With Butn
If .Value Then
.ForeColor = RGB(0, 0, 0)
.BackColor = RGB(0, 255, 0)
.Caption = "Toggle on Visible."
Else
.ForeColor = RGB(0, 0, 0)
.BackColor = RGB(255, 255, 255)
.Caption = SCrit
End If
End With

End Sub
Sub SetPrintArea(SHName, ScRg)
With ThisWorkbook.Sheets(SHName)
    
    .PageSetup.PrintArea = ""
    .PageSetup.PrintArea = ScRg

End With
End Sub
Sub SProtection(SHName As String, Optional Tskip = True)
    On Error Resume Next
    Dim Target As Range
        Dim sh As Object
    Set sh = Sheets(SHName)
    sh.Unprotect Password:="noriya"
    TG = Sheets(SHName).PageSetup.PrintArea
    don = Chr(34) & TG & Chr(34)
    Set Target = Range(TG)
If Tskip Then
    For Each cell In Target
            cell.Select
            If cell.HasFormula Then
                With Selection
                    .Locked = True
                    .FormulaHidden = True
                End With
            Else
                With Selection
                    .Locked = False
                    .FormulaHidden = True
                End With
            End If
        Next cell
        sh.Protect Password:="noriya"
End If

End Sub

Sub SortRange(SHName, ScRg, keyRange)
Dim sh As Worksheet
Set sh = ThisWorkbook.Worksheets(SHName)

sh.Sort.SortFields.Clear
sh.Sort.SortFields.Add Key:=Range(keyRange), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With sh.Sort
        .SetRange Range(ScRg)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub


