VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmQuickSelect 
   Caption         =   "Quick Select"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8280
   OleObjectBlob   =   "frmQuickSelect.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmQuickSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub UnitMod()

 array1 = Split("flow,head,speed", ",")
 For i = 0 To 2
 val3 = Me.Controls("TB" & i + 1).Value
 fu3 = Me.Controls("CB" & i + 1).Value
 Chk_unt array1(i), fu3, val3
 Me.Controls("TB" & i + 1).Value = val3
 Me.Controls("CB" & i + 1).Value = fu3
 Next i
 
'
''Me.CB1.Value = "m3/hr"
' Chk_unt "head", Me.CB2.Value, Me.TB2.Value
''Me.CB2.Value = "m"
' Chk_unt "speed", Me.CB3.Value, Me.TB3.Value
''Me.CB3.Value = "rpm"
''
End Sub

Sub PowerMod()
liqP = Me.TB2.Value * Me.TB1.Value * 9.81 / (3600 * 0.7) * 1.125
arr = Array(liqP, Me.TB4.Value)
Me.TB4.Value = Int(Application.Max(arr) + 2.1)
Me.Power.Value = Round(liqP, 2)

End Sub

Sub ModelMod()
With Sheets("Calc")
Me.model.Caption = .Range("model").Value
Me.Series.Caption = .Range("Series").Value
End With
End Sub

Private Sub CheckBox1_Click()
Sheets("Calc").Range("ViscosityCorrection").Value = Me.CheckBox1.Value
End Sub

Private Sub cmdCurrent_Click()
Module1.TrSheettoForm "Calc", "J13", Me, True
PowerMod
UnitMod
ModelMod

End Sub

Private Sub CmdSearch_Click()
'FAGSHKJBFkjbdsaufg
End Sub

Private Sub CommandButton1_Click()
PowerMod
End Sub

Private Sub CommandButton2_Click()
PowerMod
End Sub

Private Sub CommandClose_Click()
Call EnterData_Click
Unload Me
End Sub


Sub EnterData_Click()
Msgr = "Selected Model does not exist!"
If Not (SMKPChecklist(Me.TB22.Value)) Then: MsgBox Msgr: Exit Sub
TrSheettoForm "Calc", "J13", Me, False
TrSheettoForm "Calc", "J13", Me, True
ModelMod
'
'Application.Wait (True)

End Sub




Private Sub ToggleButton1_Click()

Call BtColr(Me.ToggleButton1, "AutoSelect")
Me.TB22.Visible = Not (Me.ToggleButton1.Value)
Me.LB3.Visible = Not (Me.ToggleButton1.Value)
Me.TB3.Visible = Not (Me.ToggleButton1.Value)
Me.CB3.Visible = Not (Me.ToggleButton1.Value)
Me.CmdSearch.Visible = (Me.ToggleButton1.Value)

End Sub
Sub UserForm_Initialize()
TrSheettoForm "Calc", "J13", Me, True
ModelMod
End Sub

Private Sub UserForm_Click()

'Call cmdCurrent_Click

End Sub
