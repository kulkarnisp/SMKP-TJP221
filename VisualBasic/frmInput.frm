VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInput 
   Caption         =   "Input Data"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8145
   OleObjectBlob   =   "frmInput.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Sub UserForm_Initialize()
Call cmdCurrent_Click
End Sub

Private Sub cmdCurrent_Click()
Call Module1.TrSheettoForm("Calc", "J3", Me, True)
Me.TB34.Value = Environ$("username")
End Sub


Private Sub CmdImport_Click()
Call Module1.TrSheettoForm("Calc", "J3", Me, True)
Call Module1.TrHistorytoForm("History", ComboHistory.Value, Me, True)
End Sub


Private Sub CmdSave_Click()
Call Module1.TrExportD("Calc", "J2:L39", ".txt")
End Sub


Private Sub CommandButton1_Click()
        On Error Resume Next
        For i = 1 To 40: Me.Controls("TB" & i).Value = ""
        Next i
End Sub


Private Sub CommandButton2_Click()
    Call Module1.TrImportD("Calc", "J4", "txt")
End Sub


Private Sub CommandClose_Click()
  Unload Me
 End Sub

Public Sub EnterData_Click()

checkEmpty Me, Bval
If Bval Then: GoTo 20

    Call Module1.TrHistorytoForm("History", Me.TB4.Value, Me, False)
    Call Module1.TrSheettoForm("Calc", "J3", Me, False)
    Unload Me
20:
End Sub

Sub checkEmpty(xe As Object, Bval)
On Error Resume Next
With xe
For i = 1 To 23
checker = 10
checker = .Controls("TB" & i).Value
If checker = "" Then: GoTo IsZero
Next i
End With

Exit Sub
IsZero:
Namer = xe.Controls("LB" & i).Caption
essage = "Field " & Namer & " must be a non zero value"
Bval = MsgBox(essage, vbInformation, "Friday")
Bval = True

End Sub

Sub RequiredCtrl(frm As Object)
'  Sets background color for required field -> Tag = *
Dim setColour As String
setColour = RGB(255, 244, 164)
Dim ctl As control
End
For Each ctl In frm.Controls
'        If ctl.ControlType = acTextBox Or ctl.ControlType = acComboBox Or ctl.ControlType = acListBox Then
            If InStr(1, ctl.Tag, "*") = "" Then
            Debug.Print "REACH"
                ctl.BackColor = setColour
            End If

Next ctl
Set ctl = Nothing
End Sub
Private Sub UserForm_Click()
'
'Call cmdCurrent_Click

End Sub
