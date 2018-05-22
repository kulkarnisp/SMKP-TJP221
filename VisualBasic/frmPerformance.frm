VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPerformance 
   Caption         =   "Performance Data"
   ClientHeight    =   9435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9975
   OleObjectBlob   =   "frmPerformance.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmPerformance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdSave_Click()
Call Module1.TrExportD("Calc", "A1:H26", ".csv")
End Sub

Private Sub CommandButton1_Click()

Call Module1.TrCurveData0("Calc", "A1", Me)

End Sub

Private Sub CommandClose_Click()
  Unload Me
 End Sub

Private Sub EnterData_Click()

If Trim(Me.TextBox1.Value) = "" Then
Me.TextBox1.SetFocus
MsgBox "Please enter Project Number number"
Exit Sub
End If

Call Module1.TrCurveData1("Calc", "A1", Me)
Call CommandClose_Click

End Sub

Private Sub ImportDef_Click()
 Call Module1.TrCurveData0("Calc", "A1", Me)
End Sub

Sub UserForm_Initialize()
 Call Module1.TrCurveData0("Calc", "A1", Me)
End Sub

