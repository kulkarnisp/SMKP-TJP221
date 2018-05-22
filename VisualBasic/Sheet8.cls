VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Private Sub CmdBack_Click()
Call goStoD("Curve", "Details")
End Sub


Private Sub CmdCurvePDF_Click()
CreatePDF "Curve"
End Sub

Private Sub CmdDatPDF_Click()
CreatePDF "Details"
End Sub

Private Sub CmdNext_Click()
Call goStoD("Input", "Details")
End Sub

Private Sub CommandHelp_Click()
Call HelpJarvis("NoStr2")
End Sub


Call SProtection("Details", False)

Private Sub ToggleButton1_Click()
Call SProtection("Details", ToggleButton1.Value)
Call BtColr(ToggleButton1, "NPT")
End Sub
