VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Private Sub CmdBack_Click()
Call goStoD("Input", "Curve")
End Sub


Private Sub CmdCalculate_Click()
Call SMKP_Data
Call SMKP_Calc

End Sub

Private Sub CmdNext_Click()

Call goStoD("Details", "Curve")

End Sub


Private Sub CommandHelp_Click()
Call HelpJarvis("NoStr2")
End Sub

Private Sub ModifyData_Click()
    frmModify.Show
End Sub

Private Sub PlotData_Click()
frmPerformance.Show
End Sub

Private Sub UserData_Click()
frmInput.Show
End Sub
Private Sub CheckBox1_Click()
    Sheets("Calc").Range("ViscosityCorrection").Value = CheckBox1.Value
    Call BtColr(CheckBox1, "Viscosity Correction")
End Sub
