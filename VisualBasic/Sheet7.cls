VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub CmdBack_Click()
Call goStoD("Header", "Input")
End Sub

Private Sub CmdCalculate_Click()
Call SMKP_Data
Call SMKP_Calc

End Sub

Private Sub CmdNext_Click()
Application.ScreenUpdating = False
'Call Module1.Calculate
Call SProtection("Curve", False)
Call goStoD("Curve", "Input")
Application.ScreenUpdating = True
End Sub


Private Sub CmdPreSelection_Click()
Call goStoD("ShPivot", "Input")
End Sub

Private Sub CmdPlotData_Click()
Call TrImportD("Calc", "L1", ".csv")
End Sub

Private Sub CmdQuick_Click()
frmQuickSelect.Show
End Sub

Private Sub CmdSelect_Click()
Call goStoD("Selection", "Input")
End Sub

Private Sub CmdUserData_Click()
Call TrImportD("Calc", "AQ4", ".txt")
End Sub

Private Sub CmdHelp_Click()
Call HelpJarvis
End Sub


Private Sub CommandButton1_Click()
Call goStoD("SMKP_Chart", "Input")
End Sub

Private Sub ToggleButton1_Click()

Sheets("Input").ChartObjects("SMKPchart").Visible = ToggleButton1.Value
Call BtColr(ToggleButton1, "Show Plot")

End Sub

Private Sub ToggleButton2_Click()

Sheets("Calc").Range("ZZ4").Value = ToggleButton4.Value
Call BtColr(ToggleButton2, "AutoBot")

End Sub

Private Sub ToggleButton3_Click()
Call SProtection("Selection", ToggleButton3.Value)
Call BtColr(ToggleButton3, "JarvisBot")
End Sub


