VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CmdAddM_Click()

Call goStoD("Input", "Selection")

End Sub

Private Sub CmdDelete_Click()
Call DeleteChart("Selection", 1)
End Sub

Private Sub CmdHelp_Click()
Call HelpJarvis("Select")
End Sub

Private Sub CmdPlot_Click()

Call CopyCurve("Calc", "Selection")
Call DelinkDataFromChart("Selection", 1)

End Sub

Private Sub CommandButton1_Click()
Dim bApply As Boolean
valer = Sheets("calc").Range("Series").Value

'If valer = "VPF" Then
Massage = "Auto-Search models for series " & valer
'Else
'Massage = "Only defined for VPF series and current series is " & valer
'MsgBox Message, vbCritical
'End
'End If

bApply = (MsgBox(Massage, vbYesNo, "TJPSearch") = vbYes)
Debug.Print bApply
If bApply Then
Call TJPaction.SearchModel(valer)
End If

End Sub

Private Sub CommandButton2_Click()
'On Error Resume Next
'Call goStoD("Input", "Selection")
'For I = 1 To Sheets("Selection").ChartObjects.count - 1
''Sheets("Selection").ChartObjects.Item(I).Select
''Call DeleteChart("Selection", 1)
'Next I
End Sub

Private Sub PlotVPF_Click()
    Call OnSpecs
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
On Error Resume Next
n = ActiveSheet.ChartObjects.count
k = ActiveSheet.Range("C4").Value
ActiveSheet.ChartObjects.Item(k).ShapeRange.ZOrder msoBringToFront

End Sub
