Attribute VB_Name = "OutputSubs"

'Callback for customButton1 onAction
Sub dokecurve(control As IRibbonControl)
   Call SProtection("Curve", False)
   Call goStoD("Curve", "Selection")
End Sub

'Callback for customButton2 onAction
Sub dokesheet(control As IRibbonControl)
    Call goStoD("Details", "SMKP_Chart")
End Sub

'Callback for customButton3 onAction
Sub dokechart(control As IRibbonControl)
    Call goStoD("SMKP_Chart", "Curve")
'    MsgBox "Executed Successfully"
End Sub

'Callback for customButton4 onAction
Sub Macro5(control As IRibbonControl)
    Call goStoD("Input", "Details")
    Call goStoD("Input", "Curve")
    MsgBox "etc executed"
End Sub

'Callback for customButton5 onAction
Sub Macro4(control As IRibbonControl)
    Call goStoD("Selection", "Curve")
    
End Sub

'Callback for customButton6 onAction
Sub Macro6(control As IRibbonControl)
    Call goStoD("ShPivot", "Curve")
End Sub

'Callback for customButton7 onAction
Sub Macro7(control As IRibbonControl)
   'Calculate
   Call SMKP_Data
   Call SMKP_Calc
End Sub

'Callback for customButton8 onAction
Sub Macro8(control As IRibbonControl)
Call goStoD("SMKP_Data", "Details", , True)
End Sub

'Callback for customButton9 onAction
Sub Macro9(control As IRibbonControl)
Call goStoD("History", "Input", , True)
End Sub

'Callback for customButton10 onAction
Sub Macro10(control As IRibbonControl)
    Call goStoD("Calc", "Details", , True)
End Sub

'Callback for customButton11 onAction
Sub Macro11(control As IRibbonControl)
    Call goStoD("VPF_Data", "Calc")

End Sub

'Callback for customButton12 onAction
Sub Macro12(control As IRibbonControl)
'    Call SMKP_Plot
    
End Sub

'Callback for customButton4B onAction
Sub Macro4B(control As IRibbonControl)
 Call goStoD("Input", "Calc")
Call TrImportD("Calc", "A1", ".csv")
End Sub

'Callback for customButton5B onAction
Sub Macro5B(control As IRibbonControl)
Call TrImportD("Calc", "J4", ".txt")
End Sub

'Callback for customButton6B onAction
Sub Macro6B(control As IRibbonControl)
End Sub

'Callback for customButton4A onAction
Sub Macro4A(control As IRibbonControl)
frmPerformance.Show
End Sub

'Callback for customButton5A onAction
Sub Macro5A(control As IRibbonControl)
frmInput.Show
End Sub

'Callback for customButton6A onAction
Sub Macro6A(control As IRibbonControl)
frmModify.Show
End Sub

Sub MacroClose(control As IRibbonControl)
On Error GoTo msgbv
ActiveSheet.Visible = xlVeryHidden
Exit Sub
msgbv:
MsgBox "Climax:No more endings"
End Sub

'Callback for customButton10 onAction
Sub Macro14(control As IRibbonControl)
    SetPrintArea "Curve", "A2:AE81"
   CreatePDF "Curve"
End Sub

'Callback for customButton11 onAction
Sub Macro15(control As IRibbonControl)
    CreatePDF "Details"
End Sub

'Callback for customButton12 onAction
Sub Macro16(control As IRibbonControl)
    RunMacro "Export"
End Sub

Sub MacroQuickSelect(control As IRibbonControl)
    frmQuickSelect.Show
End Sub
'Sub isalkhd()
'Dim i%, xx%
'i = 1#
'xx = i / 2
'
'
'
'End Sub

Sub Macro17(control As IRibbonControl)

Range("theory").Value = Not (Range("theory").Value)

End Sub
'Callback for ViscToggle getPressed
Sub ViscToggle_getPressed(control As IRibbonControl, ByRef returnedVal)
End Sub

'Callback for ViscToggle onAction
Sub ViscToggle_onAction(control As IRibbonControl, pressed As Boolean)

Range("ViscosityCorrection").Value = pressed
End Sub


'Callback for Labley
Sub Macro18(control As IRibbonControl)
MsgBox "Jarvis"
End Sub

'Callback for Label
Sub Macro19(control As IRibbonControl)
MsgBox "Friday"
End Sub

'Callback for customButton20 onAction
Sub Macro20(control As IRibbonControl)
End Sub

'Callback for customButton21 onAction
Sub Macro21(control As IRibbonControl)
End Sub

'Callback for customButton22 onAction
Sub Macro22(control As IRibbonControl)
End Sub

'Callback for FaceHappy onAction
Sub Macro23(control As IRibbonControl)
With Sheets("Calc")
.PivotTables("PvTb2").PivotCache.Refresh
.PivotTables("PvTb1").PivotCache.Refresh
End With
MsgBox "Adamantium"

End Sub
'Callback for FilesTR onAction
Sub Macro13(control As IRibbonControl)
End Sub
'Callback for Hz50Toggle getPressed
Sub freq_getPressed(control As IRibbonControl, ByRef returnedVal)
End Sub

'Callback for Hz50Toggle onAction
Sub IndiaFreq_onAction(control As IRibbonControl, pressed As Boolean)
Range("Hz").Value = 50
End Sub

'Callback for Hz60Toggle onAction
Sub JapFreq_onAction(control As IRibbonControl, pressed As Boolean)
Range("Hz").Value = 60
End Sub



