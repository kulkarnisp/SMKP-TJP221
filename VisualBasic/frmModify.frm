VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmModify 
   Caption         =   "Modifications"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7290
   OleObjectBlob   =   "frmModify.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmModify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub PlotUnit()
Dim sh As Worksheet
Set sh = ThisWorkbook.Sheets(3)
If Not Me.CB11.Value = "" Then
    Multx = 1
    InpU = Sheets("Calc").Range("funit").Value
    OutU = Me.CB11.Value
    Call Chk_unt("flow", InpU, Multx) ' Checks if everything is in m-hr standard
    Call Cnv_unt("flow", OutU, Multx) 'Converts m-hr to putput unit
    
With sh
    For i = 2 To 60: .Cells(i, 37).Value = .Cells(i, 37).Value * Multx: Next i
    For i = 6 To 15: .Cells(i, 53).Value = .Cells(i, 53).Value * Multx: Next i
    For j = 46 To 49: .Cells(5, j).Value = .Cells(5, j).Value * Multx: Next j
    For j = 46 To 49: .Cells(13, j).Value = .Cells(13, j).Value * Multx: Next j
End With

Sheets("Calc").Range("flow").Value = Sheets("Calc").Range("flow").Value * Multx
Sheets("Calc").Range("funit").Value = OutU

End If
End Sub


Sub AddNPSH()
'Sheets("Calc").Range("shNPSH").Value = ToggleButton1.Value
Dim sh As Worksheet
Set sh = ThisWorkbook.Sheets(3)
If Not Me.NPSHA.Value = "" Then
    Multx = 1
    Naval = Me.NPSHA.Value
 Else
 Naval = ""
End If
With sh
    For i = 28 To 38:
    If .Cells(i, 43).Value = "" Then
    Else
    .Cells(i, 44).Value = Naval:
    End If
    Next i
End With
Sheets("Calc").Range("npsha").Value = Naval
End Sub
Sub ChangeVisc()
Dim Vdat()
Chang = CB1.Value
If Sheets("Calc").Range("ViscosityCorrection").Value Then
    If Chang = True Then
        Call Inp_Data(Vdat, "AJ41", "Curve")
      
        Call Inp_Data(Edat, "AT23", "Curve")
        Qr = Sheets("Calc").Range("flow").Value
        Hr = Sheets("Calc").Range("head").Value
        
        Call Cal_Edat(Vdat, Vdat, Edat, False)
        Call Cal_Rdat(Vdat, Rdat, Edat, Qr, Hr)
        Copy_AnyDat Rdat, "AT3", "Curve"
    ElseIf Chang = False Then
    
    Else
    MsgBox "Select Pull down Value"
    End If
End If
End Sub


Sub RealTheory()

Call TJPamod.SMKP_Calc(, , , True)
Application.ScreenUpdating = True
Sheets("Calc").Range("isoAffiniity").Value = False

End Sub


Sub DrawSysCurve()
Dim sh As Worksheet
Set sh = ThisWorkbook.Sheets("Calc")
Dim Hst As Double
If Not Me.Hs.Value = "" Then: Hst = Me.Hs.Value
Qr = sh.Range("flow").Value
Hr = sh.Range("head").Value
Cal_Sdat Rdat, Qr, Hr, Hst
Copy_AnyDat Rdat, "AT38", "Curve"
'Sheets("Calc").Range("shsysCurve").Value = ToggleButton2.Value
End Sub


Sub ChangeAOR()

Call Inp_Data(Fdat, "AJ28", "Curve")
Call Inp_Data(Edat, "AT23", "Curve")
Qr = Sheets("Calc").Range("flow").Value

Call Inp_Data(Cdat, "AORmax", "Calc")
Call Cal_R2dat(Fdat, Rdat, Edat, Cdat, Qr)
Copy_AnyDat Rdat, "AZ5", "Curve"

End Sub

Sub ChangeRatedDia()
Dim RtDia#
'If Not Me.RatedDia.Value = "" Then: RtDia = Me.RatedDia.Value
'Cal_Fdat Mdat, Fdat, RtDia
'Copy_AnyDat Fdat, "AT38", "Curve"
MsgBox "request pending"
'Sheets("Calc").Range("shsysCurve").Value = ToggleButton2.Value
End Sub





Private Sub CmdSave_Click()
Call Module1.TrExportD("History", "A1:VV300", ".csv")
End Sub

Sub UserForm_Initialize()
Call Module1.TrLoadDefaults
Call Module1.TrUserModify(Me, False)
End Sub
Private Sub CommandButton1_Click()
Call Module1.TrLoadDefaults
Call Module1.TrUserModify(Me, False)
End Sub

Private Sub CommandButton2_Click()
arr = Split("POR1,POR2,MCSF", ",")
For i = 0 To 2
'Call ShowDataLabels("Curve", "Head", arr(i), 2)
Next i
End Sub

Private Sub CommandButton3_Click()
Call DrawSysCurve
End Sub

Private Sub CommandButton4_Click()
Call ChangeRatedDia
End Sub

Private Sub CommandButton5_Click()
Call RealTheory
End Sub

Private Sub CommandButton6_Click()
Call ChangeVisc
End Sub

Private Sub CommandButton7_Click()
Call PlotUnit
End Sub

Private Sub CommandClose_Click()
 Unload Me
End Sub

Private Sub EnterData_Click()
Call Module1.TrUserModify(Me, True)
Call ChangeAOR
Call CommandClose_Click
End Sub

Private Sub ToggleButton1_Click()
Call AddNPSH
End Sub

Private Sub UserForm_Click()
Call Module1.TrUserModify(Me, False)
End Sub

