VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Dim Series As String
Dim model As String
Dim Captions() As String
Dim Values()
Dim Units()
Dim InArray() As String

'Captions = Split("CustomerName,ItemNo,Service,JobNo", ",")
'InArray = Split("flow,head,speed,model", ",")

Dim Spgr As Long
Dim Vicsosity As Long
Private Sub Auto_Open() ' Define a range named welcome in sheet
    Dim vntToday As Variant
    vntToday = WorksheetFunction.Text(Date, "mmmm dd")
    On Error Resume Next
    Sheets("Input").Select
    If Err <> 0 Then
        MsgBox "Click: Return to Input."
    Else
        Range("Welcome").Select
        ActiveWindow.Zoom = 10
        MsgBox "Welcome User j."
        Range("Welcome").Select
        ActiveWindow.Zoom = 110
    End If

End Sub


Sub SMKPData(override As Boolean, a As String)
    model = Range(a & "model").Value
    With Sheets("Calc")
    If Not (.Cells(11, 8).Value = model And override) Then
        .Cells(7, 34).Value = model
        .Cells(24, 34).Value = model
        .Cells(1, 34).Value = model

        For j = 1 To 7
            For i = 1 To 11: .Cells(4 + i, 11 + j).Value = .Cells(9 + i, 33 + j).Value: Next i
        Next j
        For j = 1 To 3
            For i = 1 To 6: .Cells(17 + i, 11 + j).Value = .Cells(26 + i, 33 + j).Value: Next i
        Next j
        For j = 1 To 4: .Cells(2, 13 + j).Value = .Cells(4, 32 + j).Value: Next j
        
        .Cells(2, 12).Value = .Cells(4, 39).Value 'PUMP ID for SMKP series
        .Cells(2, 13).Value = model
    End If
    End With
    Debug.Print Chr(13) & model & " " & Date & Chr(13)
End Sub

Sub GoalSeek()
    SHName = "Calc"
    ScRg = "cutr"
    ChRg = "error"
     Sheet1.Visible = True
     Range("cutdia").Value = 0
     
        With Sheets(SHName)
            .Range(ScRg).Value = 0.85
           .Range(ChRg).GoalSeek Goal:=0, ChangingCell:=.Range(ScRg)
        End With
        
   If Range("ViscosityCorrection") = True Then
        Range("ViscCorrect") = "Yes"
        ChRg = "errorvisc"
        GoTo GoalSeeker
    Else
        Range("ViscCorrect") = "No"
    End If
    
GoalSeeker:
        With Sheets(SHName)
            .Range(ScRg).Value = 0.85
           .Range(ChRg).GoalSeek Goal:=0, ChangingCell:=.Range(ScRg)
        End With
     Sheet1.Visible = xlVeryHidden
End Sub


Sub Calculate()
Application.ScreenUpdating = False
On Error GoTo ErrHdl
Dim model As String
If Not Sheets("Calc").Range("ZZ4").Value Then Call SMKPData(True, "")
    Call GoalSeek
    Call TrLoadDefaults 'Load defaul baulae
'    Call ScaleActiveChart("Curve", "Graf1")
'    Call ScaleActiveChart("Curve", "Graf2")
'    Call ScaleActiveChart("Graph3", "Graf3")
If Not Sheets("Calc").Range("ZZ4").Value Then Call CheckRPM
exthdl:
Exit Sub
ErrHdl:
MsgBox "Jarvis:1503 Viscosity Corrr Not Required"
GoTo exthdl
Application.ScreenUpdating = False
End Sub



Sub goStoD(d As String, Leave As String)
    On Error Resume Next
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
    With ActiveWindow
        .SplitColumn = 31.2127659574468
        .SplitRow = 35.7142857142857
       .Panes(4).Activate
      .Panes(4).ScrollRow = 112
      .Panes(2).ScrollColumn = 32
     
    End With
    
  
    
End Sub
Sub CheckRPM()
    SHName = "Input"
    sArray = Split("model,speed", ",")
    a = "j"
    b = "i"
 For j = 0 To UBound(sArray): Range(a & sArray(j)).Value = Range(b & sArray(j)).Value: Next j

    nRPM = Range("ispeed").Value
  With Sheets(SHName)
    For i = 15 To 18
        If .Cells(i, 9).Value = nRPM Then
            .Range("K9").Value = Messg
            Exit Sub
        Else
            Messg = "Check RPM in Product Dialogue"
            .Range("K9").Value = Messg
        End If
    Next i
  End With
Mbox:
    MsgBox Messg, vbCritical, "TJP-RPM err"
End Sub



Sub TrExportD(SHName As String, ScRg As String, ftp As String)
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
 
    On Error Resume Next
    
    'SHName = "Input"
    'ScRg = "T2:Z24"
    
    Dim wb, wbI As Workbook
    Dim wsI As Worksheet
    Dim wRg As Range
        
    Set wbI = ThisWorkbook
    Set wsI = wbI.Sheets(SHName)
    Set wRg = Sheets(SHName).Range(ScRg)
    
    
    Set wb = Application.Workbooks.Add
    wRg.Copy
    wb.Worksheets(1).Paste
    saveFile = Application.GetSaveAsFilename(FileFilter:="Enter file Name (*" & ftp & "),*" & ftp)
    wb.SaveAs Filename:=saveFile, FileFormat:=xlNormal, CreateBackup:=False
    wb.Close

    MsgBox "Saved as CommaSepFile " & fName, vbDefaultButton1
    
    Application.CutCopyMode = False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
End Sub

Sub TrImportD(SHName As String, ScRg As String, ftp As String)
    On Error Resume Next
    'SHName = "History2"
    'ScRg = "D23"

    Dim wbI As Workbook, wbO As Workbook
    Dim wsI As Worksheet
    Dim Rg As Range
    Set wbI = ThisWorkbook
    Set wsI = wbI.Sheets(SHName)
    
    Filt = "Select File (*" & ftp & "),*" & ftp
    Title = "Select a boolean/cst File to Import"
    Filename = Application.GetOpenFilename(FileFilter:=Filt, Title:=Title)
    
    If Filename = False Then
    MsgBox "No File Was Selected"
    Exit Sub
    End If
   
   Set wbO = Workbooks.Open(Filename)
    
    With wbO.Sheets(1)
    lRow = .Cells(Rows.count, 2).End(xlUp).Row
    lCol = .Cells(1, Columns.count).End(xlToLeft).Column
    .Range(Cells(1, 1), Cells(lRow, lCol)).Copy wsI.Range(ScRg)
    End With
  wbO.Close SaveChanges:=False
      
End Sub


Function TrRowofHistory(SHName As String, SCrit As String) As Integer
   
   Dim iRow As Integer
        Dim adrs As String
        Dim WS As Worksheet
        On Error GoTo MyErrorHandler:
        Application.ScreenUpdating = False

        Set WS = Worksheets(SHName)
        iRow = WS.Cells.Find(What:=SCrit, SearchOrder:=xlRows, _
        SearchDirection:=xlPrevious, LookIn:=xlValues).Row
        'MsgBox "The required value is in " & iRow, , "Search Result"
        TrRowofHistory = iRow
        
exitHandler:
        Exit Function
MyErrorHandler:
        MsgBox "Friday: Contact Server Error"
        Resume exitHandler

    
    
End Function


Sub TrSheettoForm(SHName As String, ScRg As String, Form As Object, inverse As Boolean)
   On Error Resume Next
    ' Total number of Textboxes
    n = 40
    m = Range(ScRg).Row
    c = Range(ScRg).Column
     With Sheets(SHName)
  If inverse Then
    ' Row number of Data Start
          For j = 1 To n
           Form.Controls("TB" & j).Value = .Cells(m + j, c + 1).Value
           Form.Controls("LB" & j).Caption = .Cells(m + j, c).Value
           Form.Controls("CB" & j).Value = .Cells(m + j, c + 2).Value
        Next j
   Else
  
        For j = 1 To n
          .Cells(m + j, c).Value = Form.Controls("LB" & j).Caption
          .Cells(m + j, c + 1).Value = Form.Controls("TB" & j).Value
          .Cells(m + j, c + 2).Value = Form.Controls("CB" & j).Value
        Next j
  End If
     End With
     
     
End Sub

Sub TrHistorytoForm(SHName As String, SCText As String, Form As Object, inverse As Boolean)
    On Error Resume Next
    ' Total number of Textboxes
    n = 40
    ' Row number of Data Start
    
    m = TrRowofHistory(SHName, SCText)
    c = 1
    
    With Sheets(SHName)
    
        If inverse Then
        
        For j = 1 To n
        Form.Controls("TB" & j).Value = .Cells(m, c + j).Value
        Next j
        
        Else
        For j = 1 To n
        .Cells(m, c + j).Value = Form.Controls("TB" & j).Value
        Next j 'One line is one statement and on error a complete statement is skipped
        
        End If
        
    End With
End Sub


Sub TrCurveData0(SHName As String, ScRg As String, Box As Object)
Application.ScreenUpdating = False
    On Error Resume Next
    n = Range(ScRg).Column - 1
    m = Range(ScRg).Row - 1

    With Sheets(SHName)
 '.Unprotect Password:="password"
    r = 12
    For j = 1 To 6
    For i = 1 To 11
    r = r + 1
   Box.Controls("TextBox" & r).Value = .Cells(m + 4 + i, n + j).Value
    Next i
    Next j
    
    r = 67
    For j = 1 To 3
    For i = 1 To 6
    r = r + 1
    Box.Controls("TextBox" & r).Value = .Cells(m + 17 + i, n + j).Value
    Next i
    Next j
    
    For j = 1 To 8
   Box.Controls("TextBox" & j).Value = .Cells(m + 2, n - 1 + j).Value
    Next j
     ' .Protect Password:="password"
    End With
 Application.ScreenUpdating = True
End Sub
Sub TrCurveData1(SHName As String, ScRg As String, Box As Object)
    On Error Resume Next
    n = Range(ScRg).Column - 1
    m = Range(ScRg).Row - 1

    With Sheets(SHName)
 '.Unprotect Password:="password"
     r = 12
    For j = 1 To 5
    For i = 1 To 11
    r = r + 1
    .Cells(m + 4 + i, n + j).Value = Box.Controls("TextBox" & r).Value
    Next i
    Next j
    
    r = 67
    For j = 1 To 3
    For i = 1 To 6
    r = r + 1
    .Cells(m + 17 + i, n + j).Value = Box.Controls("TextBox" & r).Value
    Next i
    Next j
    
    For j = 1 To 8
    .Cells(m + 2, n - 1 + j).Value = Box.Controls("TextBox" & j).Value
    Next j
      
   ' .Protect Password:="password"
    End With

End Sub

Sub TrUserData(a As String, b As String)
    Dim sArray() As String
    'A = ""
    'B = "i"
    sArray = Split("flow,head,speed,SpGr,power,Viscosity,model,Stages,Series,ViscosityCorrection", ",")
    For i = 0 To UBound(sArray): Range(a & sArray(i)).Value = Range(b & sArray(i)).Value: Next i

End Sub


Sub TrLoadDefaults()
    Dim sArray() As String
    a = ""
    b = "i"
    sArray = Split("AORmax,AORmin,PORmax,PORmin,pullx,effcorrect,cutdia,pullx2", ",")
    For j = 0 To UBound(sArray): Range(a & sArray(j)).Value = Range(b & sArray(j)).Value: Next j

End Sub


Sub TrUserModify(Box As Object, inverse As Boolean)

    sArray = Split("AORmax,AORmin,PORmax,PORmin,pullx,effcorrect,cutdia,pullx2", ",")
    n = UBound(sArray) + 1
      
    If inverse Then
            For j = 1 To n: Range(sArray(j - 1)).Value = Box.Controls("TextBox" & j).Value: Next j
    Else
            For j = 1 To n: Box.Controls("TextBox" & j).Value = Range(sArray(j - 1)).Value: Next j
    End If
        
End Sub



Sub TrDetails()
    ' Go TO Details Temporary *Alternrative
      Sheet7.Unprotect Password:="noriya"
    Dim pf As PivotField
    Set pf = Sheets("Details").PivotTables("configuration").PivotFields("Model")
    'pf.ClearAllFilters
    pf.CurrentPage = Range("model").Value
       
End Sub



Sub InsideSearch(SHName As String, SCrit As String)
'SHName = "History"

Dim iRow As Integer
Dim adrs As String
Dim WS As Worksheet
On Error GoTo MyErrorHandler:
Application.ScreenUpdating = False
'Sheets("History2").Visible = True
Set WS = Worksheets(SHName)
iRow = WS.Cells.Find(What:=SCrit, SearchOrder:=xlRows, _
SearchDirection:=xlPrevious, LookIn:=xlValues).Row + 1
MsgBox "The required value is in " & iRow, , "Search Result"

exitHandler:
Exit Sub
MyErrorHandler:
MsgBox "Friday: Contact Server Error"
Resume exitHandler

End Sub


Sub TrExportD1(Form As Object, TBName As String, txtName As Integer)
    
    On Error Resume Next
    'Total number of Textboxes
    TBX = TBName
    If TBX = "TB" Then LB = "LB" Else LB = "Label"
    If TBX = "TB" Then CB = "LB" Else CB = "ComboBox"
    n = tbCount(TBName, Form)
    k = txtName
    fName = Form.Controls(TBX & k).Value & ".txt"
    fPath = Application.ThisWorkbook.path
    fName = Replace(fName, "/", "-")
    Set fs = CreateObject("Scripting.FileSystemObject")
    FilePath = fs.BuildPath(fPath, fName)
    Set a = fs.CreateTextFile(FilePath, True)
    
    With Form
     For j = 1 To n
      If tbExists(LB & j, Form) Then Label = .Controls(LB & j).Caption Else Label = " "
      If tbExists(TBX & j, Form) Then Value1 = .Controls(TBX & j).Value Else Value1 = " "
      If tbExists(CB & j, Form) Then Unit = .Controls(CB & j).Value Else Unit = " "
    
      fWrite = Label & Chr(9) & Value1 & Chr(9) & Unit
      a.WriteLine (fWrite)
     Next j
    End With
    a.Close
    MsgBox "Saved as " & fName
    
End Sub
Function tbExists(argName As String, Form As Object) As Boolean
 
   Dim ctrl As control 'As TextBox doesn't work '(inspiartion from erik.van.geit
    On Error Resume Next
    With Form
        Set ctrl = .Controls(argName)
'        For Each ctrl In .Controls
'            If ctrl.Name = argName Then
'                tbExists = True
'                Exit For
'                Else
'                tbExists = True
'            End If
'        Next ctrl
    End With
        If Err = 0 Then tbExists = True Else: tbExists = False
        On Error GoTo 0
     
End Function
Function tbCount(argName As String, Form As Object)
Dim cCont As control
Dim lCount As Long

    For Each cCont In Form.Controls
        If TypeName(cCont) = argName Then
            lCount = lCount + 1
        End If
     Next cCont
        tbCount = lCount
End Function



