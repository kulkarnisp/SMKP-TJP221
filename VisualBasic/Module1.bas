Attribute VB_Name = "Module1"

Dim Series As String
Dim model As String
Dim Captions() As String
Dim Values()
Dim Units()
Public ModelArray() As String

'Captions = Split("CustomerName,ItemNo,Service,JobNo", ",")
'Values = Split("flow,head,speed,model", ",")

Dim Spgr As Long
Dim Vicsosity As Long

' add npsh viscosity correction 12 March
' add frequency option 30 March
' Specific gravity in power calculation
' add diameter margin

Private Sub Auto_Open() ' Define a range named welcome in sheet
   'On Error Resume Next
Call WakeupCall(ModelArray)
Call SProtection("SMKP_Chart", False)
Call SProtection("Selection", True)

End Sub

Sub WakeupCall(ModelArray)
ReDim ModelArray(46)
Dim SHName As String
    SHName = "SMKP_Chart"
    With Sheets(SHName)
    If Err <> 0 Then
        MsgBox "Click: TJPxtc-Cpation 3."
    Else
        .Visible = True
        .Select
'        .ChartObjects("GrRg").Visible = True
'        .Range("C4").Select
        ActiveWindow.Zoom = 100
    End If
    End With
'For i = 1 To 46
'ModelArray(i) = Sheets("Calc").Cells(28 + i, 1).Value
'Next i
'Sheets("Calc").Range("theory").Value = "True"
i = 6
ThisWorkbook.Sheets(i).Visible = True
For i = 1 To 9
If i = 6 Then: i = i + 1
ThisWorkbook.Sheets(i).Visible = xlVeryHidden
Next i
ThisWorkbook.Sheets("SMKP_Chart").Visible = True

Debug.Print "WakeupCall: Models Loaded " '& ModelArray(Rnd(46))



End Sub

Sub GoalSeek(SHName, ScRg)
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

Sub CheckRPM()
Application.ScreenUpdating = True
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
        On Error GoTo CleanFail:
CleanRun:
        
i = 0
iRow = 0
        Set WS = Worksheets(SHName)
        iRow = WS.Cells.Find(What:=SCrit, SearchOrder:=xlRows, _
        SearchDirection:=xlPrevious, LookIn:=xlValues).Row
        'MsgBox "The required value is in " & iRow, , "Search Result"
        TrRowofHistory = iRow
        
        
        Exit Function
CleanFail:
      SCrit = ""
      i = i + 1
     If i = 3 Then: MsgBox "Friday: Server Error Module1TrRoH": End
     Resume CleanRun
    
    
End Function


Sub TrSheettoForm(SHName As String, ScRg As String, Form As Object, inverse As Boolean)
   On Error Resume Next
'   Application.DisplayAlerts = False
    ' Total number of Textboxes
    n = 40
    
     With Sheets(SHName)
    
    m = .Range(ScRg).Row
    c = .Range(ScRg).Column
    
  If inverse Then
    ' Row number of Data Start
          For j = 1 To n
           Form.Controls("TB" & j).Value = .Cells(m + j, c + 1).Value
            Form.Controls("LB" & j).Value = .Cells(m + j, c).Value
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
      Application.DisplayAlerts = True
     
End Sub

Sub TrHistorytoForm(SHName As String, SCText As String, Form As Object, inverse As Boolean)
    On Error Resume Next
    ' Total number of Textboxes
    n = 40
    ' Row number of Data Start

    m = TrRowofHistory(SHName, SCText)

'    m = TrRowofHistory(SHName, "")

    Debug.Print "TJP;module1 " & m

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
        .Cells(m, c).Value = Now()
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
   Box.Controls("TextBox" & r).Value = .Cells(m + 1 + i, n + 1 + j).Value
    Next i
    Next j
    
    r = 67
    For j = 1 To 3
    For i = 1 To 6
    r = r + 1
    Box.Controls("TextBox" & r).Value = .Cells(m + 14 + i, n + 1 + j).Value
    Next i
    Next j
    
    For j = 2 To 8
   Box.Controls("TextBox" & j).Value = .Cells(m + 23, n - 1 + j).Value
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
    .Cells(m + 1 + i, n + 1 + j).Value = Box.Controls("TextBox" & r).Value
    Next i
    Next j
    
    r = 67
    For j = 1 To 3
    For i = 1 To 6
    r = r + 1
    .Cells(m + 14 + i, n + 1 + j).Value = Box.Controls("TextBox" & r).Value
    Next i
    Next j
    
    For j = 2 To 8
    .Cells(m + 23, n - 1 + j).Value = Box.Controls("TextBox" & j).Value
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
    sArray = Split("AORmax,AORmin,PORmax,PORmin,isoEff", ",")
    For j = 0 To UBound(sArray): Range(a & sArray(j)).Value = Range(b & sArray(j)).Value: Next j
    Range("isoAffiniity").Value = True
    
End Sub


Sub TrUserModify(Box As Object, inverse As Boolean)

    sArray = Split("AORmax,AORmin,PORmax,PORmin,isoEff", ",")
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





