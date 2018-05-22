Attribute VB_Name = "TJPintact"
Sub SMKP_Data(Optional SModel = 0)

 LocalSource SModel
   
 
'    Accdb_Dat "modelname", "A79", "A1:H26"
'    WhiteWash "A78", "A1:H26"
'   CopyData "A78", "A1", "A1:H26"

    
End Sub
Sub Inp_Data(Cdat, ScRg, Sh2Name)
'On Error GoTo errhdl
Dim i%, j%, k%, rn%, cn%
'SCRg = "A2"
Dim Xd%, Yd%
With ThisWorkbook.Sheets(Sh2Name)


rn = .Range(ScRg).Row
cn = .Range(ScRg).Column
i = rn
j = cn

Do While Not IsEmpty(.Cells(i, cn + 1).Value): i = i + 1: Loop
Xd = i - .Range(ScRg).Row
Do While Not IsEmpty(.Cells(rn, j).Value): j = j + 1: Loop
Yd = j - .Range(ScRg).Column


    ReDim Cdat(0, Yd, Xd)
    
    Cdat(0, 0, 0) = .Cells(rn, cn)

 
    For j = 0 To Yd
        For i = 1 To Xd
        Cdat(0, j, i) = .Cells(rn + i - 1, cn + j).Value
        Next i
    Next j
End With
    
Exit Sub
ErrHdl:
MsgBox "Error code 1506: Check Input Data"
End
    
End Sub

Sub Inp_UDat(Qr, Hr, Nr, Nd, Visc, ViscReq As Boolean, Optional r, Optional Spgr)

'Nd = 5

With ThisWorkbook.Sheets("Calc")
    array1 = Split("flow,head,speed", ",")
    array2 = Split("funit,hunit,sunit", ",")
    
    Qr = .Range("flow").Value
    fu = .Range("funit").Value
    Chk_unt "flow", fu, Qr
    .Range("flow").Value = Qr
    .Range("funit").Value = fu
        
    Hr = .Range("head").Value
    fu = .Range("hunit").Value
     Chk_unt "head", fu, Hr
    .Range("head").Value = Hr
    .Range("hunit").Value = fu
    
    
    Nr = .Range("speed").Value
    fu = .Range("sunit").Value
     Chk_unt "speed", fu, Nr
    .Range("speed").Value = Nr
    .Range("sunit").Value = fu 'S-speed
    
    Checkzero Nr
    
'   Qmx = .Range("pullx").Value
    
    Spgr = .Range("SpGr").Value
    Visc = .Range("Viscosity").Value
    ViscReq = .Range("ViscosityCorrection")
      
'    if .
    r = .Range("Qindex").Value
     .Range("isoAffiniity").Value = True
     
End With

End Sub
Sub Checkzero(qt)

If qt = 0 Then
MsgBox "Zero input parameters"
End
End If

End Sub

Sub Inp_TempSpeedList(Nlist)
ReDim Preserve Nlist(4) As Integer

With Sheets("Calc")
If .Range("Hz").Value = 60 Then
m = 1
n = 3
Else
m = 2
n = 4
End If

For i = m To n Step 2
Nlist(i) = .Cells(40 + i, 10).Value
Next i
End With

End Sub
 
Sub Copy_AnyDat(Adat, ScRg, SHName)
On Error GoTo ErrHdl
Dim i%, cn%, rn%, j%, t%
i = 1
  
 With ThisWorkbook.Sheets(SHName)
cn = .Range(ScRg).Column
rn = .Range(ScRg).Row

For k = 0 To UBound(Adat, 1)
   t = 0
    For i = i To i + UBound(Adat, 3)
    For j = 0 To UBound(Adat, 2)
        If Not Adat(k, j, t) = 0 Then: .Cells(rn + i, cn + j).Value = Adat(k, j, t)
    Next j
    t = t + 1
    Next i
    i = i + 1 'For VPF
Next k
End With
Exit Sub

ErrHdl:
MsgBox "TJPintact: cannot input data to excel sheet"
End

End Sub

Sub VPF_data(model)
'On Error GoTo errhdl
With ThisWorkbook.Sheets("Calc")
           
        arr = Split("a,b,", ",")
        For i = 0 To UBound(arr) - 1
        .Range(arr(i) & "VPF").Value = model
        Next i

        CopyData "J112", "A1", "A1:H26"
End With

Exit Sub
ErrHdl:
Debug.Print "Missed; Model" & model

End Sub

Sub LocalSource(modelid)
'On Error GoTo ErrHdl:
With ThisWorkbook.Sheets("Calc")
If modelid <> 0 Then: .Range("model").Value = modelid
   If Not SMKPChecklist(.Range("model").Value) Then GoTo ErrHdl
   If Not (.Range("model").Value = .Range("B23").Value) Then
    
        arr = Split("a,b,", ",")
        For i = 0 To UBound(arr) - 1
        .Range(arr(i) & "Model").Value = .Range("model").Value
        Next i
                        
        CopyData "A112", "A1", "A1:H26"
    Else
        Debug.Print "TJPintact:skipped " & .Range("model").Value
    End If

End With

Exit Sub
ErrHdl:
MsgBox "Error in Model Selected: Reselect Model"
End


End Sub

    
Sub Accdb_Dat(ScCrit, destiny, SccRg)
'
' ImportCurveData Macro ''modelname As String
' For importing from Access file saved as Database3

On Error GoTo ErrHdl
Dim Location1, Location As String
Dim address1, address2, address3 As String
Dim orig As Range


Application.ScreenUpdating = False

With ThisWorkbook.Sheets("Calc")

    
Set orig = .Range(destiny)

RcRG = orig.Offset(13, 1).Address
ScRg = orig.Offset(, 1).Address
TcRg = orig.Offset(21, 0).Address

     orig.Offset(-1).Range(SccRg).Clear
'    ActiveSheet.ChartObjects.Delete
    cn = .Range(ScRg).Column
    rn = .Range(ScRg).Row
    .Cells(rn + 1, cn - 1).Value = 1
    cn = .Range(RcRG).Column
    rn = .Range(RcRG).Row
    .Cells(rn + 1, cn - 1).Value = 1

    'modelname = InputBox("Enter Model Number")
    modelname = .Range("model").Value
    
    address1 = "WHERE (`Curve DA`.Model Like '" & modelname & "')"
    'MsgBox modelname
    
    'Location = "\\Sv09c04\710_êÖóÕã@äBê›åvïî\710_ã§í \7109_ÇªÇÃëº\SMKP\Database3.accdb"
    Location1 = Application.ThisWorkbook.path
    Location = Location1 & "\Database3.accdb"
    
 '======================================================================================================================
    
 With .ListObjects.Add(SourceType:=0, Source:=Array(Array( _
        "ODBC;DSN=MS Access Database;DBQ=" & Location & ";DefaultDir=C:\Data;DriverId=25;FIL=MS Access;MaxBufferSize=2048;PageTimeout=" _
        ), Array("5;")), Destination:=.Range(ScRg)).QueryTable
        '.CommandType = 0
        .CommandText = Array( _
        "SELECT  `Curve DA`.Flow, `Curve DA`.Head, `Curve DA`.Power, `Curve DA`.RPM, `Curve DA`.Efficiency, `Curve DA`.Correction" & Chr(13) & "" & Chr(10) & "FROM `" & Location & "`.`Curve D" _
        , _
        "A` `Curve DA`" & Chr(13) & "" & Chr(10) & address1 & Chr(13) & "" & Chr(10) & "ORDER BY `Curve DA`.Flow" _
        )
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = False
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = "TableA"
        .Refresh BackgroundQuery:=False
    End With



  '  ===========================================================================================================================

    address2 = "(Pumps.Model Like '" & modelname & "'))"
'

    With .ListObjects.Add(SourceType:=0, Source:=Array(Array( _
        "ODBC;DSN=MS Access Database;DBQ=" & Location & ";DefaultDir=C:\Data;DriverId=25;F" _
        ), Array("IL=MS Access;MaxBufferSize=2048;PageTimeout=5;")), Destination:= _
        .Range(TcRg)).QueryTable
        '.CommandType = 0
        .CommandText = Array( _
        "SELECT `Outline Dimensions`.`Pump ID`, Pumps.Model, Pumps.Size, `Outline Dimensions`.RatedDia, `Outline Dimensions`.`Imp Max`, `Outline Dimensions`.`Imp Min`, Pumps.`Group`" & Chr(13) & "" & Chr(10) & "FROM " _
        , _
        "`" & Location & "`.`Outline Dimensions` `Outline Dimensions`, `" & Location & "`.Pumps Pumps" & Chr(13) & "" & Chr(10) & "WHERE `Outline Dimensions`.`Pump ID` = Pumps.`Pump ID` AND (" _
        , address2 & Chr(13) & "" & Chr(10) & "ORDER BY `Outline Dimensions`.`Pump ID`")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = False
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = "TableC"
        .Refresh BackgroundQuery:=False
    End With
    
      .Columns("A:M").Select
    Selection.ColumnWidth = 9

  '  ===================================================================================================================
        address3 = "WHERE (`NPSH DA`.Model Like '" & modelname & "')"
    
     With .ListObjects.Add(SourceType:=0, Source:=Array(Array( _
        "ODBC;DSN=MS Access Database;DBQ=" & Location & ";DefaultDir=C:\Data;DriverId=25;FIL=MS Access;MaxBufferSize=2048;PageTimeout=" _
        ), Array("5;")), Destination:=.Range(RcRG)).QueryTable
        '.CommandType = 0
        .CommandText = Array( _
        "SELECT `NPSH DA`.Flow, `NPSH DA`.NPSH, `NPSH DA`.RPM" & Chr(13) & "" & Chr(10) & "FROM `" & Location & "`.`NPSH DA` `NPSH DA`" & Chr(13) & "" & Chr(10) & address3 & Chr(13) & "" & Chr(10) & "ORDER BY `NPSH DA`.Flow" _
        )
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = False
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = "TableB"
        .Refresh BackgroundQuery:=False
    End With
    
'  =========================================================================================================================
    arr = Split("TableA,TableB,TableC", ",")
    For i = 0 To 2
    Name = arr(i)
    .Range(Name & "[#Headers]").Select
    Selection.AutoFilter
     .ListObjects(Name).TableStyle = "TableStyleLight1"
   Next i
   
    .Columns("A:O").Select
    Selection.ColumnWidth = 9

 
End With

Application.ScreenUpdating = True

Exit Sub
ErrHdl:
MsgBox "Friday:Server Contact Error"
        
End Sub
Sub WhiteWash(destiny, ScRg)
'
'Makecolor and range white
'
'ActiveWindow.ScrollColumn = 16
'' ActiveWindow.SmallScroll ToRight:=18
With ThisWorkbook.Sheets(1)
Range(destiny).Offset().Range(ScRg).Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.49
        
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End With
    
End Sub

Sub CreatePDF(Name, Optional ScRg)
    Application.ScreenUpdating = False
    Dim wsA As Worksheet
    Dim wbA As Workbook
    Dim strTime As String
    Dim strName As String
    Dim strPath As String
    Dim strFile As String
    Dim strPathFile As String
    Dim myFile As Variant
    On Error Resume Next 'GoTo errHandler
    Dim Output As String
    
    Output = Name
    
    Sheets(Output).Visible = True
    
    Sheets(Output).Activate
    
    
    Set wbA = ActiveWorkbook
    Set wsA = Sheets(Output)
    strTime = Format(Now(), "yyyy\-mm\-dd\-hhmm\-")
    
    'get active workbook folder, if saved
    strPath = wbA.path
    If strPath = "" Then
      strPath = Application.DefaultFilePath
    End If
    strPath = strPath & "\"
    Item = Sheets("Calc").Range("noitem").Value
    
    'replace spaces and periods in sheet name
    strName = Replace(wsA.Name, " ", "")
    strName = strName & Item
    strName = Replace(strName, ".", "_")
    strName = Replace(strName, "/", "-")
    strName = Replace(strName, " ", "_")
    
    'create default name for savng file
    strFile = strName & ".pdf"   'strTime & service  ##Rev 02
    strPathFile = strPath & strFile
    
    
    'use can enter name and
    ' select folder for file
    myFile = Application.GetSaveAsFilename _
        (InitialFileName:=strPathFile, _
            FileFilter:="PDF Files (*.pdf), *.pdf", _
            Title:="Select Folder and FileName to save")
    
    'export to PDF if a folder was selected
    If myFile <> "False" Then
        wsA.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            Filename:=myFile, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=True
        'confirmation message with file info
      '  MsgBox "PDF file has been created: " _
          & vbCrLf _
          & myFiled
    
       With wsA
        .Range("A1").Select
        .Range("A1").Value = "Printed " & Now()
         End With
    
    
    End If
    
    'Sheets(Output).Visible = False
    'Sheets("Input").Select
    
exitHandler:
        Exit Sub
errHandler:
        MsgBox "Could not create PDF file"
        Resume exitHandler
        
    Application.ScreenUpdating = True

End Sub
 

Sub CopyData(UpRow, destiny, Content)
Application.ScreenUpdating = False
   With ThisWorkbook.Sheets("Calc")

    .Range(UpRow).Offset(1).Range(Content).Copy
 
    .Range(destiny).Offset().Range("A1").PasteSpecial Paste _
     :=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
            .Range("J1").Value = ""
      .Range("J1").Copy
   .Range("J1").Value = Now()

       
     End With
        
Application.ScreenUpdating = True
End Sub


Sub TrialD(SHName As String)
modelname = "50C0010"
Call InputCurveData(modelname)
'
With Sheets
   Range("B1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    
    Dim WS As Worksheet
    Set WS = ThisWorkbook.Sheets.Add(After:= _
             ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    WS.Name = "Tempo"
End With
End Sub

Sub Run_Access_Macro(Name As String)
    Location1 = Application.ThisWorkbook.path

    Location = Location1 & "\Database3.accdb"
    Complete = "MSACCESS.EXE " & Location
    MsgBox Location
         'Opens Microsoft Access and the file nwind.mdb
            Shell (Complete)
         'Initiates a DDE channel to Microsoft Access
            Chan = DDEInitiate("MSACCESS", "system")
         'Activates Microsoft Access
             Application.ActivateMicrosoftApp xlMicrosoftAccess
         'Runs the macro "Sample AutoExec" from the NWIND.MDB file
             Application.DDEExecute Chan, "Export 2"
         'Terminates the DDE channel
          Application.DDETerminate Chan

      End Sub
  Sub RunMacro(Name As String)
    Location1 = Application.ThisWorkbook.path
    Location = Location1 & "\TJPData.accde"
    
          Dim a As Object
    Set a = CreateObject("Access.Application")
    '//Use next statment only needed if you want to see Access - default is to be "invisible"
    a.Visible = False
    a.OpenCurrentDatabase (Location)
    a.DoCmd.RunMacro Name
    a.CloseCurrentDatabase
    a.Quit
    Set a = Nothing


  End Sub

Sub OpenFolder(strDirectory As String)
    'DESCRIPTION: Open folder if not already open. Otherwise, activate the already opened window
    'DEVELOPER: Ryan Wells (wellsr.com)
    'INPUT: Pass the procedure a string representing the directory you want to open
        Dim Dir As String
        Location1 = Application.ThisWorkbook.path
        Dir = Location1 & strDirectory & Range("model").Value
    Dim pID As Variant
    Dim sh As Variant
    On Error GoTo 102:
    Set sh = CreateObject("shell.application")
    For Each w In sh.Windows
        If w.Name = "Windows Explorer" Or w.Name = "File Explorer" Then
            If w.Document.folder.self.path = Dir Then
                'if already open, bring it front
                w.Visible = False
                w.Visible = True
                Exit Sub
            End If
        End If
    Next
    'if you get here, the folder isn't open so open it
    pID = Shell("explorer.exe " & Dir, vbNormalFocus)
102:
End Sub


' Excel macro to export all VBA source code in this project to text files for proper source control versioning
' Requires enabling the Excel setting in Options/Trust Center/Trust Center Settings/Macro Settings/Trust access to the VBA project object model
Public Sub ExportVisualBasicCode()
    Const Module = 1
    Const ClassModule = 2
    Const Form = 3
    Const Document = 100
    Const Padding = 24
    
    Dim VBComponent As Object
    Dim count As Integer
    Dim path As String
    Dim directory As String
    Dim extension As String
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    directory = ActiveWorkbook.path & "\VisualBasic"
    count = 0
 
    
    
    If Not FSO.FolderExists(directory) Then
        Call FSO.CreateFolder(directory)
    End If
    Set FSO = Nothing
    
       'OR
    'ALternative approach
'     If Dir(directory, vbDirectory) = "" Then
'      MkDir directory
'    End If
    
    
    For Each VBComponent In ActiveWorkbook.VBProject.VBComponents
        Select Case VBComponent.Type
            Case ClassModule, Document
                extension = ".cls"
            Case Form
                extension = ".frm"
            Case Module
                extension = ".bas"
            Case Else
                extension = ".txt"
        End Select
            
                
        On Error Resume Next
        Err.Clear
        
        path = directory & "\" & VBComponent.Name & extension
        Call VBComponent.Export(path)
        
        If Err.Number <> 0 Then
            Call MsgBox("Failed to export " & VBComponent.Name & " to " & path, vbCritical)
        Else
            count = count + 1
            Debug.Print "Exported " & Left$(VBComponent.Name & ":" & Space(Padding), Padding) & path
        End If

        On Error GoTo 0
    Next
    
    Application.StatusBar = "Successfully exported " & CStr(count) & " VBA files to " & directory
    Application.OnTime Now + TimeSerial(0, 0, 10), "ClearStatusBar"
End Sub


Public Sub ImportModules()
    Dim wkbTarget As Excel.Workbook
    Dim objFSO As Scripting.FileSystemObject
    Dim objFile As Scripting.File
    Dim szTargetWorkbook As String
    Dim szImportPath As String
    Dim szFileName As String
    Dim cmpComponents As VBIDE.VBComponents

    If ActiveWorkbook.Name = ThisWorkbook.Name Then
        MsgBox "Select another destination workbook" & _
        "Not possible to import in this workbook "
        Exit Sub
    End If

    'Get the path to the folder with modules
    If FolderWithVBAProjectFiles = "Error" Then
        MsgBox "Import Folder not exist"
        Exit Sub
    End If

    ''' NOTE: This workbook must be open in Excel.
    szTargetWorkbook = ActiveWorkbook.Name
    Set wkbTarget = Application.Workbooks(szTargetWorkbook)
    
    If wkbTarget.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to Import the code"
    Exit Sub
    End If

    ''' NOTE: Path where the code modules are located.
    szImportPath = FolderWithVBAProjectFiles & "\"
        
    Set objFSO = New Scripting.FileSystemObject
    If objFSO.GetFolder(szImportPath).Files.count = 0 Then
       MsgBox "There are no files to import"
       Exit Sub
    End If

    'Delete all modules/Userforms from the ActiveWorkbook
    Call DeleteVBAModulesAndUserForms

    Set cmpComponents = wkbTarget.VBProject.VBComponents
    
    ''' Import all the code modules in the specified path
    ''' to the ActiveWorkbook.
    For Each objFile In objFSO.GetFolder(szImportPath).Files
    
        If (objFSO.GetExtensionName(objFile.Name) = "cls") Or _
            (objFSO.GetExtensionName(objFile.Name) = "frm") Or _
            (objFSO.GetExtensionName(objFile.Name) = "bas") Then
            cmpComponents.Import objFile.path
        End If
        
    Next objFile
    
    MsgBox "Import is ready"
End Sub

Function FolderWithVBAProjectFiles() As String
    Dim WshShell As Object
    Dim FSO As Object
    Dim SpecialPath As String

    Set WshShell = CreateObject("WScript.Shell")
    Set FSO = CreateObject("scripting.filesystemobject")

    SpecialPath = WshShell.SpecialFolders("MyDocuments")

    If Right(SpecialPath, 1) <> "\" Then
        SpecialPath = SpecialPath & "\"
    End If
    
    If FSO.FolderExists(SpecialPath & "VBAProjectFiles") = False Then
        On Error Resume Next
        MkDir SpecialPath & "VBAProjectFiles"
        On Error GoTo 0
    End If
    
    If FSO.FolderExists(SpecialPath & "VBAProjectFiles") = True Then
        FolderWithVBAProjectFiles = SpecialPath & "VBAProjectFiles"
    Else
        FolderWithVBAProjectFiles = "Error"
    End If
    
End Function

Function DeleteVBAModulesAndUserForms()
        Dim VBProj As VBIDE.VBProject
        Dim VBComp As VBIDE.VBComponent
        
        Set VBProj = ActiveWorkbook.VBProject
        
        For Each VBComp In VBProj.VBComponents
            If VBComp.Type = vbext_ct_Document Then
                'Thisworkbook or worksheet module
                'We do nothing
            Else
                VBProj.VBComponents.Remove VBComp
            End If
        Next VBComp
End Function

