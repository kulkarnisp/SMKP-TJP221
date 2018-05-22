Attribute VB_Name = "TJPgrpro"
Sub SMKP_Plot()
Attribute SMKP_Plot.VB_ProcData.VB_Invoke_Func = " \n14"
SHName = "Calc"
On Error Resume Next

PlotChart SHName, "Head"
PlotSeries SHName, "Head", "Headr", "AL28:AL38", "AK28:AK38"
PlotSeries SHName, "Head", "Headmx", "AL3:AL13", "AK3:AK13"
PlotSeries SHName, "Head", "Headmn", "AL15:AL25", "AK15:AK25"
PlotSeries SHName, "Head", "Ratedpt", "Head", "Capacity"
ChartAxes SHName, "Head", "Flow (m3/hr)", "Head (m)"
PlotSeries SHName, "Head", "Isor", "R54:R154", "Q54:Q154"

PlotChart SHName, "Power"
PlotSeries SHName, "Power", "Headr", "AM28:AM38", "AK28:AK38"
PlotSeries SHName, "Power", "Headmx", "AM3:AM13", "AK3:AK13", 252
PlotSeries SHName, "Power", "Ratedpt", "Power", "Capacity"
ChartAxes SHName, "Power", "Flow (m3/hr)", "Power (KW)"

PlotChart SHName, "NPSH"
PlotSeries SHName, "NPSH", "Headr", "AQ28:AQ38", "AK28:AK38"
PlotSeries SHName, "NPSH", "Headmx", "AQ3:AQ13", "AK3:AK13", 252
PlotSeries SHName, "NPSH", "Ratedpt", "NPSHr", "Capacity"
ChartAxes SHName, "NPSH", "Flow (m3/hr)", "NPSH (m)"

PlotChart SHName, "Effi"
PlotSeries SHName, "Effi", "Headr", "A028:A038", "AK28:AK38"
PlotSeries SHName, "Effi", "Headmx", "A03:A013", "AK3:AK13", 252
PlotSeries SHName, "Effi", "Ratedpt", "BEP", "BEPQ"
ChartAxes SHName, "Effi", "Flow (m3/hr)", "Efficiency (%)"

PlotSeries SHName, "Head", "ViscH", "AU3:AU13", "AT3:AT13"
PlotSeries SHName, "Effi", "ViscE", "AX3:AX13", "AT3:AT13"
PlotSeries SHName, "Power", "ViscP", "AV3:AV13", "AT3:AT13"
PlotSeries SHName, "Head", "ViscMax", "AD3:AD13", "AC3:AC13"
PlotSeries SHName, "Head", "ViscMin", "AL3:AL13", "AK3:AK13"

End Sub

Sub PlotSeries(SHName, GRName, SCName, YcRg, XcRg, Optional clr = 0)
 On Error Resume Next
 
'     SHame = "Graph3"
'    GRName = "graf1"
'    SCName = "ViscHead"
'    YcRg = "=plotTable[Hvisc]"
'    XcRg = "=plotTable[Qvisc]"
    
    Sheets(SHName).Activate
    ActiveSheet.ChartObjects(GRName).Activate
    t = ActiveChart.SeriesCollection.count
    ActiveChart.SeriesCollection.NewSeries
    t = t + 1
     With ActiveChart.SeriesCollection.Item(t)
        .Name = SCName
        .XValues = Range(XcRg)
        .Values = Range(YcRg)
        .Border.Color = RGB(clr, clr, clr)
        .AxisGroup = 1
        .Format.Line.DashStyle = msoLineSysDash
        .Format.Line.Weight = 1
    End With
'   ActiveChart.FullSeriesCollection(7).Trendlines(1).Select
'    With Selection
'        .Type = xlPolynomial
'        .Order = 2
'    End With

End Sub
Sub ChartAxes(SHName, GRName, XName, YName)

'
    Sheets(SHName).ChartObjects(GRName).Activate
    With ActiveChart
    
    .ChartArea.Select
     .Axes(xlValue).Select
    .SetElement (msoElementPrimaryValueAxisTitleAdjacentToAxis)
   .Axes(xlValue, xlPrimary).AxisTitle.Text = YName '"Head (m)"
       
    .Axes(xlCategory).Select
    .SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
     .Axes(xlCategory, xlPrimary).AxisTitle.Text = XName '"Flow (m3/hr)"
      
'    .ChartTitle.Select
'    .ChartTitle.Text = "Performance Curve"
    
    
'    .ChartArea.Select
'    .ChartTitle.Select
'    Selection.Delete

    .Axes(xlCategory).Select
    Selection.TickLabels.NumberFormat = "#,##0.00"
    Selection.TickLabels.NumberFormat = "#,##0.0"
    .Axes(xlValue).Select
    Selection.TickLabels.NumberFormat = "#,##0.00"
    Selection.TickLabels.NumberFormat = "#,##0.0"
    .ChartArea.Select
    
    End With
    


End Sub

Sub ScaleChart(SHName As String, GRName As String)
    'Sheets(SHName).Unprotect Password:="noriya"
    On Error Resume Next
    Dim Max, Min As Double
    With Sheet1
        With Sheet1.ListObjects("plot_var")
        a = .ListColumns(8).DataBodyRange.Value
        b = .ListColumns(9).DataBodyRange.Value
        End With
        .Cells(45, 11).Value = a
        .Cells(46, 11).Value = b
        Max = .Cells(45, 11).Value
        Min = .Cells(46, 11).Value
    End With

    Sheets(SHName).Activate
    ActiveSheet.ChartObjects.Item(1).Activate
    ActiveSheet.ChartObjects(GRName).Activate
    With ActiveChart.Axes(xlCategory)
     .MinimumScale = Min
     .MaximumScale = Max
    End With
        ActiveChart.Axes(xlValue, xlSecondary).MinimumScale = 0
        ActiveChart.Axes(xlValue, xlSecondary).MaximumScale = 100
        ActiveChart.Axes(xlValue).MinimumScale = 0
        ActiveChart.Axes(xlValue).MaximumScaleIsAuto = True
    With ActiveSheet
        .Range("A2").Select
        .Range("A2").Value = "Scaled " & Now()
    End With

End Sub
Sub DeleteSeries(SHName, GRName, SCName)
 On Error Resume Next
'
'    SHName = "Curve"
'    GRName = "Graf1"
'    SCName = "ViscEff"
    
    Sheets(SHName).Activate
    ActiveSheet.ChartObjects(GRName).Activate
    With ActiveChart
    DoEvents
    For c = .SeriesCollection.count To 1 Step -1
        If .SeriesCollection(c).Name Like SCName Then .SeriesCollection(c).Delete
    Next c
    End With

End Sub

Sub ShowDataLabels(SHName, GRName, SCNameR, Optional k = 2)
On Error Resume Next
 '   k = 2
'    SHName = "Graph3" '"Curve"
'    GRName = "graf1"
'    SCNameR = Split("POR1,POR2,MCSF", ",")
    SCNameR = Split(SCNameR, ",")
For Each SCName In SCNameR
Dim lst As Variant
Dim WS As Worksheet
Dim cht As Chart
Dim srs As Series
Dim pt As Point
Dim P As Integer
Set WS = Worksheets(SHName)
    Sheets(SHName).Activate
    ActiveSheet.ChartObjects(GRName).Activate
Set cht = ActiveChart
    
    With cht
    DoEvents
    For c = .SeriesCollection.count To 1 Step -1
        If .SeriesCollection(c).Name Like SCName Then
            Set srs = cht.SeriesCollection(c)
            GoTo fasterfene
         End If
    Next c
fasterfene:
    End With
      srs.ApplyDataLabels
            With srs
            .DataLabels.SeriesName = True
            .HasDataLabels = True
            .DataLabels.Position = xlLabelPositionAbove 'Left,Right
             .DataLabels.NumberFormat = "##.#"
              .DataLabels.ShowCategoryName = True
              .DataLabels.ShowValue = False
              .DataLabels.ShowSeriesName = False
             .DataLabels.Orientation = xlUpwards
                    With .DataLabels.Format
                        .TextFrame2.TextRange.Font.Size = 10
                        .TextFrame2 = True
                        .TextFrame2.Orientation = msoTextOrientationUpward
                        .Line.Visible = msoTrue
                        .Line.Weight = 0.75
                        .Fill.Visible = msoTrue
                        .Fill.Solid
                        .Fill.ForeColor.ObjectThemeColor = msoThemeColorLight1
                     End With
            End With
    
    
    
    For P = 0 To srs.DataLabels.count
        If Not P = k Then
        Set pt = srs.Points(P)
        '## remove the datalabel for this point
        pt.DataLabel.Delete
        End If
      Next
    
Next SCName
End Sub
Sub DeleteChart(SHName, GRName)

With ThisWorkbook
    On Error GoTo errHandler
    .Sheets(SHName).ChartObjects(GRName).Activate
    ActiveChart.Parent.Delete
End With

exitHandler:
    Exit Sub
errHandler:
    MsgBox "Everything Deleted!", vbExclamation, "Infinity Stone"
    Resume exitHandler
End Sub
Sub DelinkDataFromChart(SHName, GRName)

    ''' Thanks to Tushar Mehta
    
    Dim mySeries As Series
 
    Sheets(SHName).ChartObjects(GRName).Activate
    ''' Make sure a chart is selected
    If Not ActiveChart Is Nothing Then
        ''' Loop through all series in active chart
        For Each mySeries In ActiveChart.SeriesCollection
            '''' Convert X and Y Values to arrays of values
            mySeries.XValues = mySeries.XValues
            mySeries.Values = mySeries.Values
            mySeries.Name = mySeries.Name
        Next mySeries

        n = a
        ActiveChart.Shapes("ModelName").Select
        ActiveChart.Shapes("ModelName").TextFrame.Characters.Text = Range("model").Value & " Group-" & Range("group").Value

    End If

End Sub
Sub PlotChart(SHName, GRName)
 
 'On Error Resume Next
With ThisWorkbook.Sheets(SHName)
n = .ChartObjects.count
.Shapes.AddChart2(240, xlXYScatterSmoothNoMarkers).Select
.ChartObjects(n + 1).Select
Selection.Name = GRName
End With

End Sub
Sub CopyCurve(SHName As String, CMName As String)

    Sheets(SHName).ChartObjects.Item(2).Activate
    ActiveChart.ChartArea.Select
    ActiveChart.ChartArea.Copy
    Sheets(CMName).Select
    Range("B11").Select
    ActiveSheet.Paste
    
    'Special(DataType:=ppPasteShape, Link:=False)
    '   n = Sheets(CMName).ChartObjects.Count
    ' Sheets(CMName).Charts(n).BreakLink
    
End Sub

Sub ShowChart(t As Boolean)
Sheets("Input").ChartObjects.Item(1).Select
If t Then
Set RngToCover = ActiveSheet.Range("G5:P46")
    With Selection
         .Height = RngToCover.Height ' resize
         .Width = RngToCover.Width   ' resize
         .Top = RngToCover.Top       ' reposition
         .Left = RngToCover.Left
     End With
End If

Sheets("Input").ChartObjects.Item(1).Visible = t
Range("iSeries").Select

End Sub

Sub connect_all(SHName As String)
On Error Resume Next
Dim sh As Shape
For Each sh In Sheets(SHName).Shapes
  sh.OnAction = "picture_click"
Next sh
End Sub
Sub picture_click(SHName)

  Sheets(SHName).Shapes(Application.Caller).Select
ActiveSheet.Shapes("GraphC").ZOrder msoBringToFront
End Sub
