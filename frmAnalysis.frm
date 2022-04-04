
Private Sub cmdBack3_Click()

Me.Hide
frmSample.Hide
frmSample.Show

End Sub


Private Sub cmdChart_Click()

''declare worksheet 1 and delete worksheet 2 and 3
Dim nWS As Worksheet
Set nWS = eWBook.Worksheets(1)

eWBook.Sheets("Sheet2").Delete
eWBook.Sheets("Sheet3").Delete

Excel.Application.ScreenUpdating = False


''set row number
Dim LastRow As Long
LastRow = nWS.Range("A" & nWS.Rows.count).End(xlUp).Row


'''standard deviation, variance and correlation coefficient chart'''

''add new chart
Dim eChart As Excel.Chart
Set eChart = eWBook.Charts.Add


''set chart properties

    With eChart
        .ChartType = xlLineMarkers
        
        .Location xlLocationAsNewSheet
        .HasTitle = True
        .ChartTitle.Characters.Text = "Measure of Dispersion from Interpolation for all Raingauges"
        .Name = "Dispersion Graph"
    
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Rainfall Station ID"
        
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Dispersion Measures"
        .Axes(xlValue, xlPrimary).MinimumScaleIsAuto = True
        .Axes(xlValue, xlPrimary).MaximumScaleIsAuto = True
                        
        .Axes(xlCategory).HasMajorGridlines = False
        .Axes(xlCategory).HasMinorGridlines = False
        .Axes(xlValue).HasMajorGridlines = True
        .Axes(xlValue).HasMinorGridlines = False
        
        '.Axes(xlCategory).CategoryType = xlCategoryScale
        '.Axes(xlCategory).TickLabelSpacing = 1
        .Axes(xlCategory).TickLabels.NumberFormat = "General"
        .Axes(xlCategory).TickLabels.NumberFormatLinked = True
        
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
            
        
        '' Remove any series created with the chart
            Do Until .SeriesCollection.count = 0
                .SeriesCollection(1).Delete
            Loop

                
        '' Add each series
        Dim srs1 As Series, srs2 As Series, srs3 As Series
        Dim srs4 As Series, srs5 As Series
        
        Set srs1 = .SeriesCollection.NewSeries
                With srs1
                    .Name = nWS.Cells(1, "D")
                    .Values = nWS.Range("D2", "D" & LastRow)
                    .XValues = nWS.Range("A2", "A" & LastRow)
                    .HasDataLabels = True
                End With
       
        Set srs2 = .SeriesCollection.NewSeries
                With srs2
                    .Name = nWS.Cells(1, "E")
                    .Values = nWS.Range("E2", "E" & LastRow)
                    .XValues = nWS.Range("A2", "A" & LastRow)
                    .HasDataLabels = True
                End With
    
        Set srs3 = .SeriesCollection.NewSeries
                With srs3
                    .Name = nWS.Cells(1, "F")
                    .Values = nWS.Range("F2", "F" & LastRow)
                    .XValues = nWS.Range("A2", "A" & LastRow)
                    .HasDataLabels = True                    
                    .AxisGroup = xlSecondary        'turn this series into secondary y-axis with different scale
                End With
                
        Set srs4 = .SeriesCollection.NewSeries
                With srs4
                    .Name = nWS.Cells(1, "G")
                    .Values = nWS.Range("G2", "G" & LastRow)
                    .XValues = nWS.Range("A2", "A" & LastRow)
                    .HasDataLabels = True
                    '.AxisGroup = xlSecondary        'turn this series into secondary y-axis with different scale
                End With
        
        Set srs5 = .SeriesCollection.NewSeries
                With srs5
                    .Name = nWS.Cells(1, "H")
                    .Values = nWS.Range("H2", "H" & LastRow)
                    .XValues = nWS.Range("A2", "A" & LastRow)
                    .HasDataLabels = True                    
                    '.AxisGroup = xlSecondary        'turn this series into secondary y-axis with different scale
                End With
                
        
        Call FormatSrsNo(srs1)
        Call FormatSrsNo(srs2)
        Call FormatSrsNo(srs3)
        Call FormatSrsNo(srs4)
        Call FormatSrsNo(srs5)
                
    End With


'Excel.Application.CutCopyMode = False
Excel.Application.ScreenUpdating = True

'''save workbook
eWBook.Save

'''prevent button from being clicked twice
Me.cmdChart.Enabled = False

End Sub


Private Sub cmdCombAnalysis_Click()
        
        Dim hParentHwnd As OLE_CANCELBOOL
        
On Error GoTo ErrorHandler


'' Save the new workbook with combined data

    'dialog box to save file (as)
    Dim pGxDialog As IGxDialog
    Set pGxDialog = New GxDialog

    Dim pFilterCol As IGxObjectFilterCollection
    Set pFilterCol = pGxDialog

    Dim pEnumGx As IEnumGxObject

    'declare filter types
    Dim pFilter As IGxObjectFilter
    Set pFilter = New GxFilterDatasets

    'add filter
    pFilterCol.AddFilter pFilter, True    'the default filter

    'dialog box properties
    Dim sLocation As String
    pGxDialog.Title = "Create New File for Combined Data"
    pGxDialog.RememberLocation = True
    pGxDialog.StartingLocation = sLocation
    pGxDialog.DoModalSave (hParentHwnd)
    
    Dim pSaveFile As IGxObject
    
        'set filename and directory
        Set pSaveFile = pGxDialog.FinalLocation
        sLocation = pSaveFile.FullName

        'save input to text file
        Dim DataFile As String
        DataFile = sLocation & "\" & pGxDialog.Name
    
        'save excel file
        Set nBook = Excel.Application.Workbooks.Add
        nBook.Application.Visible = True
        nBook.SaveAs filename:=DataFile, FileFormat:=56


ErrorHandler:
    If Err.Number <> 0 Then
        MsgBox "Error Number " & Err.Number & vbNewLine & Err.Description
    ElseIf Err.Number = 1004 Then
        Exit Sub
    End If
    

'''prevent button from being clicked twice
Me.cmdCombAnalysis.Enabled = False

End Sub


Private Sub cmdErrorAnalysis_Click()

'''name fields
nBook.Sheets(1).Activate          'target sheet


'''loop through all worksheets (except first)
Dim iws As Long, j As Long
iws = nBook.Worksheets.count

For j = 2 To iws

    Dim ws As Worksheet
    Set ws = nBook.Worksheets(j)
    
    ''column headers
    ws.Range("I1") = "Error (Pts)"
    ws.Range("J1") = "Error (MAR)"
    ws.Range("K2") = "Standard Deviation (Points)"
    ws.Range("K3") = "Variance (Points)"
    ws.Range("K4") = "R (Points)"
    ws.Range("K6") = "Standard Deviation (MAR)"
    ws.Range("K7") = "Variance (MAR)"
    ws.Range("K8") = "R (MAR)"
    ws.Range("K12") = "Max"
    ws.Range("K13") = "Zero"


    '''calculate error analysis
    ''get row count
    Dim LastRow As Long
    LastRow = ws.Cells(Rows.count, "A").End(xlUp).Row


    ''calculations

    'errors
    Dim i As Long

    For i = 2 To LastRow
        ws.Range("I" & i).Formula = "=ABS(E" & i & "-F" & i & ")"
        ws.Range("J" & i).Formula = "=ABS(G" & i & "-H" & i & ")"
    Next i

    'standard deviation
    ws.Range("L2").Formula = "=STDEV(I2:I" & LastRow & ")"
    ws.Range("L6").Formula = "=STDEV(J2:J" & LastRow & ")"

    'variance
    ws.Range("L3").Formula = "=L2^2"
    ws.Range("L7").Formula = "=L6^2"

    'correlation coefficient
    ws.Range("L4").Formula = "=CORREL(E2:E" & LastRow & ",F2:F" & LastRow & ")"
    ws.Range("L8").Formula = "=CORREL(G2:G" & LastRow & ",H2:H" & LastRow & ")"
        
    'get maximum value to do x=y line
    ws.Range("L12") = "=MAX(E2:F" & LastRow & ")"       'Points
    ws.Range("L13") = "0"
    ws.Range("M12") = "=MAX(G2:H" & LastRow & ")"       'MAR
    ws.Range("M13") = "0"
    
    'format Interpolated value number
    ws.Range("E2:J" & LastRow).NumberFormat = "0.00"
        
    ''autofit columns
    ws.Columns("H:L").EntireColumn.AutoFit

Next j


'clear variables
i = 0
j = 0


nBook.Save

'''prevent button from being clicked twice
Me.cmdErrorAnalysis.Enabled = False

End Sub


Private Sub cmdErrorComb_Click()

On Error GoTo ErrorHandler

'''add new workbook for error analysis combination
Set eWBook = Workbooks.Add


'' Save the new workbook with combined data

    'dialog box to save file (as)
    Dim pGxDialog As IGxDialog
    Set pGxDialog = New GxDialog

    Dim pFilterCol As IGxObjectFilterCollection
    Set pFilterCol = pGxDialog

    Dim pEnumGx As IEnumGxObject

    'declare filter types
    Dim pFilter As IGxObjectFilter
    Set pFilter = New GxFilterDatasets

    'add filter
    pFilterCol.AddFilter pFilter, True    'the default filter

    'dialog box properties
    Dim sLocation As String
    pGxDialog.Title = "Create New File for Combined Error Analysis Data"
    pGxDialog.RememberLocation = True
    pGxDialog.StartingLocation = sLocation
    pGxDialog.DoModalSave hParentHwnd
    
    Dim pSaveFile As IGxObject
   
        Set pSaveFile = pGxDialog.FinalLocation
        sLocation = pSaveFile.FullName

        'save input to text file
        Dim DataFile As String
        DataFile = sLocation & "\" & pGxDialog.Name

        'save excel file
        eWBook.SaveAs filename:=DataFile, FileFormat:=56
        eWBook.Application.Visible = True
    

''worksheet headers
Dim eCombWS As Worksheet
Set eCombWS = eWBook.Worksheets(1)

With eCombWS
    .Name = "Error Analysis"
    
    .Range("A1") = "Stn_ID"
    .Range("B1") = "Longitude"
    .Range("C1") = "Latitude"
    .Range("D1") = "Std Dev (Pts)"
    .Range("E1") = "Variance (Pts)"
    .Range("F1") = "R (Pts)"
    .Range("G1") = "Std Dev (MAR)"
    .Range("H1") = "Variance (MAR)"
    .Range("I1") = "R (MAR)"
    .Range("J1") = "Analysis Count"

'header format
    With .Rows(1)
        .HorizontalAlignment = xlCenter
        With .Font
            .ColorIndex = 5
            .Bold = True
        End With
    End With
End With


'''copy data from nBook to eWBook

'worksheet count from Error Analysis workbook
Dim WScount As Long
WScount = nBook.Worksheets.count


Dim i As Long

For i = 2 To WScount

    Dim nWS As Worksheet
    Set nWS = nBook.Worksheets(i)
    
    'copy values
    With nWS
        .Range("A2:C2").Copy Destination:=eCombWS.Range("A" & i & ":C" & i)
        .Range("L2").Copy
        eCombWS.Range("D" & i).PasteSpecial xlPasteValues       'so that it copies values instead of formulas
        .Range("L3").Copy
        eCombWS.Range("E" & i).PasteSpecial xlPasteValues
        .Range("L4").Copy
        eCombWS.Range("F" & i).PasteSpecial xlPasteValues 
        .Range("L6").Copy                                       'MAR - cropped
        eCombWS.Range("G" & i).PasteSpecial xlPasteValues
        .Range("L7").Copy
        eCombWS.Range("H" & i).PasteSpecial xlPasteValues
        .Range("L8").Copy
        eCombWS.Range("I" & i).PasteSpecial xlPasteValues
        
        eCombWS.Range("J" & i).Value = nWS.Cells(Rows.count, "A").End(xlUp).Row - 1
    End With

Next i


'clear variables
i = 0


'''cleanup null values'''
Dim counter As Long
Dim RO As Long
RO = eCombWS.Range("A" & Rows.count).End(xlUp).Row

For counter = RO To 1 Step -1
    If eCombWS.Cells(counter, "J").Value = 1 Then
        eCombWS.Cells(counter, "J").EntireRow.Delete
    End If
Next counter


'autofit columns
Columns("A:J").EntireColumn.AutoFit

'save workbooks
nBook.Save
eWBook.Save

'''prevent button from being clicked twice
Me.cmdErrorComb.Enabled = False

Exit Sub


ErrorHandler:
    If Err.Number <> 0 Then
      MsgBox "Error Number " & Err.Number & vbNewLine & Err.Description
    End If

End Sub


Private Sub cmdExit2_Click()

Unload Me

End Sub


Private Sub cmdFolderAnalysis_Click()

On Error GoTo ErrorHandler

'''name fields
nBook.Sheets(1).Activate          'target sheet
Range("A1") = "Stn_ID"
Range("B1") = "Longitude"
Range("C1") = "Latitude"
Range("D1") = "RF_Date"
Range("E1") = "Interpolated Value (mm)"
Range("F1") = "Measured Value (mm)"
Range("G1") = "MAR (mm)"
Range("H1") = "MAR (complete)"


'rename sheet in new workbook
Dim strExtWS As String
strExtWS = "Combination"
nBook.Sheets(1).Name = strExtWS


'''copy all xls from specified folder

    'dialog box to open folder
    Dim pGxDialog As IGxDialog
    Set pGxDialog = New GxDialog

    Dim pFilterCol As IGxObjectFilterCollection
    Set pFilterCol = pGxDialog

    Dim pEnumGx As IEnumGxObject

    'declare filter types (folders)
    Dim pFilter As IGxObjectFilter
    Set pFilter = New GxFilterBasicTypes

    'add filter
    pFilterCol.AddFilter pFilter, True    'the default filter

    'dialog box properties
    pGxDialog.Title = "Select Directory with Files to be Analyzed"
    pGxDialog.ButtonCaption = "Select"
    pGxDialog.RememberLocation = True
    pGxDialog.DoModalOpen 0, pEnumGx

    Dim pFolderPath As IGxObject
    Set pFolderPath = pEnumGx.Next

    If pFolderPath Is Nothing Then
        Exit Sub
    End If

    'specify directory
    Dim FolderName As String, FName As String
    Dim RFwb As Workbook

    FolderName = pFolderPath.FullName & "\"


'get row counts
Dim DB As Long
Dim LastRow As Long


'''search through subfolders

Dim fso As Scripting.FileSystemObject
Dim SourceFolder As Scripting.Folder, SubFolder As Scripting.Folder

    Set fso = New Scripting.FileSystemObject
    Set SourceFolder = fso.GetFolder(FolderName)

    
''loop through all xls files in subfolders (only)

    If SourceFolder.SubFolders.count > 0 Then
        For Each SubFolder In SourceFolder.SubFolders
        
            FName = Dir(SubFolder.Path & "\" & "*.xls")
            
                Do While FName <> ""
                    Set RFwb = Workbooks.Open(filename:=SubFolder.Path & "\" & FName)
                    
                    'row counts for new workbook and data workbooks
                    LastRow = nBook.Sheets(1).Range("A" & Rows.count).End(xlUp).Row
                    DB = RFwb.Sheets(2).Range("A" & Rows.count).End(xlUp).Row
        
                    RFwb.Activate
                    
                    'LR start frm previous op's last row
                    Range("A2:G" & DB).Copy Destination:=nBook.Sheets(1).Range("A" & LastRow + 1)
                    Range("G" & DB + 1).Copy Destination:=nBook.Sheets(1).Range("H" & LastRow + 1 & ":H" & LastRow + DB - 1)
                                      
                    RFwb.Close savechanges:=False
                    FName = Dir()

                Loop
                
         Next SubFolder

    End If
 
        
    'clear variables
    Set SourceFolder = Nothing
    Set fso = Nothing


'''sort according to Stn_ID
Dim finalrow As Long
finalrow = nBook.Sheets(1).Range("A" & Rows.count).End(xlUp).Row

nBook.Worksheets(1).Activate
Range(Cells(2, "A"), Cells(finalrow, "H")).Sort Key1:=Cells(2, "A")


'autofit columns
Columns("A:H").EntireColumn.AutoFit

nBook.Save
    

ErrorHandler:
    If Err.Number <> 0 Then
      MsgBox "Error Number " & Err.Number & vbNewLine & Err.Description
    End If

    
'''prevent button from being clicked twice
Me.cmdFolderAnalysis.Enabled = False

End Sub


Private Sub cmdRGraphs_Click()

'''declare new workbook specifically for R graphs
        ' Add a new workbook
        Dim RWB As Workbook
        Dim hParentHwnd As OLE_CANCELBOOL


'' Save the new workbook with combined data

    'dialog box to save file (as)
    Dim pGxDialog As IGxDialog
    Set pGxDialog = New GxDialog

    Dim pFilterCol As IGxObjectFilterCollection
    Set pFilterCol = pGxDialog

    Dim pEnumGx As IEnumGxObject

    'declare filter types
    Dim pFilter As IGxObjectFilter
    Set pFilter = New GxFilterDatasets

    'add filter
    pFilterCol.AddFilter pFilter, True    'the default filter

    'dialog box properties
    Dim sLocation As String
    pGxDialog.Title = "Create New Workbook for R Graphs"
    pGxDialog.RememberLocation = True
    pGxDialog.StartingLocation = sLocation
    pGxDialog.DoModalSave hParentHwnd
    
    Dim pSaveFile As IGxObject
    
        'set filename and directory
        Set pSaveFile = pGxDialog.FinalLocation
        sLocation = pSaveFile.FullName
 
        'save input to text file
        Dim DataFile As String
        DataFile = sLocation & "\" & pGxDialog.Name
    
        'save excel file
        Set RWB = Excel.Application.Workbooks.Add
        RWB.Application.Visible = True
        RWB.SaveAs filename:=DataFile, FileFormat:=56

        
''delete worksheet 2 and 3
RWB.Sheets("Sheet2").Delete
RWB.Sheets("Sheet3").Delete


'''loop through all worksheets (except first)

''get source data workbook
Dim iws As Long, j As Long
Dim ws As Worksheet
nBook.Activate
iws = nBook.Worksheets.count

For j = 2 To iws

    Set ws = nBook.Worksheets(j)
    
    ''add new chart
    Dim gChart As Excel.Chart

    RWB.Activate
    RWB.Charts.Add after:=RWB.Sheets(RWB.Sheets.count)

    Set gChart = RWB.ActiveChart


    ''set row number
    Dim LastRow As Long
    LastRow = ws.Range("E" & ws.Rows.count).End(xlUp).Row


    ''set chart properties

    With gChart
        .ChartType = xlXYScatterLines
        
        .Location xlLocationAsNewSheet
        .HasTitle = True
        .ChartTitle.Characters.Text = "Measured vs Interpolated Rainfall Depth"
        .ChartTitle.Characters.Font.size = 12                                               'import to thesis
        .ChartTitle.Characters.Font.Name = "Times New Roman"                                'import to thesis
        .Name = ws.Name
    
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Interpolated Values"
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Font.size = 9                     'import to thesis
        
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Measured Values"
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Font.size = 9                        'import to thesis
        .Axes(xlValue, xlPrimary).MinimumScaleIsAuto = True
        .Axes(xlValue, xlPrimary).MaximumScaleIsAuto = True
                        
        .Axes(xlCategory).HasMajorGridlines = False
        .Axes(xlCategory).HasMinorGridlines = False
        .Axes(xlValue).HasMajorGridlines = True
        .Axes(xlValue).HasMinorGridlines = False
                
        .Axes(xlCategory).TickLabelSpacing = 1
        
        .HasLegend = False
 
       
        '' Add each series
        Dim srs1 As Series, srs2 As Series
                
        Set srs1 = .SeriesCollection.NewSeries
                With srs1
                    .Name = ws.Cells(1, "F")
                    '.Type = xlXYScatter
                    .XValues = ws.Range("E2", "E" & LastRow)
                    .Values = ws.Range("F2", "F" & LastRow)
                    .Format.Line.Visible = msoFalse                    
                    .MarkerSize = 2                                 'import to thesis
                    .MarkerForegroundColor = RGB(0, 0, 0)           'import to thesis
                End With
                     
        Set srs2 = .SeriesCollection.NewSeries
                With srs2
                    .Values = ws.Range("L12:L13")
                    .XValues = ws.Range("L12:L13")
                    .MarkerStyle = xlMarkerStyleNone
                    .Trendlines.Add
                End With
        
                
        With .Shapes.AddTextbox(msoTextOrientationHorizontal, 30.75, 11.25, 100, 21).TextFrame 
            .Characters.Text = "R=" & ws.Range("L4").Text
            .AutoSize = True
        End With
            
    End With


    ''clear variables
    LastRow = 0
    Set srs1 = Nothing
    Set srs2 = Nothing
    Set ws = Nothing
    Set gChart = Nothing
        
Next j


'clear variables
j = 0


'''save workbooks
RWB.Sheets("Sheet1").Delete
RWB.Save

Excel.Application.ScreenUpdating = True

'''prevent button from being clicked twice
Me.cmdRGraphs.Enabled = False


Exit Sub
ErrorHandler:
    If Err.Number <> 0 Then
        MsgBox "Error Number " & Err.Number & vbNewLine & Err.Description
    ElseIf Err.Number = 1004 And Err.Description = "Application-defined or object-defined error" Then
        Resume Next
    End If
Exit Sub

End Sub


Private Sub cmdRGraphsMAR_Click()

'''declare new workbook specifically for R graphs
        ' Add a new workbook
        Dim RWBmar As Workbook
        Dim hParentHwnd As OLE_CANCELBOOL


'' Save the new workbook with combined data

    'dialog box to save file (as)
    Dim pGxDialog As IGxDialog
    Set pGxDialog = New GxDialog

    Dim pFilterCol As IGxObjectFilterCollection
    Set pFilterCol = pGxDialog

    Dim pEnumGx As IEnumGxObject

    'declare filter types
    Dim pFilter As IGxObjectFilter
    Set pFilter = New GxFilterDatasets

    'add filter
    pFilterCol.AddFilter pFilter, True    'the default filter

    'dialog box properties
    Dim sLocation As String
    pGxDialog.Title = "Create New Workbook for R Graphs"
    pGxDialog.RememberLocation = True
    pGxDialog.StartingLocation = sLocation
    pGxDialog.DoModalSave hParentHwnd
    
    Dim pSaveFile As IGxObject
    
        'set filename and directory
        Set pSaveFile = pGxDialog.FinalLocation
        sLocation = pSaveFile.FullName

        'save input to text file
        Dim DataFile As String
        DataFile = sLocation & "\" & pGxDialog.Name
    
        'save excel file
        Set RWBmar = Excel.Application.Workbooks.Add
        RWBmar.Application.Visible = True
        RWBmar.SaveAs filename:=DataFile, FileFormat:=56


''delete worksheet 2 and 3
RWBmar.Sheets("Sheet2").Delete
RWBmar.Sheets("Sheet3").Delete


'''loop through all worksheets (except first)

''get source data workbook
Dim iws As Long, j As Long
Dim ws As Worksheet
nBook.Activate
iws = nBook.Worksheets.count

For j = 2 To iws

    Set ws = nBook.Worksheets(j)
    
    ''add new chart
    Dim gChart As Excel.Chart

    RWBmar.Activate
    RWBmar.Charts.Add after:=RWBmar.Sheets(RWBmar.Sheets.count)

    Set gChart = RWBmar.ActiveChart

    ''set row number
    Dim LastRow As Long
    LastRow = ws.Range("E" & ws.Rows.count).End(xlUp).Row

    
    ''set chart properties
    With gChart
        .ChartType = xlXYScatterLines
        '.SetSourceData Source:=ws.Range("E1:F" & LastRow), PlotBy:=xlColumns
        
        .Location xlLocationAsNewSheet
        .HasTitle = True
        .ChartTitle.Characters.Text = "Measured vs Interpolated Rainfall Depth"
        .ChartTitle.Characters.Font.size = 12                                               'import to thesis
        .ChartTitle.Characters.Font.Name = "Times New Roman"                                'import to thesis
        .Name = ws.Name
    
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Interpolated Values"
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Font.size = 9                     'import to thesis
        
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Measured Values"
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Font.size = 9                        'import to thesis
        .Axes(xlValue, xlPrimary).MinimumScaleIsAuto = True
        .Axes(xlValue, xlPrimary).MaximumScaleIsAuto = True
                        
        .Axes(xlCategory).HasMajorGridlines = False
        .Axes(xlCategory).HasMinorGridlines = False
        .Axes(xlValue).HasMajorGridlines = True
        .Axes(xlValue).HasMinorGridlines = False

        .Axes(xlCategory).TickLabelSpacing = 1
        
        .HasLegend = False
   

        '' Add each series
        Dim srs1 As Series, srs2 As Series
                
        Set srs1 = .SeriesCollection.NewSeries
                With srs1
                    .Name = ws.Cells(1, "H")
                    '.Type = xlXYScatter
                    .XValues = ws.Range("G2", "G" & LastRow)
                    .Values = ws.Range("H2", "H" & LastRow)
                    .Format.Line.Visible = msoFalse
                    
                    .MarkerSize = 2                                 'import to thesis
                    .MarkerForegroundColor = RGB(0, 0, 0)           'import to thesis
                End With
              
       
        Set srs2 = .SeriesCollection.NewSeries
                With srs2
                    .Values = ws.Range("M12:M13")
                    .XValues = ws.Range("M12:M13")
                    .MarkerStyle = xlMarkerStyleNone
                    .Trendlines.Add
                End With
                 
        With .Shapes.AddTextbox(msoTextOrientationHorizontal, 30.75, 11.25, 100, 21).TextFrame
            .Characters.Text = "R=" & ws.Range("L8").Text
            .AutoSize = True
        End With
            
    End With


    ''clear variables
    LastRow = 0
    Set srs1 = Nothing
    Set srs2 = Nothing
    Set ws = Nothing
    Set gChart = Nothing
       
Next j


'clear variables
j = 0


'''save workbooks
RWBmar.Sheets("Sheet1").Delete
RWBmar.Save

Excel.Application.ScreenUpdating = True

'''prevent button from being clicked twice
Me.cmdRGraphsMAR.Enabled = False


Exit Sub
ErrorHandler:
    If Err.Number <> 0 Then
        MsgBox "Error Number " & Err.Number & vbNewLine & Err.Description
    ElseIf Err.Number = 1004 And Err.Description = "Application-defined or object-defined error" Then
        Resume Next
    End If
Exit Sub

End Sub


Private Sub cmdSplitWS_Click()

Dim LastRow As Long, LastCol As Integer, i As Long, iStart As Long, iEnd As Long
Dim ws As Worksheet

Excel.Application.ScreenUpdating = False
nBook.Sheets("Sheet2").Delete
nBook.Sheets("Sheet3").Delete


With ActiveSheet
    'declare number of rows and columns in combined data table
    LastRow = .Cells(Rows.count, "A").End(xlUp).Row
    LastCol = .Cells(1, Columns.count).End(xlToLeft).Column
    
    .Range(.Cells(2, 1), Cells(LastRow, LastCol)).Sort Key1:=Range("A2"), Order1:=xlAscending, _
        Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
    iStart = 2
    
    For i = 2 To LastRow
    
        'if Stn_ID is different between subsequent rows, start splitting
        If .Range("A" & i).Value <> .Range("A" & i + 1).Value Then
            iEnd = i                                                'row number terminates when Stn_ID is different
            Sheets.Add after:=Sheets(Sheets.count)                  'add new sheet
            
            Set ws = ActiveSheet
            
            On Error Resume Next
            ws.Name = .Range("A" & iStart).Value                    'name new sheet according to Stn_ID
            On Error GoTo 0
            
            'copy worksheet header
            ws.Range(Cells(1, 1), Cells(1, LastCol)).Value = .Range(.Cells(1, 1), .Cells(1, LastCol)).Value
            
            'worksheet header format
            With ws.Rows(1)
                .HorizontalAlignment = xlCenter
                With .Font
                    .ColorIndex = 5
                    .Bold = True
                End With
            
            End With
            
            'copy data from specified row in data table to new sheet
            .Range(.Cells(iStart, 1), .Cells(iEnd, LastCol)).Copy Destination:=ws.Range("A2")
                      
            'autofit columns
            Columns("A:H").EntireColumn.AutoFit
            
            iStart = iEnd + 1
        
        End If
    
    Next i

End With

Excel.Application.CutCopyMode = False
Excel.Application.ScreenUpdating = True


'save workbook
Excel.ActiveWorkbook.Save

'''prevent button from being clicked twice
Me.cmdSplitWS.Enabled = False

End Sub


Module
ModAll.bas


Public pubDate As String
Public preDate As String
Public pubSaveDir As String
Public pubGrpLayerName As String
Public pExcel As Excel.Application
Public pAccess As Access.Application
Public wb As Workbook
Public nBook As Workbook
Public eWBook As Workbook


Public Function FindLayerByNameInGroup(pMap As IMap, sName As String) As ILayer
Dim pLayer As ILayer
Dim pCompositeLayer As ICompositeLayer
Dim i, j As Integer


' loop through each layer in the focus map
For i = 0 To pMap.LayerCount - 1
    Set pLayer = pMap.Layer(i)
    ' check to see if the layer is a group layer
    If TypeOf pLayer Is IGroupLayer Then
        ' if so use the composite layer interface to extract each layer from the group
        Set pCompositeLayer = pLayer
        ' loop through each layer in the group layer
        For j = 0 To pCompositeLayer.count - 1
            If pCompositeLayer.Layer(j).Name = sName Then
                Set FindLayerByNameInGroup = pCompositeLayer.Layer(j)
            End If
        Next j
        
    End If

Next i


End Function


Public Function FindLayerByName(pMap As IMap, sName As String) As ILayer
  Dim i As Integer
  For i = 0 To pMap.LayerCount - 1
    If pMap.Layer(i).Name = sName Then
      Set FindLayerByName = pMap.Layer(i)
    End If
  Next
End Function


Public Function FindLayerByNameNGroup(pMap As IMap, sName As String, sGroupName As String) As ILayer
Dim pLayer As ILayer
Dim pCompositeLayer As ICompositeLayer
Dim i, j As Integer


' loop through each layer in the focus map
For i = 0 To pMap.LayerCount - 1
    Set pLayer = pMap.Layer(i)
    ' check to see if the layer is a group layer
    If TypeOf pLayer Is IGroupLayer And pLayer.Name = sGroupName Then
        ' if so use the composite layer interface to extract each layer from the group
        Set pCompositeLayer = pLayer
        ' loop through each layer in the group layer
        For j = 0 To pCompositeLayer.count - 1
            If pCompositeLayer.Layer(j).Name = sName Then
                Set FindLayerByNameNGroup = pCompositeLayer.Layer(j)
            End If
        Next j
        
    End If

Next i

End Function


Public Sub AlignY(FreeParam As Integer)
      '' FreeParam: AXIS ALLOWED TO VARY
      '' 1: Y1 (PRI) MIN
      '' 2: Y1 (PRI) MAX
      '' 3: Y2 (SEC) MIN
      '' 4: Y2 (SEC) MAX

      Dim Y1min As Double
      Dim Y1max As Double
      Dim Y2min As Double
      Dim Y2max As Double

    With ActiveChart
        With .Axes(2, 1)
          Y1min = .MinimumScale
          Y1max = .MaximumScale
          .MinimumScaleIsAuto = False
          .MaximumScaleIsAuto = False
        End With
        
        With .Axes(2, 2)
          Y2min = .MinimumScale
          Y2max = .MaximumScale
          .MinimumScaleIsAuto = False
          .MaximumScaleIsAuto = False
        End With
        
        Select Case FreeParam
          Case 1
            If Y2max <> 0 Then _
              .Axes(2, 1).MinimumScale = Y2min * Y1max / Y2max
          Case 2
            If Y2min <> 0 Then _
              .Axes(2, 1).MaximumScale = Y1min * Y2max / Y2min
          Case 3
            If Y1max <> 0 Then _
              .Axes(2, 2).MinimumScale = Y1min * Y2max / Y1max
          Case 4
            If Y1min <> 0 Then _
              .Axes(2, 2).MaximumScale = Y2min * Y1max / Y1min
        End Select
      
     End With


End Sub


Public Sub FormatSrsNo(srs As Series)

Dim pt As Point, pts As Points
Dim i As Integer
Set pts = srs.Points

For i = 1 To pts.count
    srs.Points(i).DataLabel.NumberFormat = "#,##0.00"
Next i

End Sub


' Returns the map or scene document file path. Can be unreliable if the document has
' yet to be saved to disk.

Public Function GetDocPath(pApp As IApplication) As String

  Dim pTemplates As ITemplates
  Dim lTempCount As Long
  
  Set pTemplates = pApp.Templates
  lTempCount = pTemplates.count
    
  ' The document is always the last item
  GetDocPath = pTemplates.Item(lTempCount - 1)
  
End Function
