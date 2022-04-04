Option Explicit


Private Sub cmdBack2_Click()

Me.Hide
frmKrigeVariogram.Show

End Sub


Private Sub cmdExit3_Click()

'''clear all public variables'''
pubDate = ""
preDate = ""
pubSaveDir = ""
pubGrpLayerName = ""

Access.CloseCurrentDatabase
pExcel.Quit

Unload Me

End Sub


Private Sub cmdExport_Click()

    ''Get the focused map from MapDocument
    Dim pMxDoc As IMxDocument
    Dim pMap As IMap
    Set pMxDoc = ThisDocument
    Set pMap = pMxDoc.FocusMap


  '''start Excel operation'''
        
        On Error Resume Next
        
        Set pExcel = GetObject(, "Excel.Application")
        
        If pExcel Is Nothing Then   'No instance of Excel is available
            Set pExcel = CreateObject("Excel.Application")
        End If
        
        ' Make it visible
        pExcel.Visible = True
  
        ' Add a new workbook
        Set wb = pExcel.Workbooks.Add
        
        'Add the data here
        Dim ws As Worksheet
        Set ws = wb.Sheets(1)
        ws.Name = "RF" & pubDate
         
        'add MAR worksheet
        Dim MARWs As Worksheet
        Set MARWs = wb.Sheets(3)
        MARWs.Name = "MAR" & pubDate
      
        
'''write to Excel'''

''write to first sheet - rainfall interpolation sampled values
                
        ''Declare export function variables        
        Dim lngRow As Long, lngCol As Long
        Dim Data
        Dim i As Integer
        Dim txt As String, Char As String
        Dim filepath As String
                
        'filepath of file to be opened
        filepath = pubSaveDir & "\RF" & pubDate & ".txt"
        
        
Dim fnum1 As Integer
fnum1 = FreeFile()

Open filepath For Input As #fnum1

    lngRow = 0
    lngCol = 0
    txt = ""

    With ws.Application

        Do Until EOF(fnum1)                                     'End of File

            Line Input #fnum1, Data
            
            For i = 1 To Len(Data)
                Char = Mid(Data, i, 1)
                If Char = "," Then                              'delimiter
                    .ActiveCell.Offset(lngRow, lngCol) = txt
                    lngCol = lngCol + 1
                    txt = ""
                ElseIf i = Len(Data) Then
                    If Char <> Chr(34) Then txt = txt & Char
                    .ActiveCell.Offset(lngRow, lngCol) = txt
                    txt = ""
                ElseIf Char <> Chr(34) Then
                    txt = txt & Char
                End If
            Next i
    
            lngCol = 0
            lngRow = lngRow + 1
        
        Loop
        
    End With

Close #fnum1


''write to second sheet - mean areal

'declare variables
        Dim lngRow2 As Long, lngCol2 As Long
        Dim Data2
        Dim j As Integer
        Dim txt2 As String, Char2 As String
        Dim Filepath2 As String
                
        'filepath of file to be opened
        Filepath2 = pubSaveDir & "\MAR" & pubDate & ".txt"
        
        'clear variables
        lngRow2 = 0
        lngCol2 = 0
        txt2 = ""


Dim fnum2 As Integer
fnum2 = FreeFile()
  
Open Filepath2 For Input As #fnum2

    MARWs.Activate

    With MARWs.Application

        Do Until EOF(fnum2)                                         'End of File

            Line Input #fnum2, Data2
            For j = 1 To Len(Data2)
                Char2 = Mid(Data2, j, 1)
                If Char2 = "," Then                                 'delimiter'
                    .ActiveCell.Offset(lngRow2, lngCol2) = txt2
                    lngCol2 = lngCol2 + 1
                    txt2 = ""
                ElseIf j = Len(Data2) Then
                    If Char2 <> Chr(34) Then txt2 = txt2 & Char2
                    .ActiveCell.Offset(lngRow2, lngCol2) = txt2
                    txt2 = ""
                ElseIf Char2 <> Chr(34) Then
                    txt2 = txt2 & Char2
                End If
            Next j
    
            lngCol2 = 0
            lngRow2 = lngRow2 + 1
        
        Loop
        
    End With

Close #fnum2
   
  
'''save xls to a directory'''
Dim filename As String
filename = pubSaveDir & "\RF" & pubDate & ".xls"
wb.SaveAs filename, FileFormat:=56

'''prevent button from being clicked twice
Me.cmdExport.Enabled = False

End Sub


Private Sub cmdExtractData_Click()

''declare variables
Dim firWS As Worksheet, secWS As Worksheet, thiWS As Worksheet
Set firWS = wb.Sheets(1)
Set secWS = wb.Sheets(2)
Set thiWS = wb.Sheets(3)

firWS.Activate

' Make it visible
pExcel.Application.Visible = True

'get first sheet's name
Dim strWSName As String
strWSName = firWS.Name

'rename second sheet
Dim strExtWS As String
strExtWS = strWSName & "Ext"
secWS.Name = strExtWS

'get row counts
Dim LR As Long
LR = secWS.Range("B" & Rows.count).End(xlUp).Row            'target sheet

Dim RO As Long
RO = firWS.Range("A" & Rows.count).End(xlUp).Row            'first sheet


''copy station numbers, date, and measured rainfall depths''
Dim dbfWB As Workbook
Set dbfWB = Workbooks.Open(wb.Path & "\RF0.dbf")


With dbfWB
    .Activate

    Dim DB As Long
    DB = .Sheets(1).Range("A" & Rows.count).End(xlUp).Row

    .Sheets(1).Range("A2:A" & DB).Copy Destination:=secWS.Range("A" & LR + 1)      'target sheet
    .Sheets(1).Range("B2:B" & DB).Copy Destination:=secWS.Range("D" & LR + 1)      'target sheet
    .Sheets(1).Range("C2:C" & DB).Copy Destination:=secWS.Range("F" & LR + 1)      'target sheet

    .Close

End With


''copy station coordinates
With firWS
    .Activate
    .Range("B2:C" & RO).Copy Destination:=secWS.Range("B" & LR + 1)      'sampling of interps
End With


''copy MAR from thiWS
With thiWS
    .Activate
    .Range("B2:B" & RO).Copy Destination:=secWS.Range("G" & LR + 1)     'MAR values from cropped and uncropped (cropped only)
    .Range("B1").Copy Destination:=secWS.Range("G" & RO + 1)            'cropped mean of complete interp
    .Range("B1").Copy Destination:=secWS.Range("H2:H" & RO)             'column of cropped mean of complete interp
End With


''copy interpolated values
Dim IV As Long
IV = 2
    
    Dim strABC(1 To 26) As String
    Dim intCounter As Integer

    For intCounter = 4 To 21
        strABC(intCounter) = Chr$(intCounter + 64)
                
            firWS.Range(strABC(intCounter) & IV).Copy Destination:=secWS.Range("E" & IV)      'target sheet
            
            IV = IV + 1
                        
    Next intCounter


With secWS              'target sheet
    .Activate
    
    ''name fields
    .Range("A1") = "Stn_ID"
    .Range("B1") = "Longitude"
    .Range("C1") = "Latitude"
    .Range("D1") = "RF_Date"
    .Range("E1") = "Interpolated Value (mm)"
    .Range("F1") = "Measured Value (mm)"
    .Range("G1") = "Mean Areal Rainfall (MAR)"
    .Range("H1") = "MAR (complete)"
    

    ''verify Interpolated values (delete rows with "0", change negative values to "0")
    Dim counter As Long
    For counter = RO To 1 Step -1
        If .Cells(counter, "E").Value = 0 Then
            .Cells(counter, "E").EntireRow.Delete
        ElseIf Cells(counter, "E").Value < 0 Then
            .Cells(counter, "E").Value = 0
        Else
        End If
    
    Next counter
    
End With


''autofit columns
secWS.Columns("A:I").EntireColumn.AutoFit

''clear variables
strWSName = ""
strExtWS = ""
LR = 0
RO = 0
DB = 0
IV = 0
intCounter = 0

'save workbook
wb.Save

'''prevent button from being clicked twice
Me.cmdExtractData.Enabled = False

End Sub


Private Sub cmdMAR_Click()

    '''Get the focused map from MapDocument'''
    Dim pMxDoc As IMxDocument
    Dim pMap As IMap
    Set pMxDoc = ThisDocument
    Set pMap = pMxDoc.FocusMap


  ''' Get table row count'''    
    Dim pTable As ITable
    Dim pLayer As ILayer
    Dim pOriFLayer As IFeatureLayer
    Dim LayerName As String
        
    LayerName = "RF0" 
    Set pLayer = FindLayerByNameNGroup(pMap, LayerName, pubGrpLayerName)
    Set pOriFLayer = pLayer
    Set pTable = pOriFLayer.FeatureClass

    
    'Query filter to get row count
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter

    pQueryFilter.WhereClause = "Stn_ID is not null"

    Dim pRowCount As Integer
    pRowCount = pTable.RowCount(pQueryFilter)

   
'''search and calculate all rasters in specified group layer'''

''declare uncropped rasters
Dim il As Integer
Dim RasLyrName As String
Dim fnum As Integer
fnum = FreeFile()

Dim pRasterlayer As IRasterLayer
Dim pRasterBandCollection As IRasterBandCollection
Dim pRasterBand As IRasterBand
Dim pRasterStat As IRasterStatistics


''declare cropped rasters
Dim MRasLyrName As String
Dim pMRasterlayer As IRasterLayer
Dim pMRasterBandCollection As IRasterBandCollection
Dim pMRasterBand As IRasterBand
Dim pMRasterStat As IRasterStatistics


''save input to text file
Dim MARFile As String
MARFile = pubSaveDir & "\MAR" & pubDate & ".txt"


''loop through all raster layers
For il = 0 To pRowCount

    RasLyrName = "VarRF" & il
    MRasLyrName = "ExtRF" & il

    Set pRasterlayer = FindLayerByNameNGroup(pMap, RasLyrName, pubGrpLayerName)
    Set pMRasterlayer = FindLayerByNameNGroup(pMap, MRasLyrName, pubGrpLayerName)
    
    Set pRasterBandCollection = pRasterlayer.Raster
    Set pRasterBand = pRasterBandCollection.Item(0)
    Set pRasterStat = pRasterBand.Statistics

    Set pMRasterBandCollection = pMRasterlayer.Raster
    Set pMRasterBand = pMRasterBandCollection.Item(0)
    Set pMRasterStat = pMRasterBand.Statistics


    'write mean values to text file
    If il = 0 Then
        Open MARFile For Output As #fnum
    Else
        Open MARFile For Append As #fnum
    End If
        
        Write #fnum, pRasterStat.Mean, pMRasterStat.Mean
        Close #fnum


    'clear variables
    Set pRasterlayer = Nothing
    Set pRasterBand = Nothing
    Set pRasterStat = Nothing
    Set pMRasterlayer = Nothing
    Set pMRasterBand = Nothing
    Set pMRasterStat = Nothing

Next il
    

'''prevent button from being clicked twice
Me.cmdMAR.Enabled = False

End Sub


Private Sub cmdMask_Click()

    'Get the focused map from MapDocument
    Dim pDoc As IMxDocument
    Dim pMap As IMap
    Set pDoc = ThisDocument
    Set pMap = pDoc.FocusMap
                
    'Get table with n rows    
    Dim pTable As ITable
    Dim pLayer As ILayer
    Dim pOriFLayer As IFeatureLayer
    Dim LayerName As String
        
    LayerName = "RF0" 
    Set pLayer = FindLayerByNameNGroup(pMap, LayerName, pubGrpLayerName)
    Set pOriFLayer = pLayer
    Set pTable = pOriFLayer.FeatureClass

    'Query filter to get row count
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter

    pQueryFilter.WhereClause = "Stn_ID is not null"

    Dim pRowCount As Integer
    pRowCount = pTable.RowCount(pQueryFilter)
        
        
'statusbar progress message
Application.StatusBar.Message(0) = "Extracting raster files... Please wait..."
        
        
'start loop for n tables
Dim il As Integer
For il = 0 To pRowCount
        
    'Get the feature class from layer name in Group Layer
    Dim sName As String
    sName = "VarRF" & il
           
    Dim pInRasLyr As IRasterLayer
    Dim pInRasBCol As IRasterBandCollection
    Dim pInRasB As IRasterBand
    
    Set pInRasLyr = FindLayerByNameNGroup(pMap, sName, pubGrpLayerName)
    Set pInRasBCol = pInRasLyr.Raster
    Set pInRasB = pInRasBCol.Item(0)
    
    ' Create the RasterExtractionOp object
    Dim pExtractionOp As IExtractionOp
    Set pExtractionOp = New RasterExtractionOp

    ' Declare mask layer
    Dim pMaskRaster As IFeatureLayer
    Dim pMaskFC As IFeatureClass

    ' open mask raster datasets
    Set pMaskRaster = FindLayerByNameNGroup(pMap, "Basin", "Background")
    Set pMaskFC = pMaskRaster.FeatureClass
         
    ' Declare the output dataset
    Dim pOutputRaster As IRaster

    ' Call the method - extract by mask
    Set pOutputRaster = pExtractionOp.Raster(pInRasB, pMaskFC)
        
    'Add output into ArcMap as a raster layer
    Dim pOutRasLayer As IRasterLayer
    Set pOutRasLayer = New RasterLayer

    'set raster to be first band of the raster
    Dim pRasterDS As IRasterDataset
    Dim pRasBandCol As IRasterBandCollection

    Set pRasBandCol = pOutputRaster
    Set pRasterDS = pRasBandCol.Item(0).RasterDataset

    'Create a new raster layer and save in folder (make it permanent)  
    Dim pWorkspaceFactory As IWorkspaceFactory
    Dim pRasterWorkspace As IRasterWorkspace
    Dim pTempDS As ITemporaryDataset
    Dim pDataset As IDataset
  
    Set pWorkspaceFactory = New RasterWorkspaceFactory
    Set pRasterWorkspace = pWorkspaceFactory.OpenFromFile(pubSaveDir, Application.hWnd)   'raster directory
  
    Set pDataset = pRasterDS
    Set pTempDS = pRasterDS
    Set pRasterDS = pTempDS.MakePermanentAs("ExtRF" & il, pRasterWorkspace, "GRID")
        
    pOutRasLayer.CreateFromDataset pRasterDS
    pOutRasLayer.Name = "ExtRF" & il

    'Add layer to Group Layer in Map
    Dim pFinalLayer As IRasterLayer
    Set pFinalLayer = pOutRasLayer
  
    Dim pMapLayers As IMapLayers
    Set pMapLayers = pMap
  
    Dim pGrpLayer As IGroupLayer
    Set pGrpLayer = FindLayerByName(pMap, pubGrpLayerName)
    
    pMapLayers.InsertLayerInGroup pGrpLayer, pFinalLayer, True, 0

    'clear variables
    Set pRasterDS = Nothing
    Set pOutRasLayer = Nothing
    Set pOutputRaster = Nothing
    Set pFinalLayer = Nothing
    
Next il


pDoc.ActivatedView.PartialRefresh esriViewGeography, Nothing, Nothing

'''prevent button from being clicked twice
Me.cmdMask.Enabled = False

End Sub


Private Sub cmdNext3_Click()

Unload Me
frmAnalysis.Show

End Sub


Private Sub cmdSample_Click()

    'Get the focused map from MapDocument
    Dim pDoc As IMxDocument
    Dim pMap As IMap
    Set pDoc = ThisDocument
    Set pMap = pDoc.FocusMap

    'set input location raster object
    Dim SamplePtName As String
    SamplePtName = "RF0" 
    Dim pInLayer As ILayer
    Set pInLayer = FindLayerByNameNGroup(pMap, SamplePtName, pubGrpLayerName)

    Dim pFLayer As IFeatureLayer
    Set pFLayer = pInLayer
    
    Dim pFClass As IFeatureClass
    Set pFClass = pFLayer.FeatureClass


'''extract point values from rasters'''

' Create the RasterExtractionOp object
Dim pExtractionOp As IExtractionOp
Set pExtractionOp = New RasterExtractionOp

' Create a raster of multiple bands
Dim pBandCol As IRasterBandCollection
Set pBandCol = New Raster
Dim pInputRasterDataset As IRasterDataset
Dim pRasterlayer As IRasterLayer
Dim pRasterBand As IRasterBand
Dim pRaster As IRaster
Dim pBandColTemp As IRasterBandCollection

  ' Get table row count    
    Dim pTable As ITable
    Dim pLayer As ILayer
    Dim pOriFLayer As IFeatureLayer
    Dim LayerName As String
        
    LayerName = "RF0" 
    Set pLayer = FindLayerByNameNGroup(pMap, LayerName, pubGrpLayerName)
    Set pOriFLayer = pLayer
    Set pTable = pOriFLayer.FeatureClass

    'Query filter to get row count
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter

    pQueryFilter.WhereClause = "Stn_ID is not null"

    Dim pRowCount As Integer
    pRowCount = pTable.RowCount(pQueryFilter)
        
'start loop for n tables
Dim il As Integer
Dim RasLyrName As String


For il = 0 To (pRowCount - 1)

    RasLyrName = "VarRF" & (il + 1)

    Set pRasterlayer = FindLayerByNameNGroup(pMap, RasLyrName, pubGrpLayerName)
    Set pRaster = pRasterlayer.Raster
    Set pBandColTemp = pRaster
    Set pRasterBand = pBandColTemp.Item(0)
    pBandCol.AppendBand pRasterBand

    'clear variables
    Set pRasterlayer = Nothing
    Set pRaster = Nothing
    Set pRasterBand = Nothing
    Set pBandColTemp = Nothing

Next il


' Declare the output table object
Dim pOutputTable As ITable
 
' Calls the method
Set pOutputTable = pExtractionOp.Sample(pFClass, pBandCol, esriGeoAnalysisResampleBilinear)


    '''convert temporary table into dbf in specified folder'''
    Dim pDataset As IDataset
    Dim pInDSNAme As IDatasetName

    Dim pFeatureClassName As IFeatureClassName
    Dim pOutDatasetName As IDatasetName
    Dim pWorkspaceName As IWorkspaceName

    Dim pExportOp As IExportOperation
   
    'Get the FcName from the featureclass
    Set pDataset = pOutputTable
    Set pInDSNAme = pDataset.FullName
    
    'Define the output feature class name
    Set pFeatureClassName = New FeatureClassName
    Set pOutDatasetName = pFeatureClassName
        
    pOutDatasetName.Name = "RF" & pubDate & ".txt"
    
    Set pWorkspaceName = New WorkspaceName
    pWorkspaceName.PathName = pubSaveDir    'save directory
    pWorkspaceName.WorkspaceFactoryProgID = "esriDataSourcesOleDB.TextFileWorkspaceFactory"
    Set pOutDatasetName.WorkspaceName = pWorkspaceName
   
    'Export
    Set pExportOp = New ExportOperation
    pExportOp.ExportFeatureClass pInDSNAme, Nothing, Nothing, Nothing, pOutDatasetName, 0


'''add dbf into TOC'''

'define input table
Dim pName As IName
Dim pFinalTable As ITable
Dim pFinalWSName As IWorkspaceName
Dim pFinalDSName As IDatasetName

'get the txt file by specifying its workspace and name
Set pFinalDSName = New TableName
Set pFinalWSName = New WorkspaceName

pFinalWSName.WorkspaceFactoryProgID = "esriDataSourcesOleDB.TextFileWorkspaceFactory"

pFinalWSName.PathName = pubSaveDir   'save directory
pFinalDSName.Name = "RF" & pubDate & ".txt"

Set pFinalDSName.WorkspaceName = pFinalWSName
Set pName = pFinalDSName

'open the txt table
Set pFinalTable = pName.Open

'add table to map
Dim pStandTableCollection As IStandaloneTableCollection
Dim pStandaloneTable As IStandaloneTable

Set pStandaloneTable = New StandaloneTable
Set pStandaloneTable.Table = pFinalTable
Set pStandTableCollection = pMap

pStandTableCollection.AddStandaloneTable pStandaloneTable

pDoc.UpdateContents

'''prevent button from being clicked twice
Me.cmdSample.Enabled = False

End Sub
