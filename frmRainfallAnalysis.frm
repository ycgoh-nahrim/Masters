Option Explicit


Private Sub cmdAccess_Click()

    'Get the focused map from MapDocument
    Dim pMxDoc As IMxDocument
    Dim pMap As IMap
    Set pMxDoc = ThisDocument
    Set pMap = pMxDoc.FocusMap


''get document file path - sFolder
    
    'declare application
    Dim pApp As IApplication
    Set pApp = New AppRef

    ' Create a FileSystemObject
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
  
    ' Save file name portion of file path
    Dim filepath As String
    filepath = GetDocPath(pApp)
    Dim sFolder As String
    sFolder = fso.GetParentFolderName(filepath)


  ''start Access operation
  
        'Open Access file
        Dim pAccess As Access.Application
               
        On Error Resume Next
        
        Set pAccess = GetObject(, "Access.Application")
                
        If pAccess Is Nothing Then   'No instance of Access is available
            Set pAccess = CreateObject("Access.Application")
        End If

        pAccess.OpenCurrentDatabase sFolder & "\Database.mdb"
                
        ' Make it visible
        pAccess.Visible = True
                
        'extract data from pubDate
        Dim yy As String, mm As String, dd As String
        yy = Mid(preDate, 1, 4)
        mm = Mid(preDate, 5, 2)
        dd = Mid(preDate, 7, 2)
        
        
        ''rain gauge(s) omission
        Dim omit As String
        Dim omit1 As String, omit2 As String, omit3 As String, omit4 As String
                
        omit1 = ""
        omit2 = ""
        omit3 = ""
        omit4 = ""

        
            If cboRGONo.Text = "0" Then
                omit = ""
            ElseIf cboRGONo.Text = "1" Then
                omit1 = " And (Not (RF_Stations.Stn_ID) = 'RG" & txtRGO1.Text & "')"
            ElseIf cboRGONo.Text = "2" Then
                omit1 = " And (Not (RF_Stations.Stn_ID) = 'RG" & txtRGO1.Text & "')"
                omit2 = " And (Not (RF_Stations.Stn_ID) = 'RG" & txtRGO2.Text & "')"
            ElseIf cboRGONo.Text = "3" Then
                omit1 = " And (Not (RF_Stations.Stn_ID) = 'RG" & txtRGO1.Text & "')"
                omit2 = " And (Not (RF_Stations.Stn_ID) = 'RG" & txtRGO2.Text & "')"
                omit3 = " And (Not (RF_Stations.Stn_ID) = 'RG" & txtRGO3.Text & "')"
            ElseIf cboRGONo.Text = "4" Then
                omit1 = " And (Not (RF_Stations.Stn_ID) = 'RG" & txtRGO1.Text & "')"
                omit2 = " And (Not (RF_Stations.Stn_ID) = 'RG" & txtRGO2.Text & "')"
                omit3 = " And (Not (RF_Stations.Stn_ID) = 'RG" & txtRGO3.Text & "')"
                omit4 = " And (Not (RF_Stations.Stn_ID) = 'RG" & txtRGO4.Text & "')"
            End If
        
        omit = omit1 & omit2 & omit3 & omit4
        
          
        'Run SQL command - First Query
        Dim DB As DAO.Database
        Set DB = pAccess.CurrentDb
           
        Dim strSQL1 As String
        
        strSQL1 = "SELECT RF_Stations.Stn_ID, RF_Stations.JPS_Code, RF_Data.RF_Date, RF_Data.RF_Depth " & _
                    "FROM (RF_Stations INNER JOIN RF_Data ON RF_Stations.Stn_ID = RF_Data.Stn_ID) INNER JOIN " & _
                    "StnType_Table ON (RF_Stations.Stn_ID = StnType_Table.Stn_ID) AND (RF_Data.Stn_Type = StnType_Table.Stn_Type) " & _
                    "WHERE (((RF_Data.RF_Date) = #" & mm & "/" & dd & "/" & yy & "#)" & omit & ") " & _
                    "ORDER BY RF_Stations.Stn_ID;"
                
        Dim qdf1 As DAO.QueryDef
        Set qdf1 = DB.QueryDefs("Daily_RF_Depth")
        qdf1.Sql = strSQL1
        DoCmd.OpenQuery "Daily_RF_Depth"
                

        'Second Query
        Dim strSQL2 As String
        
        strSQL2 = "SELECT RF_Data.Stn_ID, RF_Data.RF_Date, " & _
                    "Round(StDev(RF_Data.RF_Depth)*Var(RF_Data.RF_Depth)) AS [StDev * Var] " & _
                    "FROM RF_Stations INNER JOIN RF_Data ON RF_Stations.Stn_ID = RF_Data.Stn_ID " & _
                    "WHERE (((RF_Data.RF_Date) = #" & mm & "/" & dd & "/" & yy & "#)) " & _
                    "GROUP BY RF_Data.Stn_ID, RF_Data.RF_Date;"
              
        Dim qdf2 As DAO.QueryDef
        Set qdf2 = DB.QueryDefs("Daily_RF_Depth_StDev")
        qdf2.Sql = strSQL2
        DoCmd.OpenQuery "Daily_RF_Depth_StDev"


        'Third Query
        Dim strSQL3 As String
        
        strSQL3 = "SELECT Daily_RF_Depth_StDev.Stn_ID, Daily_RF_Depth_StDev.RF_Date, " & _
                    "Avg(RF_Data.RF_Depth) AS [Avg RF Depth], RF_Stations.Latitude, RF_Stations.Longitude " & _
                    "FROM RF_Stations INNER JOIN (Daily_RF_Depth_StDev INNER JOIN RF_Data ON Daily_RF_Depth_StDev.Stn_ID = RF_Data.Stn_ID) " & _
                    "ON (Daily_RF_Depth_StDev.Stn_ID = RF_Stations.Stn_ID) AND (RF_Stations.Stn_ID = RF_Data.Stn_ID)" & _
                    "WHERE (((RF_Data.RF_Date) = #" & mm & "/" & dd & "/" & yy & "#) And ((Daily_RF_Depth_StDev.[StDev * Var]) " & _
                    "< 50 Or (Daily_RF_Depth_StDev.[StDev * Var]) Is Null) And (([RF_Data.RF_Depth]) Is Not Null)" & omit & ") " & _
                    "GROUP BY Daily_RF_Depth_StDev.Stn_ID, Daily_RF_Depth_StDev.RF_Date, RF_Stations.Latitude, RF_Stations.Longitude;"
       
        Dim qdf3 As DAO.QueryDef
        Set qdf3 = DB.QueryDefs("Daily_RF_Depth_Corrected")
        qdf3.Sql = strSQL3
        DoCmd.OpenQuery "Daily_RF_Depth_Corrected"


        DB.Close
        
        'clear variables
        Set qdf1 = Nothing
        Set qdf2 = Nothing
        Set qdf3 = Nothing
        Set DB = Nothing


End Sub


Private Sub cmdAddData_Click()

''get document file path - sFolder
    
    'declare application
    Dim pApp As IApplication
    Set pApp = New AppRef

    ' Create a FileSystemObject
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
  
    ' Save file name portion of file path
    Dim filepath As String
    filepath = GetDocPath(pApp)
    Dim sFolder As String
    sFolder = fso.GetParentFolderName(filepath)


''Start operation

  Dim pPropset As IPropertySet
  Set pPropset = New PropertySet

  'set properties for OLE DB Connection
  pPropset.SetProperty "CONNECTSTRING", "Provider=Microsoft.Jet.OLEDB.4.0;Data source=" & sFolder & "\Database.mdb"

  Dim pWorkspaceFact As IWorkspaceFactory
  Set pWorkspaceFact = New OLEDBWorkspaceFactory

  'Create the new workspace/feature workspace objects
  Dim pWorkspace As IWorkspace
  Set pWorkspace = pWorkspaceFact.Open(pPropset, 0)
  
  Dim pFeatWorkspace As IFeatureWorkspace
  Set pFeatWorkspace = pWorkspace
  
  'Create the new table object from the dataset name
  Dim pTable As ITable
  Set pTable = pFeatWorkspace.OpenTable("Daily_RF_Depth_Corrected")


  ' add the table

    Dim pMxDoc As IMxDocument
    Set pMxDoc = ThisDocument

    Dim pMap As IMap
    Set pMap = pMxDoc.FocusMap

    Dim pStTab As IStandaloneTable
    Set pStTab = New StandaloneTable
    Set pStTab.Table = pTable
        
    Dim pStTabCol As IStandaloneTableCollection
    Set pStTabCol = pMap

    pStTabCol.AddStandaloneTable pStTab
    pMxDoc.UpdateContents


End Sub


Private Sub cmdAllAnalysis_Click()

Unload Me
frmAnalysis.Show

End Sub


Private Sub cmdBrowse_Click()

''dialog box to open file
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
    pGxDialog.Title = "Select New Folder Directory"
    pGxDialog.ButtonCaption = "Set"
    pGxDialog.RememberLocation = True
    pGxDialog.DoModalOpen 0, pEnumGx

    Dim pFolderPath As IGxObject
    Set pFolderPath = pEnumGx.Next

    If pFolderPath Is Nothing Then
        Exit Sub
    End If

''write new folder path and name to textbox
Me.txtNewFolder.Text = pFolderPath.FullName & "\RF" & pubDate


End Sub


Private Sub cmdDate_Click()

''get document file path - sFolder
    
    'declare application
    Dim pApp As IApplication
    Set pApp = New AppRef

    ' Create a FileSystemObject
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
  
    ' Save file name portion of file path
    Dim filepath As String
    filepath = GetDocPath(pApp)
    Dim sFolder As String
    sFolder = fso.GetParentFolderName(filepath)


'extract pubDate year
preDate = frmRainfallAnalysis.txtDate.Text


If preDate = "" Then
    MsgBox "Please enter date", vbRetryCancel, "Error"
ElseIf Me.txtDate.TextLength <> 8 Then
    MsgBox "Please enter date according to this format: yyyymmdd", vbRetryCancel, "Error"
Else
    pubDate = Mid(preDate, 3, 6)
    frmRainfallAnalysis.txtNewFolder.Text = sFolder & "\RF" & pubDate
End If


End Sub


Private Sub cmdExit1_Click()

Unload Me

End Sub


Private Sub cmdGenTables_Click()

  Dim pDoc As IMxDocument
  Dim pMap As IMap
  Set pDoc = ThisDocument
  Set pMap = pDoc.FocusMap
  
  
  '''Create New Folder'''
Dim SaveDirAdd As String
Dim fldr, i
Dim flg As Boolean

SaveDirAdd = pubSaveDir & "\nTables"

flg = False
Set fldr = CreateObject("Scripting.FileSystemObject")
fldr.CreateFolder (SaveDirAdd)
  
  
  ' Get the table (selected)
    
    Dim pTable As ITable
    Dim pLayer As ILayer
    Dim pOriFLayer As IFeatureLayer
    Dim LayerName As String
        
    LayerName = "RF0" '& pubDate
    Set pLayer = FindLayerByNameNGroup(pMap, LayerName, pubGrpLayerName)
    Set pOriFLayer = pLayer
    Set pTable = pOriFLayer.FeatureClass

    'Query filter to get row count
    Dim pQueryFilter As IQueryFilter
    Set pQueryFilter = New QueryFilter

    pQueryFilter.WhereClause = "Stn_ID is not null"

    Dim pRowCount As Integer
    pRowCount = pTable.RowCount(pQueryFilter)


    'find source layer by name                      
    Dim OriLayerName As String
    OriLayerName = "RF0" '& pubDate

    Dim pOriLayer As ILayer
    Set pOriLayer = FindLayerByNameNGroup(pMap, OriLayerName, pubGrpLayerName)


    'declare variables
    Dim pFLayer As IFeatureLayer
    Dim pFc As IFeatureClass
    Dim pInFeatureClassName As IFeatureClassName

    Dim pDataset As IDataset
    Dim pInDatasetName As IDatasetName

    Dim pFeatureClassName As IFeatureClassName
    Dim pOutDatasetName As IDatasetName
    Dim pWorkspaceName As IWorkspaceName

    Dim pExportOp As IExportOperation

    Set pFLayer = pOriLayer
    Set pFc = pFLayer.FeatureClass
    

    'Get the FcName from the featureclass
    Set pDataset = pFc
    Set pInFeatureClassName = pDataset.FullName
    Set pInDatasetName = pInFeatureClassName


'''Start row deleting loop'''

Dim il As Integer
For il = 0 To (pRowCount - 1)
        

    'Define the output feature class name
    Set pFeatureClassName = New FeatureClassName
    Set pOutDatasetName = pFeatureClassName
    pOutDatasetName.Name = "RF" & (il + 1)
        
    'set output file properties    
    Set pWorkspaceName = New WorkspaceName
    pWorkspaceName.PathName = SaveDirAdd
    pWorkspaceName.WorkspaceFactoryProgID = "esriDataSourcesFile.shapefileworkspacefactory.1"
    
    Set pOutDatasetName.WorkspaceName = pWorkspaceName

    pFeatureClassName.FeatureType = esriFTSimple
    pFeatureClassName.ShapeType = esriGeometryAny
    pFeatureClassName.ShapeFieldName = "Shape"
    
    'query filter for export operation
    Dim pQFilter As IQueryFilter
    Set pQFilter = New QueryFilter
    pQFilter.SubFields = "*"

    'Export
    Set pExportOp = New ExportOperation
    pExportOp.ExportFeatureClass pInDatasetName, pQFilter, Nothing, Nothing, pOutDatasetName, 0

    'Create a new feature layer and assign shapefile to it
    Dim pFeatureWorkspace As IFeatureWorkspace
    Set pFeatureWorkspace = pDataset.Workspace
  
    Dim pNewFLayer As IFeatureLayer
    Set pNewFLayer = New FeatureLayer
  
    Dim pWorkspaceFactory As IWorkspaceFactory
    Set pWorkspaceFactory = New ShapefileWorkspaceFactory
  
    Dim ShpName As String
    ShpName = "RF" & (il + 1) & ".shp"
  
    Set pFeatureWorkspace = pWorkspaceFactory.OpenFromFile(SaveDirAdd, Application.hWnd)
    Set pNewFLayer.FeatureClass = pFeatureWorkspace.OpenFeatureClass(ShpName)
    
    
    '''delete 1 row from each table'''
    Dim pEditLayer As IFeatureLayer
    Set pEditLayer = pNewFLayer
  
    'start editing and delete queried row
    Dim pNewTable As ITable
    Dim pEditDataSet As IDataset
    Dim pWorkspaceEdit As IWorkspaceEdit

    Set pEditDataSet = pEditLayer.FeatureClass

    Set pWorkspaceEdit = pEditDataSet.Workspace
    Set pNewTable = pEditLayer.FeatureClass


    'get row to be deleted
    pWorkspaceEdit.StartEditing (False)

        Dim pRow As IRow
        Set pRow = pNewTable.GetRow(il)
        pRow.Delete
  
    pWorkspaceEdit.StopEditing (True)


  'Add layer to Group Layer in Map
  Dim pFinalLayer As ILayer
  Dim pFinalFLayer As IFeatureLayer
  Set pFinalFLayer = New FeatureLayer
  Set pFinalFLayer.FeatureClass = pNewTable
  Set pFinalLayer = pFinalFLayer
  pFinalLayer.Name = "RF" & (il + 1)
  
  Dim pMapLayers As IMapLayers
  Set pMapLayers = pMap
  
  Dim pGrpLayer As IGroupLayer
  Set pGrpLayer = FindLayerByName(pMap, pubGrpLayerName)
    
  pMapLayers.InsertLayerInGroup pGrpLayer, pFinalLayer, True, 0
  pDoc.ActivatedView.PartialRefresh esriViewGeography, Nothing, Nothing


    'clear variables
    Set pNewFLayer = Nothing
    Set pFinalLayer = Nothing
    Set pEditLayer = Nothing
    Set pNewTable = Nothing
    Set pRow = Nothing

Next il


'''prevent button from being clicked twice
Me.cmdGenTables.Enabled = False


End Sub


Private Sub cmdNewFolder_Click()

On Error GoTo ErrorHandler


'''create new Group Layer'''
Dim pMxDoc As IMxDocument
Set pMxDoc = ThisDocument

Dim pGrpLyr As IGroupLayer
Set pGrpLyr = New GroupLayer


'Group Layer Name
pGrpLyr.Name = "RF" & pubDate
pubGrpLayerName = pGrpLyr.Name

pMxDoc.AddLayer pGrpLyr
pMxDoc.UpdateContents


'''Create New Folder'''
Dim fldr, i
Dim flg As Boolean

pubSaveDir = frmRainfallAnalysis.txtNewFolder.Text

flg = False
Set fldr = CreateObject("Scripting.FileSystemObject")
fldr.CreateFolder (pubSaveDir)


ErrorHandler:

If Err.Number > 0 Then
  MsgBox Err.Description
Else: MsgBox "New folder created"
End If


End Sub


Private Sub cmdNext1_Click()

Unload Me
'Me.Hide
frmKrigeVariogram.Show

End Sub


Private Sub cmdRGO_Click()

''Labels visibilities
    lblRGO2.Visible = True
    lblRGO3.Visible = True

''textbox visibilities
If cboRGONo = "0" Then

    lblRGO2.Visible = False
    lblRGO3.Visible = False

    txtRGO1.Visible = False
    txtRGO2.Visible = False
    txtRGO3.Visible = False
    txtRGO4.Visible = False

ElseIf cboRGONo = "1" Then
    txtRGO1.Visible = True
    txtRGO2.Visible = False
    txtRGO3.Visible = False
    txtRGO4.Visible = False

ElseIf cboRGONo = "2" Then
    txtRGO1.Visible = True
    txtRGO2.Visible = True
    txtRGO3.Visible = False
    txtRGO4.Visible = False
    
ElseIf cboRGONo = "3" Then
    txtRGO1.Visible = True
    txtRGO2.Visible = True
    txtRGO3.Visible = True
    txtRGO4.Visible = False
    
ElseIf cboRGONo = "4" Then
    txtRGO1.Visible = True
    txtRGO2.Visible = True
    txtRGO3.Visible = True
    txtRGO4.Visible = True
    
End If


End Sub


Private Sub cmdXYEvent_Click()

  Dim pDoc As IMxDocument
  Dim pMap As IMap
  Set pDoc = ThisDocument
  Set pMap = pDoc.FocusMap
    
  
  'Get table from Access database named "Daily_RF_Depth_Corrected"    
  Dim pStTabCol As IStandaloneTableCollection
  Dim pStandaloneTable As IStandaloneTable
  Dim intCount As Integer
  Dim pTable As ITable
  Set pStTabCol = pMap
  For intCount = 0 To pStTabCol.StandaloneTableCount - 1
    Set pStandaloneTable = pStTabCol.StandaloneTable(intCount)
    If pStandaloneTable.Name = "Daily_RF_Depth_Corrected" Then  'table name may change
      Set pTable = pStandaloneTable.Table
      Exit For
    End If
  Next
  If pTable Is Nothing Then
    MsgBox "The table was not found"
    Exit Sub
  End If
      
  
  ' Get the table name object
  Dim pDataset As IDataset
  Dim pTableName As IName
  Set pDataset = pTable
  Set pTableName = pDataset.FullName

  
  ' Specify the X and Y fields
  Dim pXYEvent2FieldsProperties As IXYEvent2FieldsProperties
  Set pXYEvent2FieldsProperties = New XYEvent2FieldsProperties
  With pXYEvent2FieldsProperties
    .XFieldName = "Longitude"    'Field Name changes
    .YFieldName = "Latitude"
    .ZFieldName = ""
  End With


  ' Specify the projection
  Dim pSpatialReferenceFactory As ISpatialReferenceFactory
  Dim pGeographicCoordinateSystem As IGeographicCoordinateSystem
  Set pSpatialReferenceFactory = New SpatialReferenceEnvironment
  Set pGeographicCoordinateSystem = pSpatialReferenceFactory.CreateGeographicCoordinateSystem(esriSRGeoCS_WGS1984)


  ' Create the XY name object and set it's properties
  Dim pXYEventSourceName As IXYEventSourceName
  Dim pXYName As IName
  Dim pXYEventSource As IXYEventSource
  
  Set pXYEventSourceName = New XYEventSourceName
  With pXYEventSourceName
    Set .EventProperties = pXYEvent2FieldsProperties
    Set .SpatialReference = pGeographicCoordinateSystem
    Set .EventTableName = pTableName
  End With
  Set pXYName = pXYEventSourceName
  Set pXYEventSource = pXYName.Open
           
  
  ' Create a new Map Layer
  Dim pFLayer As IFeatureLayer
  Set pFLayer = New FeatureLayer
  Set pFLayer.FeatureClass = pXYEventSource
  pFLayer.Name = "RF" & pubDate & " XY Event Layer"
        
  
  'Add the layer extension (this is done so that when you edit
  'the layer's Source properties and click the Set Data Source
  'button, the Add XY Events Dialog appears)
  Dim pLayerExt As ILayerExtensions
  Dim pRESPageExt As New XYDataSourcePageExtension
  Set pLayerExt = pFLayer
  pLayerExt.AddExtension pRESPageExt
    
  
  'Add layer to Group Layer in Map  
  Dim pMapLayers As IMapLayers
  Set pMapLayers = pMap
  
  Dim pGrpLayer As IGroupLayer
  Set pGrpLayer = FindLayerByName(pMap, pubGrpLayerName)
    
  pMapLayers.InsertLayerInGroup pGrpLayer, pFLayer, True, 0
  pDoc.ActivatedView.PartialRefresh esriViewGeography, Nothing, Nothing
    

'''prevent button from being clicked twice
cmdXYEvent.Enabled = False

  
End Sub


Private Sub cmdXYtoShape_Click()

'declare variables
    Dim pDoc As IMxDocument
    Dim pMap As IMap
    Dim pFLayer As IFeatureLayer
    Dim pFc As IFeatureClass
    Dim pInFeatureClassName As IFeatureClassName

    Dim pDataset As IDataset
    Dim pInDSNAme As IDatasetName

    Dim pFSel As IFeatureSelection
    Dim pSelSet As ISelectionSet

    Dim pFeatureClassName As IFeatureClassName
    Dim pOutDatasetName As IDatasetName
    Dim pWorkspaceName As IWorkspaceName

    Dim pExportOp As IExportOperation

    
    'find xy layer to convert (inside group layer)    
    Set pDoc = ThisDocument
    Set pMap = pDoc.FocusMap
    
    Dim XYLyrName As String
    XYLyrName = "RF" & pubDate & " XY Event Layer"
    
    Set pFLayer = FindLayerByNameNGroup(pMap, XYLyrName, pubGrpLayerName)
    Set pFc = pFLayer.FeatureClass


    'Get the FcName from the featureclass
    Set pDataset = pFc
    Set pInFeatureClassName = pDataset.FullName
    Set pInDSNAme = pInFeatureClassName
    

    'Get the selection set
    Set pFSel = pFLayer
    Set pSelSet = pFSel.SelectionSet


    'Define the output feature class name
    Set pFeatureClassName = New FeatureClassName
    Set pOutDatasetName = pFeatureClassName

    pOutDatasetName.Name = "RF0" 
    
    Set pWorkspaceName = New WorkspaceName
    pWorkspaceName.PathName = pubSaveDir
    pWorkspaceName.WorkspaceFactoryProgID = "esriDataSourcesFile.shapefileworkspacefactory.1"

    Set pOutDatasetName.WorkspaceName = pWorkspaceName

    pFeatureClassName.FeatureType = esriFTSimple
    pFeatureClassName.ShapeType = esriGeometryAny
    pFeatureClassName.ShapeFieldName = "Shape"
    

    'Export
    Set pExportOp = New ExportOperation
    pExportOp.ExportFeatureClass pInDSNAme, Nothing, pSelSet, Nothing, pOutDatasetName, 0


'Create a new feature layer and assign shapefile to it
  Dim pFeatureWorkspace As IFeatureWorkspace
  Set pFeatureWorkspace = pDataset.Workspace
  
  Dim pNewFLayer As IFeatureLayer
  Set pNewFLayer = New FeatureLayer
  
  Dim pWorkspaceFactory As IWorkspaceFactory
  Set pWorkspaceFactory = New ShapefileWorkspaceFactory
  
  Dim XYtoShpName As String
  XYtoShpName = "RF0.shp"  
  
  Set pFeatureWorkspace = pWorkspaceFactory.OpenFromFile(pubSaveDir, Application.hWnd)
  Set pNewFLayer.FeatureClass = pFeatureWorkspace.OpenFeatureClass(XYtoShpName)

  pNewFLayer.Name = "RF0" 


  'Add layer to Group Layer in Map  
  Dim pMapLayers As IMapLayers
  Set pMapLayers = pMap
  
  Dim pGrpLayer As IGroupLayer
  Set pGrpLayer = FindLayerByName(pMap, pubGrpLayerName)
    
  pMapLayers.InsertLayerInGroup pGrpLayer, pNewFLayer, True, 0
  pDoc.UpdateContents


'''prevent button from being clicked twice
Me.cmdXYtoShape.Enabled = False


End Sub


Private Sub txtDate_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

'make sure textbox accept numbers only
Dim str As String

str = ".0123456789"
If KeyAscii > 26 Then
   If InStr(str, Chr(KeyAscii)) = 0 Then
       KeyAscii = 0
   End If
End If

End Sub


Private Sub txtRGO1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

'make sure textbox accept numbers only
Dim str As String

str = ".0123456789"
If KeyAscii > 26 Then
   If InStr(str, Chr(KeyAscii)) = 0 Then
       KeyAscii = 0
   End If
End If

End Sub


Private Sub txtRGO2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

'make sure textbox accept numbers only
Dim str As String

str = ".0123456789"
If KeyAscii > 26 Then
   If InStr(str, Chr(KeyAscii)) = 0 Then
       KeyAscii = 0
   End If
End If

End Sub


Private Sub txtRGO3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

'make sure textbox accept numbers only
Dim str As String

str = ".0123456789"
If KeyAscii > 26 Then
   If InStr(str, Chr(KeyAscii)) = 0 Then
       KeyAscii = 0
   End If
End If

End Sub


Private Sub txtRGO4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

'make sure textbox accept numbers only
Dim str As String

str = ".0123456789"
If KeyAscii > 26 Then
   If InStr(str, Chr(KeyAscii)) = 0 Then
       KeyAscii = 0
   End If
End If

End Sub


Private Sub UserForm_Activate()
    
''get document file path - sFolder
    
    'declare application
    Dim pApp As IApplication
    Set pApp = New AppRef

    ' Create a FileSystemObject
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
  
    ' Save file name portion of file path
    Dim filepath As String
    filepath = GetDocPath(pApp)
    Dim sFolder As String
    sFolder = fso.GetParentFolderName(filepath)


frmRainfallAnalysis.txtDate.HideSelection = False
frmRainfallAnalysis.txtDate.SelStart = 0
frmRainfallAnalysis.txtDate.SelLength = Len(frmRainfallAnalysis.txtDate.Text)
frmRainfallAnalysis.txtNewFolder.Text = sFolder & "\"


''RG Omission combo box
Dim rgo As Integer

For rgo = 0 To 4
    cboRGONo.AddItem rgo
Next rgo

cboRGONo.Text = "0"

End Sub
