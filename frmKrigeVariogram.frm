Private Sub cmdBack1_Click()

Me.Hide
frmRainfallAnalysis.Show

End Sub


Private Sub cmdExit2_Click()

Unload Me

End Sub



Private Sub cmdInterpolateAll_Click()

    'Get the focused map from MapDocument
    Dim pDoc As IMxDocument
    Dim pMap As IMap
    Set pDoc = ThisDocument
    Set pMap = pDoc.FocusMap
        
        
  ' Get table with n rows    
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
        
        
'statusbar progress message
Application.StatusBar.Message(0) = "Interpolating... Please wait..."
        
        
'start loop for n tables
Dim il As Integer
For il = 0 To pRowCount
        
    'Get the feature class from layer name in Group Layer
    Dim sName As String
    sName = "RF" & il
    Dim pInLayer As ILayer
    Set pInLayer = FindLayerByNameNGroup(pMap, sName, pubGrpLayerName)
        
    Dim pFLayer As IFeatureLayer
    Set pFLayer = pInLayer
    
    Dim pFClass As IFeatureClass
    Set pFClass = pFLayer.FeatureClass
    
    'Specify the fieldname
    Dim sFieldName As String
    sFieldName = "Avg_RF_Dep"   'field name changes

    'Create FeatureClassDescriptor using a value field
    Dim pFCDescriptor As IFeatureClassDescriptor
    Set pFCDescriptor = New FeatureClassDescriptor
    pFCDescriptor.Create pFClass, Nothing, sFieldName
 
    'Create the Semi-variogram
    Dim pSemiVariogram As IGeoAnalysisSemiVariogram
    Set pSemiVariogram = New GeoAnalysisSemiVariogram
    
    'Define and set the parameters
    Dim dRange As Double
    Dim dSill As Double
    Dim dNugget As Double
    Dim dLag As Double
    Dim dRadius As Double
    Dim VarModel As esriGeoAnalysisSemiVariogramEnum
    
    dRange = Me.txtRange.Value
    dSill = Me.txtSill.Value
    dNugget = Me.txtNugget.Value
    dLag = Me.txtLagSize.Value
    
    
    'choose constant value from combo box
    If Me.cboSemiVarModel.Text = "esriGeoAnalysisNoneVariogram" Then
        VarModel = 1
    ElseIf Me.cboSemiVarModel.Text = "esriGeoAnalysisSphericalSemiVariogram" Then
        VarModel = 2
    ElseIf Me.cboSemiVarModel.Text = "esriGeoAnalysisCircularSemiVariogram" Then
        VarModel = 3
    ElseIf Me.cboSemiVarModel.Text = "esriGeoAnalysisExponentialSemiVariogram" Then
        VarModel = 4
    ElseIf Me.cboSemiVarModel.Text = "esriGeoAnalysisGaussianSemiVariogram" Then
        VarModel = 5
    ElseIf Me.cboSemiVarModel.Text = "esriGeoAnalysisLinearSemiVariogram" Then
        VarModel = 6
    ElseIf Me.cboSemiVarModel.Text = "esriGeoAnalysisUniversal1SemiVariogram" Then
        VarModel = 7
    ElseIf Me.cboSemiVarModel.Text = "esriGeoAnalysisUniversal2SemiVariogram" Then
        VarModel = 8
    End If
         
    
    pSemiVariogram.Lag = dLag
    pSemiVariogram.DefineVariogram VarModel, dRange, dSill, dNugget
       
    'Create radius using variable distance
    Dim pRadius As IRasterRadius
    Set pRadius = New RasterRadius
    pRadius.SetVariable Me.txtRadiusPts.Text
        
    'Create a RasterInterpolationOp object
    Dim pInterpolationOp As IInterpolationOp
    Set pInterpolationOp = New RasterInterpolationOp
  
    'Create Raster Analysis Environment
    Dim pEnv As IRasterAnalysisEnvironment
    Set pEnv = pInterpolationOp
    
    'set output cell size
    Dim CellSize As Double
    CellSize = Me.txtCellSize.Value
    pEnv.SetCellSize esriRasterEnvValue, CellSize

    'Perform kriging interpolation using the semi-variogram method
    Dim pOutRaster As IRaster
    Set pOutRaster = pInterpolationOp.Variogram(pFCDescriptor, pSemiVariogram, pRadius, False)

    'Add output into ArcMap as a raster layer
    Dim pOutRasLayer As IRasterLayer
    Set pOutRasLayer = New RasterLayer
    
    'set raster to be first band of the raster
    Dim pRasterDS As IRasterDataset
    Dim pRasBandCol As IRasterBandCollection

    Set pRasBandCol = pOutRaster
    Set pRasterDS = pRasBandCol.Item(0).RasterDataset

    'Create a new raster layer and save in folder  
    Dim pWorkspaceFactory As IWorkspaceFactory
    Dim pRasterWorkspace As IRasterWorkspace
    Dim pTempDS As ITemporaryDataset
    Dim pDataset As IDataset
    Dim OutputRasterName As String
  
    OutputRasterName = "RF" & il
  
    Set pWorkspaceFactory = New RasterWorkspaceFactory
    Set pRasterWorkspace = pWorkspaceFactory.OpenFromFile(pubSaveDir, Application.hWnd)   'raster directory
  
    Set pDataset = pRasterDS
    Set pTempDS = pRasterDS
    Set pRasterDS = pTempDS.MakePermanentAs(OutputRasterName, pRasterWorkspace, "GRID")
        
    pOutRasLayer.CreateFromDataset pRasterDS
    pOutRasLayer.Name = "Var" & OutputRasterName


    'Add layer to Group Layer in Map
    Dim pFinalLayer As ILayer
    Set pFinalLayer = pOutRasLayer
  
    Dim pMapLayers As IMapLayers
    Set pMapLayers = pMap
  
    Dim pGrpLayer As IGroupLayer
    Set pGrpLayer = FindLayerByName(pMap, pubGrpLayerName)
    
    pMapLayers.InsertLayerInGroup pGrpLayer, pFinalLayer, True, 0


    'clear variables
    Set pRasterDS = Nothing
    Set pOutRasLayer = Nothing

Next il


pDoc.ActivatedView.PartialRefresh esriViewGeography, Nothing, Nothing

'''prevent button from being clicked twice
Me.cmdInterpolateAll.Enabled = False

End Sub


Private Sub cmdNext2_Click()

Unload Me
frmSample.Hide
frmSample.Show

End Sub


Private Sub cmdOpenInput_Click()

''dialog box to open file
    Dim pGxDialog As IGxDialog
    Set pGxDialog = New GxDialog

    Dim pFilterCol As IGxObjectFilterCollection
    Set pFilterCol = pGxDialog

    Dim pEnumGx As IEnumGxObject

    'declare filter types
    Dim pFilter As IGxObjectFilter
    Set pFilter = New GxFilterTextFiles

    'add filter
    pFilterCol.AddFilter pFilter, True    'the default filter

    'dialog box properties
    pGxDialog.Title = "Insert Input Values"
    pGxDialog.RememberLocation = True
    pGxDialog.DoModalOpen 0, pEnumGx

    Dim pAddFile As IGxObject
    Set pAddFile = pEnumGx.Next
    
    'handle error
    If pAddFile Is Nothing Then
        Exit Sub
        On Error Resume Next
    Else
    End If


''open text files as input
Dim DataFile As String
Dim SemiVarModel As String, RadiusPts As String, CellSize As String
Dim LagSize As String, Range As String, Sill As String, Nugget As String

DataFile = pAddFile.FullName

'open and insert values from text files

Open DataFile For Input As #1

Input #1, SemiVarModel, RadiusPts, CellSize, LagSize, Range, Sill, Nugget

    Me.cboSemiVarModel.Text = SemiVarModel
    Me.txtRadiusPts.Text = RadiusPts
    Me.txtCellSize.Text = CellSize
    Me.txtLagSize.Text = LagSize
    Me.txtRange.Text = Range
    Me.txtSill.Text = Sill
    Me.txtNugget.Text = Nugget

Close #1

End Sub


Private Sub cmdSaveInput_Click()

''dialog box to save file (as)
    Dim pGxDialog As IGxDialog
    Set pGxDialog = New GxDialog

    Dim pFilterCol As IGxObjectFilterCollection
    Set pFilterCol = pGxDialog

    Dim pEnumGx As IEnumGxObject

    'declare filter types
    Dim pFilter As IGxObjectFilter
    Set pFilter = New GxFilterTextFiles

    'add filter
    pFilterCol.AddFilter pFilter, True    'the default filter

    'dialog box properties
    Dim sLocation As String
    pGxDialog.Title = "Save Input Values As"
    pGxDialog.RememberLocation = True
    pGxDialog.StartingLocation = sLocation
    pGxDialog.DoModalSave hParentHwnd

    Dim pSaveFile As IGxObject
    Set pSaveFile = pGxDialog.FinalLocation
    sLocation = pSaveFile.FullName

    ''save input to text file
    Dim DataFile As String
    DataFile = sLocation & "\" & pGxDialog.Name & ".txt"

        Open DataFile For Output As #1
        Write #1, Me.cboSemiVarModel.Text, Me.txtRadiusPts.Text, Me.txtCellSize.Text, _
                        Me.txtLagSize.Text, Me.txtRange.Text, Me.txtSill.Text, Me.txtNugget.Text
        Close #1
 
End Sub


Private Sub UserForm_Initialize()

Me.cboSemiVarModel.Clear

'set combo box property so that user cannot enter values
Me.cboSemiVarModel.Style = fmStyleDropDownList

'add items to combo box's drop-down menu
Me.cboSemiVarModel.AddItem "esriGeoAnalysisExponentialSemiVariogram"
Me.cboSemiVarModel.AddItem "esriGeoAnalysisGaussianSemiVariogram"
Me.cboSemiVarModel.AddItem "esriGeoAnalysisNoneVariogram"
Me.cboSemiVarModel.AddItem "esriGeoAnalysisSphericalSemiVariogram"
Me.cboSemiVarModel.AddItem "esriGeoAnalysisCircularSemiVariogram"
Me.cboSemiVarModel.AddItem "esriGeoAnalysisLinearSemiVariogram"
Me.cboSemiVarModel.AddItem "esriGeoAnalysisUniversal1SemiVariogram"
Me.cboSemiVarModel.AddItem "esriGeoAnalysisUniversal2SemiVariogram"

End Sub
