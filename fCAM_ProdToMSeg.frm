VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fCAM_ProdToMSeg 
   Caption         =   "Product Allocation"
   ClientHeight    =   10875
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4905
   OleObjectBlob   =   "fCAM_ProdToMSeg.frx":0000
   StartUpPosition =   1  'CenterOwner
   Tag             =   "fCAM_ProdToMSeg"
End
Attribute VB_Name = "fCAM_ProdToMSeg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit         ' fCam_ProdToMSeg
' Used to allocated ProductCodes to the Market segments.
Private peCurSegMethod As e_MSegType
Private pbIsActive As Boolean
Private plDoc_ID As Long
Private Sub UserForm_Terminate()
    mCAM_Runtime.CategoryObject(lDoc_ID).cPGrp.SaveMSegAllocationtoDB
End Sub

Private Sub cbo_AllocateTo_Change()
    ' Sort out what allocation is to be performed-Disables and Enables the allocate Button
    If cbo_AllocateTo = "" Then cmd_Allocate.Enabled = False Else cmd_Allocate.Enabled = True
End Sub
Private Sub cbo_MSegSelect_Change()
    ' When the cbobox is changed the listbox is updated with products from Category Object
Dim col As Collection
Dim P As cCBA_Prod
Dim CurMSeg As String
Dim v As Variant
    If cbo_MSegSelect = "" Then Exit Sub
    Me.lst_Products.Clear
    If Me.cbo_MSegSelect = "Unassigned" Then CurMSeg = "" Else CurMSeg = cbo_MSegSelect.Value
    Set col = mCAM_Runtime.CategoryObject(lDoc_ID).cPGrp.getProdListing
    For Each v In col
        Set P = mCAM_Runtime.CategoryObject(lDoc_ID).cPGrp.getProdObject(v)
        If eCurSegMethod = eScanData Then
            If P.sScanDataMSeg = CurMSeg Or CurMSeg = "ALL" Then Me.lst_Products.AddItem P.lPcode & "-" & P.sProdDesc
        Else
            If P.ManualMSeg(1) = CurMSeg Or CurMSeg = "ALL" Then Me.lst_Products.AddItem P.lPcode & "-" & P.sProdDesc
        End If
    Next
End Sub
Private Sub cbo_SegMethod_Change()
    ' Dependant upon selection MSeg Method, runs MSeg to combobox function
    If cbo_SegMethod = "" Then peCurSegMethod = eNone
    If cbo_SegMethod = "Manual" Then peCurSegMethod = eManual1
    If cbo_SegMethod = "ScanData" Then peCurSegMethod = eScanData
    If cbo_SegMethod = "HomeScan" Then peCurSegMethod = eHomescan
    PopulateMSegsToComboBoxes
End Sub
Private Sub cmd_Allocate_Click()
    ' runs setProductAllocations
    SetProductAllocations
End Sub
Private Sub cmd_ApplyToMaster_Click()
    ' Sends the current setup to the respective cCAMERA_Category and the CBA_ProdGroups / CBA_Prod Objects ' @RWCam - to be added
    Stop
End Sub
Private Sub cmd_SelAll_Click()
    ' Selects all the values in the Listbox
    Dim a As Long
    For a = 0 To lst_Products.ListCount - 1
        lst_Products.Selected(a) = True
    Next
End Sub
Private Sub UserForm_Initialize()
    ' Cannot be opened unless a category is selected in the ribbon
    ' Sets 3 segmentation methods into cbo_MSegMethod (Hard Coded). Set isActive = True
    Dim Doc As cCBA_Document
    Me.StartUpPosition = 0
    Me.Top = Application.Top + (Application.Height / 2) - (Me.Height / 2)
    Me.Left = Application.Left + (Application.Width / 2) - (Me.Width / 2)
    
    If mCAM_Runtime.ActiveDO.eDocumentType = eDocuNone Then
        Me.lbl_Identifier = "Active Segmentation"
        lDoc_ID = 0
    Else
        lDoc_ID = mCAM_Runtime.ActiveDO.lDoc_ID
        Me.Width = 513.75
        Set Doc = mCAM_Runtime.cDocHolder.sdDocumentList(lDoc_ID)
        Me.lbl_Identifier.Caption = Doc.sDocumentName
    End If
    cbo_SegMethod.AddItem "Manual"
    cbo_SegMethod.AddItem "ScanData"
    cbo_SegMethod.AddItem "Homescan"
    bIsActive = True
End Sub
Private Function PopulateMSegsToComboBoxes() As Boolean
Dim tempval As String
Dim allocateval As String
    ' Updates the combo-boxes with the respective marketsegments
    tempval = Me.cbo_MSegSelect.Value
    allocateval = Me.cbo_AllocateTo.Value
    Me.cbo_MSegSelect.Clear
    Me.cbo_MSegSelect.List = mCAM_Runtime.CategoryObject(lDoc_ID).GetAllSelMSegNames(eCurSegMethod)
    Me.cbo_AllocateTo.List = Me.cbo_MSegSelect.List
    Me.cbo_MSegSelect.AddItem "Unassigned", 0
    Me.cbo_MSegSelect.AddItem "ALL", 1
    If tempval <> "" Then Me.cbo_MSegSelect.Value = tempval
    If allocateval <> "" Then Me.cbo_AllocateTo.Value = allocateval
    If Err.Number > 0 Then PopulateMSegsToComboBoxes = False Else PopulateMSegsToComboBoxes = True
End Function
Public Function AccomodateMSegStructureChanged() As Boolean
    AccomodateMSegStructureChanged = PopulateMSegsToComboBoxes
    ClearAssignmentsToDeletedMSeg
End Function
Private Function ClearAssignmentsToDeletedMSeg() As Boolean
Dim col As Collection
Dim MSegCol As Collection
Dim v As Variant, va As Variant
Dim P As cCBA_Prod
Dim bfound As Boolean
    bfound = False
    For Each v In Me.cbo_AllocateTo.List
        If Me.cbo_AllocateTo.Value = v Then bfound = True: Exit For
    Next
    If bfound = False Then Me.cbo_AllocateTo.Value = ""
    bfound = False
    For Each v In Me.cbo_MSegSelect.List
        If Me.cbo_MSegSelect.Value = v Then bfound = True: Exit For
    Next
    If bfound = False Then Me.cbo_MSegSelect.Value = "Unassigned"
    ClearAssignmentsToDeletedMSeg = True
End Function
Private Sub SetProductAllocations()
    ' If cbo_segto <> "" and there are selections in the listbox then:  sets the respective product allocation from the listbox (through the CAMERA_Runtime)
Dim a As Long
    For a = 0 To lst_Products.ListCount - 1
        If lst_Products.Selected(a) = True Then
            mCAM_Runtime.CategoryObject(lDoc_ID).cPGrp.setMSegAllocation Mid(lst_Products.List(a), 1, InStr(1, lst_Products.List(a), "-") - 1), IIf(eCurSegMethod = eScanData, True, False), Me.cbo_AllocateTo
        End If
    Next
    PopulateMSegsToComboBoxes
End Sub
Public Property Get bIsActive() As Boolean: bIsActive = pbIsActive: End Property
Private Property Let bIsActive(ByVal bNewValue As Boolean): pbIsActive = bNewValue: End Property
Private Property Get eCurSegMethod() As VBAProject.e_MSegType: eCurSegMethod = peCurSegMethod: End Property
Private Property Let eCurSegMethod(ByVal NewValue As VBAProject.e_MSegType): peCurSegMethod = NewValue: End Property
Public Property Get lDoc_ID() As Long: lDoc_ID = plDoc_ID: End Property
Public Property Let lDoc_ID(ByVal NewValue As Long): plDoc_ID = NewValue: End Property
