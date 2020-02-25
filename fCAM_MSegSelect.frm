VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fCAM_MSegSelect 
   ClientHeight    =   10785
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10035
   OleObjectBlob   =   "fCAM_MSegSelect.frx":0000
   StartUpPosition =   2  'CenterScreen
   Tag             =   "fCAM_MSegSelect"
End
Attribute VB_Name = "fCAM_MSegSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                         'fCam_MSegSelect
Private peCurSegType As e_MSegType
Private pbACG As Boolean
Private pbIsActive As Boolean
Private plDoc_ID As Long

' Allows Management of the MarketSegmentation methods.
' HomeScan is defaulted (As it correlates with CG/SCG setups), ScanData is Selected from Listbox, Manual is added via Textbox.
' Remember that initiation is always instigated by the Runtime, not the DocumentHolder

Private Sub cbo_MSegMethod_Change()
    ' if DocuType = none then Pull rqd data from runtime else pull data from DocumentHolder. dependent on selection, unhide & unhide rqd objects.
    Dim v As Variant, va As Variant, sdAD As Scripting.Dictionary, sdSD As Scripting.Dictionary, col As Collection
    Dim ND As cCBA_NielsenData
    Dim dic As Scripting.Dictionary
    
    Me.lst_SelMSeg.Clear
    Select Case cbo_MSegMethod.Value
        Case "Manual"
            eCurSegType = eManual1
            Me.Height = 245.25
            Me.lst_AvaliableMSeg.Clear
            Me.cmd_AddMSeg.Visible = True
            Me.cmd_AddtoSelected.Visible = False
            Me.lst_AvaliableMSeg.Visible = False
            Me.lst_SelMSeg.Visible = True
            Me.lbl_AvaliableMseg.Visible = False
            Me.lbl_SelMseg.Visible = True
            Me.txt_ManualSeg.Visible = True
            Me.cmd_RemoveSel.Visible = True
            Me.lst_SelMSeg.List = mCAM_Runtime.ActiveDO.GetAllSelMSegNames(eCurSegType)
        
        Case "ScanData"
            eCurSegType = eScanData
            Me.Height = 568.5
            Me.cmd_AddMSeg.Visible = False
            Me.cmd_AddtoSelected.Visible = True
            Me.lst_AvaliableMSeg.Visible = True
            Me.lst_SelMSeg.Visible = True
            Me.lbl_AvaliableMseg.Visible = True
            Me.lbl_SelMseg.Visible = True
            Me.txt_ManualSeg.Visible = False
            Me.cmd_RemoveSel.Visible = True
            Me.lst_AvaliableMSeg.List = mCAM_Runtime.cRibbonData.GetAllMSegNames(eCurSegType)
            Me.lst_SelMSeg.List = mCAM_Runtime.CategoryObject(lDoc_ID).GetAllSelMSegNames(eCurSegType)
            
        Case "HomeScan"
            eCurSegType = eHomescan
            Me.Height = 245.25
            Me.cmd_AddMSeg.Visible = False
            Me.cmd_AddtoSelected.Visible = False
            Me.lst_AvaliableMSeg.Visible = False
            Me.lst_SelMSeg.Visible = True
            Me.lbl_AvaliableMseg.Visible = False
            Me.lbl_SelMseg.Visible = True
            Me.txt_ManualSeg.Visible = False
            Me.cmd_RemoveSel.Visible = False
            Me.lst_SelMSeg.List = mCAM_Runtime.CategoryObject(lDoc_ID).GetAllSelMSegNames(eCurSegType)
    End Select

End Sub
Private Function ClearAssignmentsToDeletedMSeg() As Boolean
    ' Looks at all assignments and removes any where the segment is no longer active
    Dim col As Collection
    Dim MSegCol As Collection
    Dim v As Variant, va As Variant
    Dim P As cCBA_Prod
    Dim bfound As Boolean
    Set col = mCAM_Runtime.CategoryObject(lDoc_ID).cPGrp.getProdListing
    For Each v In col
        Set P = mCAM_Runtime.CategoryObject(lDoc_ID).cPGrp.getProdObject(v)
        If P.sScanDataMSeg <> "" Then
            bfound = False
            For Each va In mCAM_Runtime.CategoryObject(lDoc_ID).GetAllSelMSegNames(eCurSegType)
                If P.sScanDataMSeg = va Then bfound = True: Exit For
            Next
            If bfound = False Then P.sScanDataMSeg = ""
        End If
        If P.ManualMSeg(1) <> "" Then
            bfound = False
            For Each va In mCAM_Runtime.CategoryObject(lDoc_ID).GetAllSelMSegNames(eCurSegType)
                If P.ManualMSeg(1) = va Then bfound = True: Exit For
            Next
            If bfound = False Then P.ManualMSeg 1, 0
        End If
    Next
End Function
Private Sub cmd_AddMseg_Click()
    ' Add Text in txt_ManualSeg to lst_SegSelected, and add to ManualSegments dictionary in DataObject
    Dim sMData As String
    Dim ND As cCBA_NielsenData
    Dim UF As MSForms.UserForm
    Dim temp As Object
    
    If Me.txt_ManualSeg = "" Then Exit Sub
    Set ND = New cCBA_NielsenData
    ND.MSegDescription = Me.txt_ManualSeg.Value
    ND.Category = mCAM_Runtime.CategoryObject(lDoc_ID).sCategoryName
    ND.IsManual = True
    ND.SelectedForCategory = mCAM_Runtime.CategoryObject(lDoc_ID).sCategoryName
    mCAM_Runtime.cRibbonData.AmendMSeg True, ND
    mCAM_Runtime.CategoryObject(lDoc_ID).AmendMSeg True, ND
    Me.lst_SelMSeg.List = mCAM_Runtime.CategoryObject(lDoc_ID).GetAllSelMSegNames(eCurSegType)
    Me.txt_ManualSeg = ""
    
    
    Set temp = mCAM_Runtime.GetUserForm("fCam_ProdToMSeg", lDoc_ID)
    If Not temp Is Nothing Then
        If temp.bIsActive Then temp.AccomodateMSegStructureChanged
    End If

End Sub
Private Sub cmd_AddtoSelected_Click()       ' Called cmd_AddScanMSep_Click?? in document???
    ' Takes the ScanDataSegment in the lst_SegAvailable Listbox and adds it to the lst_SegSelect Listbox.
    ' Flags SelectedForCategory in NielsenData Object to be used to determine if the segment has already been used by another.
    ' If so, and continued then the previous allocation added to the previousSelectedForCategory to be populated with what is replaced.
    ' Instruction sent to the listbox to represent the Segment in Red color
    Dim a As Long
    Dim UF As MSForms.UserForm
    Dim temp As Object
    Dim ND As cCBA_NielsenData
    For a = 0 To Me.lst_AvaliableMSeg.ListCount - 1
        If Me.lst_AvaliableMSeg.Selected(a) Then
            Set ND = mCAM_Runtime.cRibbonData.GetMSeg(Me.lst_AvaliableMSeg.List(a), eCurSegType)
            If mCAM_Runtime.CategoryObject(lDoc_ID).AmendMSeg(True, ND) Then
                'Call cbo_MSegMethod_Change
            End If
        End If
    Next
    Me.lst_SelMSeg.List = mCAM_Runtime.CategoryObject(lDoc_ID).GetAllSelMSegNames(eCurSegType)
    
    Set temp = mCAM_Runtime.GetUserForm("fCam_ProdToMSeg", lDoc_ID)
    If Not temp Is Nothing Then
        If temp.bIsActive Then temp.AccomodateMSegStructureChanged
    End If
    
End Sub

Private Sub cmd_ApplyToMaster_Click()
    ' Will send the MSeg data to the database
    'BE VERY CAREFUL WHEN USING THIS FUNCTION AS YOU ARE ABOUT TO WRITE ALL THIS TO THE MASTER ALLOCATION!!
    If MsgBox("This will literally save what you have now to the Master. this cannot be undone, are you sure you want to do this?", vbYesNo) = vbYes Then
        Stop
        'PLEASE IF YOU ARE TESTING THIS, BE VERY CAREFUL TO NOT OVERWRITE ANY DATA IN THE DATABASE, USE F8
        mCAM_Runtime.CategoryObject(lDoc_ID).SaveSegmentsToDB
    End If
End Sub

Private Sub cmd_RemoveSel_Click()
    ' {curSegType=eScanData:Removes the selected ScanDataSegment from lst_SegSelect Listbox and adds it to the lst_SegAvailable Listbox}
    ' {curSegType=eManual: Removes the Segment from the lst_SegSelect Listboxand} - Sends changes to dataobject via CAMER_Runtime
    Dim a As Long
    Dim temp As Object
    If eCurSegType = eScanData Then
        For a = 0 To Me.lst_SelMSeg.ListCount - 1
            If Me.lst_SelMSeg.Selected(a) Then
                Call mCAM_Runtime.CategoryObject(lDoc_ID).AmendMSeg(False, mCAM_Runtime.cRibbonData.GetMSeg(Me.lst_SelMSeg.List(a), eCurSegType))
            End If
        Next
        Me.lst_SelMSeg.List = mCAM_Runtime.CategoryObject(lDoc_ID).GetAllSelMSegNames(eCurSegType)
    ElseIf eCurSegType = eManual1 Then
        For a = 0 To Me.lst_SelMSeg.ListCount - 1
            If Me.lst_SelMSeg.Selected(a) Then
                Call mCAM_Runtime.CategoryObject(lDoc_ID).AmendMSeg(False, mCAM_Runtime.cRibbonData.GetMSeg(Me.lst_SelMSeg.List(a), eCurSegType))
                Call mCAM_Runtime.cRibbonData.AmendMSeg(False, mCAM_Runtime.cRibbonData.GetMSeg(Me.lst_SelMSeg.List(a), eCurSegType))
            End If
        Next
        Me.lst_SelMSeg.List = mCAM_Runtime.CategoryObject(lDoc_ID).GetAllSelMSegNames(eCurSegType)
    ElseIf eCurSegType = eManual2 Then
        'TBC
    ElseIf eCurSegType = eManual3 Then
        'TBC
    End If
    
    ClearAssignmentsToDeletedMSeg
    Set temp = mCAM_Runtime.GetUserForm("fCam_ProdToMSeg", lDoc_ID)
    If Not temp Is Nothing Then
        If temp.bIsActive Then temp.AccomodateMSegStructureChanged
    End If

End Sub
Sub UserForm_Terminate()
    bIsActive = False
    If mCAM_Runtime.CategoryObject(lDoc_ID).SaveSegmentsToDB = False Then MsgBox "Error in saving Segments to Database"
End Sub
Sub UserForm_Initialize()
Dim Doc As cCBA_Document
    ' Make sure all objects other than the cbo_SegMethod is hidden.
    ' Set isActive = True.
    ' Segmentation types pulled from runtime.
    ' Populate lbl_Identifier with Master/DocumentName from the respective cCAMERA_Category Object either Ribbon or Document
    Dim CGSCG As Collection
    Me.StartUpPosition = 0
    Me.Top = Application.Top + (Application.Height / 2) - (Me.Height / 2)
    Me.Left = Application.Left + (Application.Width / 4) - (1.5 * (Me.Width / 4))
    Me.lblVersion.Caption = CBA_getVersionStatus(g_GetDB("Cam"), CBA_Cam_Ver, "Camera Tool", "Camera")  ' Get the latest version
    cbo_MSegMethod.AddItem "Manual"
    cbo_MSegMethod.AddItem "ScanData"
    cbo_MSegMethod.AddItem "HomeScan"
    If mCAM_Runtime.ActiveDO Is Nothing Then
        bIsActive = False
    Else
        Me.cmd_AddMSeg.Visible = False
        Me.cmd_AddtoSelected.Visible = False
        Me.lst_AvaliableMSeg.Visible = False
        Me.lst_SelMSeg.Visible = False
        Me.lbl_AvaliableMseg.Visible = False
        Me.lbl_SelMseg.Visible = False
        Me.txt_ManualSeg.Visible = False
        Me.cmd_RemoveSel.Visible = False
        Me.lst_ReportMethods = False
        Me.cmd_ApplyToMaster = False
        Me.lbl_ReportingMethods = False
        Me.Height = 245.25
        bACG = False
'        If mCAM_Runtime.ActiveDO.eDocumentType = eDocuNone Then
'            lDoc_ID = 0
'            Me.Width = 420
'            Me.lbl_Identifier.Caption = "Active Segmentation"
'        Else
'            lDoc_ID = mCAM_Runtime.ActiveDO.lDoc_ID
'            Me.Width = 513.75
'            Set doc = mCAM_Runtime.cDocHolder.sdDocumentList(lDoc_ID)
'            Me.lbl_Identifier.Caption = doc.sDocumentName
'        End If
'        bIsActive = True
    End If
End Sub
Private Property Get eCurSegType() As e_MSegType: eCurSegType = peCurSegType: End Property
Private Property Let eCurSegType(ByVal eNewValue As e_MSegType): peCurSegType = eNewValue: End Property
Private Property Get bACG() As Boolean: bACG = pbACG: End Property
Private Property Let bACG(ByVal bNewValue As Boolean): pbACG = bNewValue: End Property
Public Property Get bIsActive() As Boolean: bIsActive = pbIsActive: End Property
Private Property Let bIsActive(ByVal bNewValue As Boolean): pbIsActive = bNewValue: End Property
Public Property Get lDoc_ID() As Long: lDoc_ID = plDoc_ID: End Property
Public Property Let lDoc_ID(ByVal NewValue As Long)
    
    If plDoc_ID < 1 Then
        plDoc_ID = NewValue
        If NewValue = 0 Then
            Me.Width = 420: Me.lbl_Identifier.Caption = "Active Segmentation"
        Else
            Me.Width = 513.75
            Me.lbl_Identifier.Caption = mCAM_Runtime.cDocHolder.sdDocumentList(lDoc_ID).sDocumentName
        End If
        bIsActive = True
    Else
        plDoc_ID = NewValue
    End If
    
End Property
