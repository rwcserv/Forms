VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fCAM_LineCount 
   ClientHeight    =   10365
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12720
   OleObjectBlob   =   "fCAM_LineCount.frx":0000
   StartUpPosition =   2  'CenterScreen
   Tag             =   "fCAM_LineCount"
End
Attribute VB_Name = "fCAM_LineCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                                     ' fCam_LineCount
Private psCurProd As String                                         ' The ProductCode selected in the current Listbox
Private plCurPosit As Long                                          ' The index of the current ProductCode in the Listbox
''Private plRangeCount As Long                                        ' The number of items in the core range
Private plstCurLst As MSForms.ListBox                               ' The currently selected listbox
Private peCurPeriod As e_LineCountPeriod                            'The current Period Selected
Private psdPeriodNames As Scripting.Dictionary                      'A dictionary with all the periods
Private pbIsActive As Boolean                                       'A flag to tell the requestor that the userform is ready to be shown
Private psdOriginalLineCountDataObjects As Scripting.Dictionary     'A copy of the linecount data objects from when the userform is activated. This is taken only to allow for a Reset.
Private plDoc_ID As Long
    ' Allows Management of the Linecount for the category.
    ' It is basically a series of listboxes that have products in them.
    ' These products can then be reallocated to different listboxes.
    ' This is all managed in a dictionary which is passed into the userform and passed out when it is terminated.
    ' The CBA_Prod object has a function to determine what it the line designation should be, and that can be used to overlay any discrepancies.
Private Sub cbo_Period_Change()
    ' For the selected period, query the runtime for the line count data and pass to populateListBoxes
    If Me.cbo_Period.ListIndex = -1 Then Exit Sub
    eCurPeriod = CLng(sdPeriodNames(Me.cbo_Period.Value))
    Call RefreshAllListboxes
End Sub
Private Sub cmd_Reset_Click()
    ' Sets the selected line count to be equal to the initial line count' This cannot be.. what if you are loading a saved line count, change it, and then decide you updated it incorrectly? You need originals
    Call ResetLineCountObjectDic
    Call RefreshAllListboxes
End Sub
Private Function ResetLineCountObjectDic() As Boolean
    ' Reset all LineCount Objects to where they were when first loaded
    Dim v As Variant
    Dim LC As cCAM_LineCount_Data
    Dim dic As Scripting.Dictionary
'    Dim lLeft As Long, lTop As Long
    Dim ErrorInReset As Boolean
    ResetLineCountObjectDic = False
    If sdOriginalLineCountDataObjects Is Nothing Then MsgBox "Could not reset as no OriginalLineCountDataObjects available": Exit Function
    Set LC = sdOriginalLineCountDataObjects(eCurPeriod)
    Set dic = LC.sdProductAllocation
    For Each v In dic
        mCAM_Runtime.CategoryObject(lDoc_ID).AmendLineCountData eCurPeriod, CStr(v), dic(v)
        mCAM_Runtime.CategoryObject(lDoc_ID).SetLineCountObjectReset eCurPeriod
    Next
    If lDoc_ID > 0 Then Call mCAM_Runtime.ControlObject(lDoc_ID, , "RenderCell") ', lLeft, lTop)
    ResetLineCountObjectDic = True
End Function
Private Sub cmd_SendToMaster_Click()
    ' Sends the selected Line count to overwrite the ribbon selected line count for this period
    ' There needs to be some controls in place here.
    ' There needs to be some accomodation where a user may be working on a 2018 category review and that would update the 2020 YTD line count for example….' @RWCam ??????
    If MsgBox("This will literally save what you have now to the Master. this cannot be undone, are you sure you want to do this?", vbYesNo) = vbYes Then
        Stop
        'PLEASE IF YOU ARE TESTING THIS, BE VERY CAREFUL TO NOT OVERWRITE ANY DATA IN THE DATABASE, USE F8
        mCAM_Runtime.cActiveDataObject.SaveLineCountAllocation True
    End If
End Sub
Sub UserForm_Terminate()
    ' Send Line Count back to CATREV Main
    ' ISSUE HERE IS, IT ALWAYS ASKS YOU TO SAVE.. WHAT IF YOU DIDN'T CHANGE ANYTHING??
Dim v As Variant, va As Variant
Dim LC As cCAM_LineCount_Data
Dim bfound As Boolean
        If mCAM_Runtime.CategoryObject(lDoc_ID) Is Nothing Then Exit Sub
        For Each va In mCAM_Runtime.CategoryObject(lDoc_ID).sdLineCountDic
            bfound = False
            Set LC = mCAM_Runtime.CategoryObject(lDoc_ID).sdLineCountDic(va)
            If LC.IsChangedNotSaved = True Then bfound = True
            If bfound = True Then
                If MsgBox("Save?", vbYesNo) = vbYes Then
                    mCAM_Runtime.CategoryObject(lDoc_ID).SaveLineCountAllocation
                Else
'                    For Each v In mCAM_Runtime.CategoryObject(lDoc_ID).sdLineCountDic
'                        Set LC = mCAM_Runtime.CategoryObject(lDoc_ID).sdLineCountDic(v)
'                        If LC.IsChangedNotSaved = True Then eCurPeriod = LC.eLCPeriod: ResetLineCountObjectDic
'                    Next
                End If
            End If
        Next
End Sub

Sub UserForm_Initialize()
    Dim v As Variant
    ' Set position. Set lbl_Identifier. Set btn_SendToMaster invisible if lb_identifier is Master Data.
    Me.StartUpPosition = 0
    Me.Top = Application.Top + (Application.Height / 2) - (Me.Height / 2)
    Me.Left = Application.Left + (Application.Width / 2) - (Me.Width / 2)
    Me.lblVersion.Caption = CBA_getVersionStatus(g_GetDB("Cam"), CBA_Cam_Ver, "Camera Tool", "Camera")  ' Get the latest version
    Set sdPeriodNames = New Scripting.Dictionary
    sdPeriodNames.Add "MAT", e_LineCountPeriod.eMAT
    sdPeriodNames.Add "Prior MAT", e_LineCountPeriod.ePriorMAT
    sdPeriodNames.Add "YTD", e_LineCountPeriod.eYTD
    sdPeriodNames.Add "Prior YTD", e_LineCountPeriod.ePriorYTD
    sdPeriodNames.Add "QTRTD", e_LineCountPeriod.eQTRTD
    sdPeriodNames.Add "Prior QTR", e_LineCountPeriod.ePriorQTR
    sdPeriodNames.Add CStr(Year(Date) - 1), e_LineCountPeriod.ePriorCalendarYr
    sdPeriodNames.Add CStr(Year(Date) - 2), e_LineCountPeriod.ePrevPriorCalendarYr
    For Each v In sdPeriodNames
        Me.cbo_Period.AddItem CStr(v)
    Next

End Sub

Private Function DeSelAllListsButMe() As Boolean
    ' Deselects all products highlighted in other lists other than the currently selected
    Dim obj As Object, a As Long
    DeSelAllListsButMe = False
    For Each obj In Me.Controls
        If Mid(obj.Name, 1, 4) = "lst_" And obj.Name <> lstCurLst.Name Then
            For a = 0 To obj.ListCount - 1
                obj.Selected(a) = False
                DeSelAllListsButMe = True
            Next
        End If
    Next
End Function
Private Function RefreshAllListboxes() As Boolean
    ' Query the runtime for the line count data and pass to populateListBoxes
    RefreshAllListboxes = True
''    lRangeCount = 0
    If PopulateListBoxes(lst_Brand) = False Then RefreshAllListboxes = False
    If PopulateListBoxes(lst_Core) = False Then RefreshAllListboxes = False
    If PopulateListBoxes(lst_Delete) = False Then RefreshAllListboxes = False
    If PopulateListBoxes(lst_Region) = False Then RefreshAllListboxes = False
    If PopulateListBoxes(lst_Sea) = False Then RefreshAllListboxes = False
    If PopulateListBoxes(lst_Spec) = False Then RefreshAllListboxes = False
    If PopulateListBoxes(lst_TrialCur) = False Then RefreshAllListboxes = False
    If PopulateListBoxes(lst_TrialSuc) = False Then RefreshAllListboxes = False
    If RefreshAllListboxes = False Then MsgBox "Error in RefreshAllListboxes", vbOKOnly
    'Me.lbl_Identifier.Caption = "Current Default - Core Range=" & lRangeCount & " items"
End Function
Private Function PopulateListBoxes(ByVal FormName As MSForms.ListBox) As Boolean
    ' Adds the product values to the various listboxes
    
    Dim lCount As Long
    PopulateListBoxes = False
    If mCAM_Runtime.CategoryObject(lDoc_ID) Is Nothing Then PopulateListBoxes = True: Exit Function
    Select Case FormName.Name
        Case "lst_Brand"
            FormName.List = mCAM_Runtime.CategoryObject(lDoc_ID).PullLineCountAllocation(eCurPeriod, e_LineCountType.eBranded)
            GoSub GSCount
            Me.lbl_Brand.Caption = "BRANDED (" & lCount & ")"
            PopulateListBoxes = True
        Case "lst_Core"
            FormName.List = mCAM_Runtime.CategoryObject(lDoc_ID).PullLineCountAllocation(eCurPeriod, e_LineCountType.eCore)
            GoSub GSCount
            Me.lbl_Core.Caption = "CORE RANGE (" & lCount & ")"
            PopulateListBoxes = True
        Case "lst_Delete"
            FormName.List = mCAM_Runtime.CategoryObject(lDoc_ID).PullLineCountAllocation(eCurPeriod, e_LineCountType.eDeleted)
            GoSub GSCount
            Me.lbl_Delete.Caption = "DELETED/OLD (" & lCount & ")"
            PopulateListBoxes = True
        Case "lst_Region"
            FormName.List = mCAM_Runtime.CategoryObject(lDoc_ID).PullLineCountAllocation(eCurPeriod, e_LineCountType.eRegional)
            GoSub GSCount
            Me.lbl_Region.Caption = "REGIONAL (" & lCount & ")"
            PopulateListBoxes = True
        Case "lst_Sea"
            FormName.List = mCAM_Runtime.CategoryObject(lDoc_ID).PullLineCountAllocation(eCurPeriod, e_LineCountType.eSeasonal)
            GoSub GSCount
            Me.lbl_Seas.Caption = "SEASONAL (" & lCount & ")"
            PopulateListBoxes = True
        Case "lst_Spec"
            FormName.List = mCAM_Runtime.CategoryObject(lDoc_ID).PullLineCountAllocation(eCurPeriod, e_LineCountType.eSpecial)
            GoSub GSCount
            Me.lbl_Spec.Caption = "SPECIALS (" & lCount & ")"
            PopulateListBoxes = True
        Case "lst_TrialCur"
            FormName.List = mCAM_Runtime.CategoryObject(lDoc_ID).PullLineCountAllocation(eCurPeriod, e_LineCountType.eCurTrial)
            GoSub GSCount
            Me.lbl_TrialCur.Caption = "Current & Unsuccessful TRIALS (" & lCount & ")"
            PopulateListBoxes = True
        Case "lst_TrialSuc"
            FormName.List = mCAM_Runtime.CategoryObject(lDoc_ID).PullLineCountAllocation(eCurPeriod, e_LineCountType.eSucTrial)
            GoSub GSCount
            Me.lbl_TrialSuc.Caption = "Successful TRIALS (" & lCount & ")"
            PopulateListBoxes = True
    End Select
    Exit Function
GSCount:
    lCount = FormName.ListCount
    If lCount = 1 Then
        If NZ(FormName.List(0), "") = "" Then lCount = 0
    End If
''    If bAccum Then lRangeCount = lRangeCount + lCount
    Return
End Function

Private Function SetCurProd() As Boolean
    ' Steps through the listbox selected to find the productcode selected
    Dim a As Long
    sCurProd = ""
    SetCurProd = False
    For a = 0 To lstCurLst.ListCount - 1
        If lstCurLst.Selected(a) = True Then
            sCurProd = CStr(lstCurLst.List(a))
            plCurPosit = a
            SetCurProd = True
            Exit Function
        End If
    Next
    MsgBox "Error: no CurProd Set"
End Function

Private Function TransferProd(ByRef lstTo As MSForms.ListBox, ByVal eLCT As e_LineCountType) As Boolean
    ' Transfers a product from one listbox to another.
    ' Does this by affecting the ProductAllocation Dictionary in the cCAM_LineCount_Data Object and then telling it to re-populate the listboxes affected.
    If lstCurLst Is Nothing Then MsgBox "Please select a product to transfer": Exit Function
    If lstTo.Tag = lst_Spec.Tag Or lstCurLst.Tag = lst_Spec.Tag Then
        MsgBox "Products cannot be allocated away from specials to another product class"
        TransferProd = False: Exit Function
    End If
    TransferProd = True
    Call mCAM_Runtime.CategoryObject(lDoc_ID).AmendLineCountData(eCurPeriod, sCurProd, eLCT)
    If lDoc_ID > 0 Then Call mCAM_Runtime.ControlObject(lDoc_ID, , "Rendercell")
    TransferProd = RefreshAllListboxes
End Function
Private Function IfDoubleClickedList(ByRef lst As MSForms.ListBox) As Boolean
    ' Show the CBA_CR_ProductViewer UserForm for the product doubleclicked on
    Dim F As fCAM_ProductViewer
    Dim sProd As String, a As Long
    IfDoubleClickedList = False
    Set lstCurLst = lst
    DeSelAllListsButMe
    For a = 0 To lst.ListCount - 1
        If lst.Selected(a) = True Then
            sProd = lst.List(a)
            IfDoubleClickedList = True
            Exit For
        End If
    Next
    Set F = New fCAM_ProductViewer
    If F.SetData(sProd) = False Then Set F = Nothing: IfDoubleClickedList = False: Exit Function
    F.Show vbModeless
End Function
Private Sub cmd_Active_Click()
    If Me.lst_Delete.ListIndex = -1 Then Exit Sub
    TransferProd Me.lst_Core, eCore
End Sub
Private Function InformOfWarningSuppression(ByVal PCodeAndDesc As String) As Boolean
    'THIS WILL BE IMPLEMENTED IN V2
End Function
Private Sub cmd_Brand_Click(): TransferProd Me.lst_Brand, eBranded: End Sub
Private Sub cmd_Core_Click(): TransferProd Me.lst_Core, eCore: End Sub
Private Sub cmd_Delete_Click(): TransferProd Me.lst_Delete, eDeleted: End Sub
Private Sub cmd_Region_Click(): TransferProd Me.lst_Region, eRegional: End Sub
Private Sub cmd_Seas_Click(): TransferProd Me.lst_Sea, eSeasonal: End Sub
Private Sub cmd_Spec_Click(): TransferProd Me.lst_Spec, eSpecial: End Sub
Private Sub cmd_TrialCur_Click(): TransferProd Me.lst_TrialCur, eCurTrial: End Sub
Private Sub cmd_TrialSuc_Click()
Dim MS As MSForms.ListBox
    Set MS = lst_TrialSuc
    TransferProd MS, eSucTrial
End Sub
Private Sub lst_Brand_DblClick(ByVal Cancel As MSForms.ReturnBoolean): SetCurProd: IfDoubleClickedList Me.lst_Brand: End Sub
Private Sub lst_Core_DblClick(ByVal Cancel As MSForms.ReturnBoolean): SetCurProd: IfDoubleClickedList Me.lst_Core: End Sub
Private Sub lst_Delete_DblClick(ByVal Cancel As MSForms.ReturnBoolean): SetCurProd: IfDoubleClickedList Me.lst_Delete: End Sub
Private Sub lst_Region_DblClick(ByVal Cancel As MSForms.ReturnBoolean): SetCurProd: IfDoubleClickedList Me.lst_Region: End Sub
Private Sub lst_Sea_DblClick(ByVal Cancel As MSForms.ReturnBoolean): SetCurProd: IfDoubleClickedList Me.lst_Sea: End Sub
Private Sub lst_TrialCur_DblClick(ByVal Cancel As MSForms.ReturnBoolean): SetCurProd: IfDoubleClickedList Me.lst_TrialCur: End Sub
Private Sub lst_TrialSuc_DblClick(ByVal Cancel As MSForms.ReturnBoolean): SetCurProd: IfDoubleClickedList Me.lst_TrialSuc: End Sub
Private Sub lst_Spec_DblClick(ByVal Cancel As MSForms.ReturnBoolean): SetCurProd: IfDoubleClickedList Me.lst_Spec: End Sub
Private Sub lst_Brand_Click(): Set lstCurLst = lst_Brand: SetCurProd: DeSelAllListsButMe: End Sub
Private Sub lst_Core_Click(): Set lstCurLst = lst_Core: SetCurProd: DeSelAllListsButMe: End Sub
Private Sub lst_Delete_Click(): Set lstCurLst = lst_Delete: SetCurProd: DeSelAllListsButMe: End Sub
Private Sub lst_Region_Click(): Set lstCurLst = lst_Region: SetCurProd: DeSelAllListsButMe: End Sub
Private Sub lst_Sea_Click(): Set lstCurLst = lst_Sea: SetCurProd: DeSelAllListsButMe: End Sub
Private Sub lst_TrialCur_Click(): Set lstCurLst = Me.lst_TrialCur: SetCurProd: DeSelAllListsButMe: End Sub
Private Sub lst_TrialSuc_Click(): Set lstCurLst = lst_TrialSuc: SetCurProd: DeSelAllListsButMe: End Sub
Private Sub lst_Spec_Click(): Set lstCurLst = lst_Spec: SetCurProd: DeSelAllListsButMe: End Sub
Private Property Get sCurProd() As String: sCurProd = psCurProd: End Property
Private Property Let sCurProd(ByVal sNewValue As String): psCurProd = sNewValue: End Property
Private Property Get sdPeriodNames() As Scripting.Dictionary: Set sdPeriodNames = psdPeriodNames: End Property
Private Property Set sdPeriodNames(ByVal objNewValue As Scripting.Dictionary): Set psdPeriodNames = objNewValue: End Property
Private Property Get lstCurLst() As MSForms.ListBox: Set lstCurLst = plstCurLst: End Property
Private Property Set lstCurLst(ByVal objNewValue As MSForms.ListBox): Set plstCurLst = objNewValue: End Property
Private Property Get lCurPosit() As Long: lCurPosit = plCurPosit: End Property
Private Property Let lCurPosit(ByVal lNewValue As Long): plCurPosit = lNewValue: End Property
''Private Property Get lRangeCount() As Long: lRangeCount = plRangeCount: End Property
Public Property Get bIsActive() As Boolean: bIsActive = pbIsActive: End Property
Private Property Let bIsActive(ByVal bNewValue As Boolean): pbIsActive = bNewValue: End Property
''Private Property Let lRangeCount(ByVal lNewValue As Long): plRangeCount = lNewValue: End Property
Private Property Get sdOriginalLineCountDataObjects() As Scripting.Dictionary: Set sdOriginalLineCountDataObjects = psdOriginalLineCountDataObjects: End Property
Private Property Set sdOriginalLineCountDataObjects(ByVal objNewValue As Scripting.Dictionary): Set psdOriginalLineCountDataObjects = objNewValue: End Property
Private Property Get eCurPeriod() As e_LineCountPeriod: eCurPeriod = peCurPeriod: End Property
Private Property Let eCurPeriod(ByVal eNewValue As e_LineCountPeriod): peCurPeriod = eNewValue: End Property
Public Property Get lDoc_ID() As Long: lDoc_ID = plDoc_ID: End Property
Public Property Let lDoc_ID(ByVal NewValue As Long)
Dim LC As cCAM_LineCount_Data
Dim LCOriginal As cCAM_LineCount_Data
Dim v As Variant
    
    plDoc_ID = NewValue
    If mCAM_Runtime.CategoryObject(plDoc_ID) Is Nothing Then
        bIsActive = False
    Else
        bIsActive = True
        If mCAM_Runtime.CategoryObject(plDoc_ID).eDocumentType = eDocuNone Then
            Me.lbl_Identifier.Caption = "Active Line Count"
            cmd_Reset.Visible = True
            cmd_SendToMaster.Visible = False
        Else
            cmd_Reset.Visible = True
            cmd_SendToMaster.Visible = True
        End If
        Me.cmd_Spec.Visible = False
        Set sdOriginalLineCountDataObjects = New Scripting.Dictionary
        For Each v In mCAM_Runtime.CategoryObject(plDoc_ID).sdLineCountDic
            Set LC = New cCAM_LineCount_Data
            Set LCOriginal = mCAM_Runtime.CategoryObject(plDoc_ID).sdLineCountDic(v)
            If LC.Copy(LCOriginal) = False Then MsgBox "error in generating LineCountObject Copy": Exit Property
            sdOriginalLineCountDataObjects.Add v, LC
        Next
    End If
    If cbo_Period = "MAT" Then Call cbo_Period_Change Else cbo_Period = "MAT"


End Property
