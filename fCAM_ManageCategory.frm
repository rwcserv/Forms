VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fCAM_ManageCategory 
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12435
   OleObjectBlob   =   "fCAM_ManageCategory.frx":0000
   StartUpPosition =   2  'CenterScreen
   Tag             =   "fCAM_ManageCategory"
End
Attribute VB_Name = "fCAM_ManageCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                     ' fCam_ManageCategory
Private pbACG As Boolean
Private pbIsActive As Boolean
Private pbUpdateMode As Boolean
Private plACGCatNo As Long
Private plACGCGno As Long
Private plACGSCGno As Long
Private plLCGCGno As Long
Private plLCGSCGno As Long
Private psCategoryName As String
Private pbRedneringDDB As Boolean
' Used to manage the category selections available in the dropdown list
''''Points to mention:
''''You should not be using the 'cActiveDataObject'
''''All category management setups are done to the master and thats that.


Private Function ValidateCboEntries(ByRef cbo As MSForms.ComboBox) As Boolean
    ' Dependent on ACG, the ACGCat,CG and SCG Listboxes validate that a value entered is in the list of the combo box.
    ' If not, False and zero out the respective private variable, else True and populate the private variable
    Dim v As Variant
    If cbo = "" Then ValidateCboEntries = True: Exit Function
    If cbo <> "" Then
        For Each v In cbo.List
            If Trim(v) = Trim(cbo.Value) Then ValidateCboEntries = True: Exit Function
        Next
    End If
End Function
Private Function AmendCategoryList(ByVal AddNotDelete As Boolean, Optional ByRef P As cCBA_Prod) As Boolean
    ' Either Adds another instance of the CBA_Prod class to the CategoryList Scripting.Dictionary or finds the correlating CBA_Prod in the listbox selected and removes from CategoryList
    Dim arr As Variant
    Dim a As Long
    On Error GoTo Err_Routine
    CBA_Error = ""
    If AddNotDelete Then
        If bACG Then
            If lACGCGno = 0 And lACGCatNo > 0 Then
                MsgBox "Please select at least one Commodity Group", vbOKOnly: AmendCategoryList = False: Exit Function
            ElseIf lACGCGno > 0 And lACGSCGno = 0 Then
                If MsgBox("Add all SCGs for the whole Commodity Group to the CAMERA Category?", vbYesNo) = vbNo Then MsgBox "Please select a Sub-Commodty Group", vbOKOnly: AmendCategoryList = False: Exit Function
                arr = mCAM_Runtime.cRibbonData.GetCGListing(eASCGnum, True, True, , lACGCGno)
                For a = LBound(arr) To UBound(arr)
                    lACGSCGno = arr(a)
                    AmendCategoryList = AddToCategoryInRibbon
                Next
                lACGSCGno = 0
            ElseIf lACGCGno > 0 And lACGSCGno > 0 Then
                AmendCategoryList = AddToCategoryInRibbon
            End If
        Else
            If lLCGCGno > 0 And (lLCGSCGno = 0 And lLCGCGno <> 2) Then
                If MsgBox("Add all SCGs for the whole Commodity Group to the CAMERA Category?", vbYesNo) = vbNo Then MsgBox "Please select a Sub-Commodty Group", vbOKOnly: AmendCategoryList = False: Exit Function
                arr = mCAM_Runtime.cRibbonData.GetCGListing(eLegSCG, False, True, , lLCGCGno)
                For a = LBound(arr) To UBound(arr)
                    lLCGSCGno = arr(a)
                    AmendCategoryList = AddToCategoryInRibbon
                Next
                lLCGSCGno = 0
            ElseIf lLCGCGno > 0 And (lLCGSCGno > 0 Or lLCGCGno = 2) Then 'accomodated for sparkiling wine where scg is 0, eggs also?
                AmendCategoryList = AddToCategoryInRibbon
            End If
        End If
    Else
        If P Is Nothing Then
            MsgBox "Cannot remove an object from CategoryDic without a comparable cCBA_Prod Object(P)"
        Else
            AmendCategoryList = mCAM_Runtime.cRibbonData.AmendCategoryDic(True, sCategoryName, P)
        End If
    End If
Exit_Routine:

    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-fCam_ManageCategory.amendCategoryList", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Cam", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function
Private Sub lst_CGAss_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ' Offers to remove the selection. MsgBox appears that asks if want to delete selected line, if yes then run amendCateogryList(False) and run populateList
    Dim a As Long
    Dim arr As Variant
    Dim P As cCBA_Prod
    On Error GoTo Err_Routine
    CBA_Error = ""
    Set P = New cCBA_Prod: P.Build 0, "GetsLetsOnly"
    If bACG Then
        Me.cbo_ACGCat = ""
    Else
        Me.cbo_CG = "": Me.cbo_SCG = ""
    End If
    For a = 0 To lst_CGAss.ListCount - 1
        If lst_CGAss.Selected(a) Then
            arr = Split(lst_CGAss.List(a), "|")
            If bACG Then
                cbo_ACGCat = Trim(arr(0)): cbo_CG = Trim(arr(1)): cbo_SCG = Trim(arr(2))
                P.lACatNum = lACGCatNo: P.lACGNum = lACGCGno: P.lASCGNum = lACGSCGno
            Else
                cbo_CG = Trim(arr(0)): cbo_SCG = Trim(arr(1))
                P.lLegCG = lLCGCGno: P.lLegSCG = lLCGSCGno
            End If
            If AmendCategoryList(False, P) = True Then PopulateList
            Exit For
        End If
    Next
Exit_Routine:

    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-fCam_ManageCategory.lst_CGAss_DblClick", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Cam", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub
Private Function AddToCategoryInRibbon() As Boolean
    Dim P As cCBA_Prod
    Dim rerunSetData As Boolean
    On Error GoTo Err_Routine
    CBA_Error = ""
    If bACG Then
        If mCAM_Runtime.cRibbonData.GetCategoryList(sCategoryName, True) Is Nothing Then
            rerunSetData = True
        ElseIf mCAM_Runtime.cRibbonData.GetCategoryList(sCategoryName, True).Exists(Format(lACGCatNo, "000") & Format(lACGCGno, "000") & Format(lACGSCGno, "000")) Then
            Exit Function
        End If
        Set P = New cCBA_Prod: P.Build 0, "GetsLetsOnly"
        P.lACatNum = lACGCatNo: P.lACGNum = lACGCGno: P.lASCGNum = lACGSCGno
        P.sACat = mCAM_Runtime.cRibbonData.GetCGListing(eACat, True, True, lACGCatNo, lACGCGno, lACGSCGno)(0)
        P.sACGDesc = mCAM_Runtime.cRibbonData.GetCGListing(eACGdesc, True, True, lACGCatNo, lACGCGno, lACGSCGno)(0)
        P.sASCGDesc = mCAM_Runtime.cRibbonData.GetCGListing(eASCGdesc, True, True, lACGCatNo, lACGCGno, lACGSCGno)(0)
        AddToCategoryInRibbon = mCAM_Runtime.cRibbonData.AmendCategoryDic(False, sCategoryName, P)
    Else
        If mCAM_Runtime.cRibbonData.GetCategoryList(sCategoryName, False) Is Nothing Then
            rerunSetData = True
        ElseIf mCAM_Runtime.cRibbonData.GetCategoryList(sCategoryName, False).Exists(Format(lLCGCGno, "000") & Format(lLCGSCGno, "000")) Then
            Exit Function
        End If
        
        Set P = New cCBA_Prod: P.Build 0, "GetsLetsOnly"
        P.lLegCG = lLCGCGno: P.lLegSCG = lLCGSCGno
        P.sLegCGDesc = mCAM_Runtime.cRibbonData.GetCGListing(eLegCGdesc, False, True, , lLCGCGno, lLCGSCGno)(0)
        P.sLegSCGDesc = mCAM_Runtime.cRibbonData.GetCGListing(eLegSCGdesc, False, True, , lLCGCGno, lLCGSCGno)(0)
        AddToCategoryInRibbon = mCAM_Runtime.cRibbonData.AmendCategoryDic(False, sCategoryName, P)
    End If
    If rerunSetData = True Then SetData: PopulateList
Exit_Routine:

    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-fCam_ManageCategory.addToCategoryInRibbon", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Cam", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function
Private Sub cmd_Add_Click()
    ' If all the values in the comboboxes are good to go (runs function validateCboEntries), then update the categorylist (amendCategoryList(True)) and the listbox (populateList)
    If ValidateCboEntries(cbo_CG) = False Or ValidateCboEntries(cbo_SCG) = False Or ValidateCboEntries(cbo_ACGCat) = False Then
        MsgBox "Invalid Entry": Exit Sub
    Else
        If AmendCategoryList(True) = True Then PopulateList
    End If
End Sub

Private Sub UserForm_Initialize()
    ' Pulls CategoryList from CAMERA_Runtime and populates the combobox options.
    ' Sets the combobox value equal to the setting on ribbon (setData)
    Dim lTop, lLeft, lRow, lcol As Long
    Me.StartUpPosition = 0
    Me.Top = Application.Top + (Application.Height / 2) - (Me.Height / 2)
    Me.Left = Application.Left + (Application.Width / 2) - (Me.Width / 2)
    Me.lblVersion.Caption = CBA_getVersionStatus(g_GetDB("Cam"), CBA_Cam_Ver, "Camera Tool", "Camera")  ' Get the latest version
    SetData True
End Sub
Private Function SetData(Optional ByVal bSetup As Boolean = False) As Boolean
    ' Pulls data from Runtime(getAllCategoryNames&ActiveRibbonCategory).
    ' Sets the list for comboboxes. Sets isActive = True.
    ' Sets ACG to LCG as default(triggers event).
    ' Sets the category value from the ribbon(triggers event)
    Dim a As Long
    Dim arr() As Variant
    Dim col As Collection
    On Error GoTo Err_Routine
    CBA_Error = ""
    If bSetup = True Then
        bIsActive = True
        cbo_ACG.AddItem "ACG"
        cbo_ACG.AddItem "Legacy"
        cbo_ACG = "Legacy"
    Else
        cbo_CatName.List = mCAM_Runtime.cRibbonData.GetAllCategoryNames
        If bACG = True Then
            arr = mCAM_Runtime.cRibbonData.GetCGListing(eACatNum, True, True, , , , True)
            CBA_BasicFunctions.CBA_Sort1DArray arr, LBound(arr), UBound(arr)
            CBA_BasicFunctions.CBA_UniqueValuesFor1DArray arr
            Me.cbo_ACGCat.List = arr
            cbo_CG.List = mCAM_Runtime.cRibbonData.GetCGListing(eACGnum, True, True, , , , True)
        Else
            cbo_CG.List = mCAM_Runtime.cRibbonData.GetCGListing(eLegCG, False, True, , , , True)
        End If
    End If
    
'    If bSetup = True And Not mCAM_Runtime.cActiveDataObject Is Nothing Then
'        cbo_CatName.Value = mCAM_Runtime.cActiveDataObject.sCategoryName
'    End If
Exit_Routine:

    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-fCam_ManageCategory.SetData", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Cam", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function
Private Function RenderDropDownBoxes() As Boolean
    
pbRedneringDDB = True
If bACG = True Then
    cbo_CG.List = mCAM_Runtime.cRibbonData.GetCGListing(eACGnum, True, True, lACGCatNo, , , True)
    cbo_SCG.List = mCAM_Runtime.cRibbonData.GetCGListing(eASCGnum, True, True, lACGCatNo, lACGCGno, , True)
    If lACGCatNo > 0 Then
        cbo_ACGCat = mCAM_Runtime.cRibbonData.GetCGListing(eACatNum, True, True, lACGCatNo, lACGCGno, lACGSCGno, True)(0)
        If lACGCGno > 0 Then
            cbo_CG = mCAM_Runtime.cRibbonData.GetCGListing(eACGnum, True, True, lACGCatNo, lACGCGno, lACGSCGno, True)(0)
            If lACGSCGno > 0 Then
                cbo_SCG = mCAM_Runtime.cRibbonData.GetCGListing(eASCGnum, True, True, lACGCatNo, lACGCGno, lACGSCGno, True)(0)
            Else
                cbo_SCG = ""
            End If
        Else
            cbo_CG = ""
            cbo_SCG = ""
        End If
    Else
        cbo_ACGCat = ""
        cbo_CG = ""
        cbo_SCG = ""
    End If
Else
    'cbo_CG.List = mCAM_Runtime.cRibbonData.GetCGListing(eLegCG, False, True, , lLCGCGno, lLCGSCGno, True)
    cbo_SCG.List = mCAM_Runtime.cRibbonData.GetCGListing(eLegSCG, False, True, , lLCGCGno, lLCGSCGno, True)
    If lLCGCGno > 0 Then
        cbo_CG = mCAM_Runtime.cRibbonData.GetCGListing(eLegSCG, False, True, , lLCGCGno, , True)(0)
        If lLCGSCGno > 0 Then
            cbo_SCG = mCAM_Runtime.cRibbonData.GetCGListing(eLegSCG, False, True, , lLCGCGno, lLCGSCGno, True)(0)
        Else
            cbo_SCG = ""
        End If
    Else
        cbo_CG = ""
        cbo_SCG = ""
    End If
End If
pbRedneringDDB = False

End Function
Private Sub cbo_ACGCat_Change()
    If pbRedneringDDB = True Then Exit Sub
    If ValidateCboEntries(cbo_ACGCat) = False Then Exit Sub
    lACGCGno = 0
    lACGSCGno = 0
    If cbo_ACGCat = "" Then
        lACGCatNo = 0

    Else
        lACGCatNo = CLng(Mid(cbo_ACGCat, 1, InStr(1, cbo_ACGCat, "-") - 1))
    End If
    RenderDropDownBoxes
    'cbo_CG.Clear: cbo_SCG.Clear
    'cbo_CG.List = mCAM_Runtime.cRibbonData.GetCGListing(eACGnum, True, True, lACGCatNo, , , True)
End Sub
Private Sub cbo_CatName_Change()
    'If pbRedneringDDB = True Then Exit Sub
    CheckUpdateMode
    sCategoryName = cbo_CatName
End Sub
Private Sub cbo_ACG_Change()
    'If pbRedneringDDB = True Then Exit Sub
    If ValidateCboEntries(cbo_ACG) = False Then cbo_ACG = bACG: Exit Sub
    If cbo_ACG = "ACG" Then bACG = True Else bACG = False
    'cbo_ACGCat.Clear: cbo_CG.Clear: cbo_SCG.Clear: lst_CGAss.Clear
    SetData
    PopulateList
End Sub
Private Sub cbo_CG_Change()
    If pbRedneringDDB = True Then Exit Sub
    If ValidateCboEntries(cbo_CG) = False Then cbo_SCG.Clear: Exit Sub
    If bACG = True Then
        lACGSCGno = 0
        If cbo_CG = "" Then lACGCGno = 0:  Exit Sub
        lACGCGno = CLng(Mid(cbo_CG, 1, InStr(1, cbo_CG, "-") - 1))
        'Application.EnableEvents = False
        'Me.cbo_ACGCat = mCAM_Runtime.cRibbonData.GetCGListing(eACatNum, True, True, , CLng(Mid(cbo_CG, 1, InStr(1, cbo_CG, "-") - 1)))(0) & "-" & mCAM_Runtime.cRibbonData.GetCGListing(eACat, True, True, , CLng(Mid(cbo_CG, 1, InStr(1, cbo_CG, "-") - 1)))(0)
        'Application.EnableEvents = True
        'Me.cbo_SCG.List = mCAM_Runtime.cRibbonData.GetCGListing(eASCGnum, True, True, , CLng(Mid(cbo_CG, 1, InStr(1, cbo_CG, "-") - 1)), , True)
        RenderDropDownBoxes
    Else
        lLCGSCGno = 0
        If cbo_CG = "" Then lLCGCGno = 0:  Exit Sub
        lLCGCGno = CLng(Mid(cbo_CG, 1, InStr(1, cbo_CG, "-") - 1))
        Me.cbo_SCG.List = mCAM_Runtime.cRibbonData.GetCGListing(eLegSCG, False, True, , CLng(Mid(cbo_CG, 1, InStr(1, cbo_CG, "-") - 1)), , True)
    End If
End Sub
Private Sub cbo_SCG_Change()
    If pbRedneringDDB = True Then Exit Sub
    If ValidateCboEntries(cbo_SCG) = False Then Exit Sub
    If bACG = True Then
        If cbo_SCG = "" Then lACGSCGno = 0: Exit Sub
        lACGSCGno = CLng(Mid(cbo_SCG, 1, InStr(1, cbo_SCG, "-") - 1))
    Else
        If cbo_SCG = "" Then lLCGSCGno = 0: Exit Sub
        lLCGSCGno = CLng(Mid(cbo_SCG, 1, InStr(1, cbo_SCG, "-") - 1))
    End If
    RenderDropDownBoxes
End Sub
Private Function PopulateList() As Boolean
    ' Removes all in the list and repopulates it with the required values from the CategoryList
    ' (adds String value (ACGCatno-ACGCatDescription | CGNo-CGDescription | SCGNo-SCGDescription))

    Dim P As cCBA_Prod
    Dim a As Long
    Dim arr As Variant
    Dim dic As Scripting.Dictionary
    On Error GoTo Err_Routine
    CBA_Error = ""
    Me.lst_CGAss.Clear
    Set dic = mCAM_Runtime.cRibbonData.GetCategoryList(cbo_CatName, bACG)
    arr = mCAM_Runtime.cRibbonData.GetCGListing(eALL, bACG)
    If Not dic Is Nothing Then
        For a = LBound(arr, 2) To UBound(arr, 2)
            If bACG Then
                If dic.Exists(Format(CLng(arr(0, a)), "000") & Format(CLng(arr(2, a)), "000") & Format(CLng(arr(4, a)), "000")) Then
                    Me.lst_CGAss.AddItem arr(0, a) & "-" & arr(1, a) & " | " & arr(2, a) & "-" & arr(3, a) & " | " & arr(4, a) & "-" & arr(5, a)
                End If
            Else
                If dic.Exists(CStr(Format(arr(0, a), "000") & Format(arr(2, a), "000"))) Then
                    Me.lst_CGAss.AddItem arr(0, a) & "-" & arr(1, a) & " | " & arr(2, a) & "-" & arr(3, a)
                End If
            End If
        Next
    End If
Exit_Routine:

    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-fCam_ManageCategory.PopulateList", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Cam", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function
Private Function CheckUpdateMode() As Boolean
    ' Checks to see if the value in cbo_CatName is in the CategoryList, if so, sets updateMode = True, else sets updateMode = False.
    If mCAM_Runtime.cRibbonData.GetCategoryList(cbo_CatName, bACG) Is Nothing Then bUpdateMode = True Else bUpdateMode = False
End Function
Private Property Let bACG(ByVal bNewValue As Boolean)
    If bNewValue = True Then lLCGCGno = 0: lLCGSCGno = 0
    If bNewValue = False Then lACGCatNo = 0: lACGCGno = 0: lACGSCGno = 0
    RenderDropDownBoxes
    pbACG = bNewValue
    Me.cbo_ACGCat.Enabled = pbACG
End Property
Private Property Let bUpdateMode(ByVal bNewValue As Boolean)
    If sCategoryName = cbo_CatName Then Exit Property
    pbUpdateMode = bNewValue
    PopulateList
End Property
Public Property Get bIsActive() As Boolean: bIsActive = pbIsActive: End Property
Private Property Let bIsActive(ByVal bNewValue As Boolean): pbIsActive = bNewValue: End Property
Private Property Get bUpdateMode() As Boolean: bUpdateMode = pbUpdateMode: End Property
Private Property Get lACGCatNo() As Long: lACGCatNo = plACGCatNo: End Property
Private Property Let lACGCatNo(ByVal lNewValue As Long): plACGCatNo = lNewValue: End Property
Private Property Get lACGCGno() As Long: lACGCGno = plACGCGno: End Property
Private Property Let lACGCGno(ByVal lNewValue As Long): plACGCGno = lNewValue: End Property
Private Property Get lACGSCGno() As Long: lACGSCGno = plACGSCGno: End Property
Private Property Let lACGSCGno(ByVal lNewValue As Long): plACGSCGno = lNewValue: End Property
Private Property Get lLCGCGno() As Long: lLCGCGno = plLCGCGno: End Property
Private Property Let lLCGCGno(ByVal lNewValue As Long): plLCGCGno = lNewValue: End Property
Private Property Get lLCGSCGno() As Long: lLCGSCGno = plLCGSCGno: End Property
Private Property Let lLCGSCGno(ByVal lNewValue As Long): plLCGSCGno = lNewValue: End Property
Private Property Get bACG() As Boolean: bACG = pbACG: End Property
Private Property Get sCategoryName() As String: sCategoryName = psCategoryName: End Property
Private Property Let sCategoryName(ByVal sNewValue As String): psCategoryName = sNewValue: End Property
Private Sub UserForm_Terminate()
    If sCategoryName <> "" Then
        mCAM_Runtime.cRibbonData.InterfaceCategoryDicChangesToDB
        mCAM_Runtime.RefreshCategorySelectionDropDownOnRibbon
    End If
End Sub
Private Property Get bRedneringDDB() As Boolean: bRedneringDDB = pbRedneringDDB: End Property
Private Property Let bRedneringDDB(ByVal NewValue As Boolean): pbRedneringDDB = NewValue: End Property
