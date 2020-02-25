VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CBA_COM_frm_MatchingTool 
   ClientHeight    =   10470
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   22335
   OleObjectBlob   =   "CBA_COM_frm_MatchingTool.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "CBA_COM_frm_MatchingTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private CCM_WWArr() As Variant
Private CCM_ColesArr() As Variant
Private CCM_dmArr() As Variant
Private CCM_fcArr() As Variant
Private CCM_IsActiveDataSet(1 To 4) As Boolean
Private CCM_PriceType As String
Private CCM_priceonshow As String
Private CCM_Statelookup As String
Private CCM_UserDefinedState As Boolean
Private CCM_Comp2Find As String
Private CCM_CompData As MatchTypeData
Private CCM_Activation As Boolean
Private AldiProd As Aprod
Private CCM_Filter As FilterVal
Private CCM_Matches As MatchMap
Private CCM_CurSel As curSel
Private Type CompDics
    WW As Scripting.Dictionary
    Coles As Scripting.Dictionary
    DM As Scripting.Dictionary
    FC As Scripting.Dictionary
End Type

Private Type CompCols
    WW As Collection
    Coles As Collection
    DM As Collection
    FC As Collection
End Type
Private Type FilterVal
    Description As String
    Packsize As String
    Special As Boolean
    Price As Single
    Matches As CompCols
    MetaData As String
End Type
Private Type Aprod
    PCode As Long
    pCG As Long
    PSCG As Long
End Type
Private Type Cprod
    PCode As String
    Pack As String
    PDescription As String
End Type
Private Type curSel
    Name As String
    Pack As String
End Type

'#RW Added new mousewheel routines 190701
Private Sub Box_ProdList_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call mw_SetBoxHook(Box_ProdList)
End Sub
Private Sub box_Options_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call mw_SetBoxHook(box_Options)
End Sub

Sub setMatchType(ByVal OBButton As String)
    Dim ob As Control
    For Each ob In Me.Controls
        If ob.Name = OBButton Then
            If InStr(1, OBButton, "_ML") > 0 Or InStr(1, OBButton, "_CB") > 0 Or InStr(1, OBButton, "_PB") > 0 Then
                MultiPage1.Value = 2
            ElseIf InStr(1, OBButton, "ob_W") > 0 Then
                MultiPage1.Value = 1
            ElseIf InStr(1, OBButton, "ob_C") > 0 Then
                MultiPage1.Value = 0
            ElseIf InStr(1, OBButton, "ob_DM") > 0 Or InStr(1, OBButton, "ob_FC") > 0 Then
                MultiPage1.Value = 4
            ElseIf InStr(1, OBButton, "Produce") > 0 Then
                MultiPage1.Value = 3
            End If
            ob.Value = 1
            Exit For
        End If
    Next
End Sub
Function CCM_getMatched(ByVal Competitor As String) As Collection
    If Competitor = "WW" Then
        Set CCM_getMatched = CCM_Filter.Matches.WW
    ElseIf Competitor = "Coles" Then
        Set CCM_getMatched = CCM_Filter.Matches.Coles
    ElseIf Competitor = "DM" Then
        Set CCM_getMatched = CCM_Filter.Matches.DM
    ElseIf Competitor = "FC" Then
        Set CCM_getMatched = CCM_Filter.Matches.FC
    End If
End Function
Sub addProdtoList(ByVal ProdToAdd As String)
    Dim bOutput  As Boolean
    If IsNumeric(ProdToAdd) = True Then
        bOutput = CCM_SQLQueries.CBA_COM_MATCHGenPullSQL("FINDPRODDESC", , , , , ProdToAdd)
        On Error Resume Next
        If CBA_CBISarr(0, 0) < 0 Then
            Err.Clear
        Else
            Me.Box_ProdList.AddItem ProdToAdd & "-" & CBA_CBISarr(0, 0)
            Me.Box_ProdList.Selected(Me.Box_ProdList.ListCount - 1) = True
        End If
        On Error GoTo 0
    End If
End Sub
Private Sub Box_addProdtoList_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim bOutput  As Boolean
    If KeyCode = 13 Then
        bOutput = CCM_SQLQueries.CBA_COM_MATCHGenPullSQL("FINDPRODDESC", , , , , Box_addProdtoList.Value)
        If bOutput = True Then
            Me.Box_ProdList.AddItem Box_addProdtoList.Value & "-" & CBA_CBISarr(0, 0)
            Me.Box_ProdList.Selected(Me.Box_ProdList.ListCount - 1) = True
            Box_addProdtoList.Value = ""
        Else
            MsgBox "Invalid CBIS Product Code Entered", vbOKOnly
            Box_addProdtoList.Value = ""
        End If
    End If
End Sub

Private Sub Box_ProdList_Change()
    Dim lNum As Long, bOutput As Boolean
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    
    For lNum = 0 To Me.Box_ProdList.ListCount - 1
        If Me.Box_ProdList.Selected(lNum) = True And Mid(Me.Box_ProdList.List(lNum, 0), 1, InStr(1, Me.Box_ProdList.List(lNum, 0), "-") - 1) <> AldiProd.PCode Then
            AldiProd.PCode = Mid(Me.Box_ProdList.List(lNum, 0), 1, InStr(1, Me.Box_ProdList.List(lNum, 0), "-") - 1)
            bOutput = CCM_SQLQueries.CBA_COM_MATCHGenPullSQL("isCG", , , , , , CStr(AldiProd.PCode))
            If bOutput = True Then
                AldiProd.pCG = CLng(CBA_CBISarr(0, 0))
                If IsNull(CBA_CBISarr(1, 0)) Then AldiProd.PSCG = 1 Else AldiProd.PSCG = CLng(CBA_CBISarr(1, 0))
                If AldiProd.pCG < 5 Then
                    setAlcoholMode True
                ElseIf AldiProd.pCG = 58 Then
                    setProduceMode True
                Else
                    setStdMode True
                End If
            End If
            If Me.cbx_OnlyMatched = True Then
                updateAssignedMatches
                CCM_CreateCCMArrays
                CCM_PopulateData
            End If
            Exit For
    '        Comp = getSpecificMatch(CCM_Comp2Find)
    '        If Comp <> "" Then
    '            thiscomp = getCompCodeandPack(Comp)
    '            Me.box_Assdesc = thiscomp.PDescription
    '            Me.box_AssPack = thiscomp.Pack
    '        End If
    '        Exit For
        End If
    Next
Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-Box_ProdList_Change", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    'If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub

Private Sub setProduceMode(ByVal SetToProduceMode As Boolean)
    Dim a As Long, bfound As Boolean, CBA_COM_DEFAULT_FORM
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    If CCM_UserDefinedState = True Then Exit Sub
    If CCM_FormState(2) = False Or CCM_Activation = True Then
        If CCM_UserDefinedState = False Then
            ck_Coles.Enabled = True: ck_WW.Enabled = True: ck_DM.Enabled = False: ck_FC.Enabled = False: ck_UD.Enabled = False
        Else
            ck_Coles.Enabled = False: ck_WW.Enabled = False: ck_DM.Enabled = False: ck_FC.Enabled = False: ck_UD.Enabled = True
        End If

        For a = 0 To 3
            If a = 2 Then CCM_FormState(a) = True Else CCM_FormState(a) = False
        Next
        For a = 0 To 4
            If a = 3 Then MultiPage1.Pages(a).Enabled = True Else MultiPage1.Pages(a).Enabled = False
        Next
        frme_SL.Enabled = False
        MultiPage1.Value = 3
checkdatasetsagain:
        CCM_DefineActiveDataSets
            For a = 1 To 2
                If CCM_IsActiveDataSet(a) = True Then
                    bfound = True
                    Exit For
                End If
            Next
            If bfound = True Then
                If a = 1 Then Me.ob_W_NAT_Produce1 = 1
                If a = 2 Then Me.ob_C_Nat_Produce1 = 1
            Else
                If CCM_UserDefinedState = False Then
                    Me.Hide
getDD:
                    Set CBA_COM_DEFAULT_FORM = New CBA_frm_SelectBaseData
                    If CCM_Runtime.CMM_getDefaultDataset > 2 Then CBA_COM_DEFAULT_FORM.Show
                    If CCM_Runtime.CMM_getDefaultDataset > 2 Then GoTo getDD Else Unload CBA_COM_DEFAULT_FORM
                    If CBA_BasicFunctions.isRunningSheetDisplayed = False Then CBA_BasicFunctions.CBA_Running
                    CCM_CreateCCMArrays
                    If CCM_Activation = False Then Me.Show vbModeless
                    GoTo checkdatasetsagain
                Else
                    CCM_CreateCCMArrays
                End If
            End If
        If CCM_Activation = False Then Me.Show vbModeless
    End If
Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-setProduceMode", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
'    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub
Private Sub setAlcoholMode(ByVal SetToAlcoholMode As Boolean)
    Dim a As Long, bfound As Boolean, CBA_COM_DEFAULT_FORM
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    If CCM_UserDefinedState = True Then Exit Sub
    If CCM_FormState(3) = False Or CCM_Activation = True Then
        If CCM_UserDefinedState = False Then
            ck_Coles.Enabled = False: ck_WW.Enabled = False: ck_DM.Enabled = True: ck_FC.Enabled = True: ck_UD.Enabled = False
        Else
            ck_Coles.Enabled = False: ck_WW.Enabled = False: ck_DM.Enabled = False: ck_FC.Enabled = False: ck_UD.Enabled = True
        End If
    
        For a = 0 To 3
            If a = 3 Then CCM_FormState(a) = True Else CCM_FormState(a) = False
        Next
        For a = 0 To 4
            If a = 4 Then MultiPage1.Pages(a).Enabled = True Else MultiPage1.Pages(a).Enabled = False
        Next
        frme_SL.Enabled = True
        MultiPage1.Value = 4
checkdatasetsagain:
        CCM_DefineActiveDataSets
            If CCM_IsActiveDataSet(3) = True Then ck_DM.Value = 1 Else ck_DM.Value = 0
            If CCM_IsActiveDataSet(4) = True Then ck_FC.Value = 1 Else ck_FC.Value = 0
            For a = 3 To 4
                If CCM_IsActiveDataSet(a) = True Then
                    bfound = True
                    Exit For
                End If
            Next
            If bfound = True Then
                If a = 3 Then Me.ob_DM1.Value = 1
                If a = 4 Then Me.ob_FC1.Value = 1
            Else
                If CCM_UserDefinedState = False Then
                    Me.Hide
getDD:
                    Set CBA_COM_DEFAULT_FORM = New CBA_frm_SelectBaseData
                    If CCM_Runtime.CMM_getDefaultDataset < 3 Then CBA_COM_DEFAULT_FORM.Show
                    If CCM_Runtime.CMM_getDefaultDataset < 3 Then GoTo getDD Else Unload CBA_COM_DEFAULT_FORM
                    If CBA_BasicFunctions.isRunningSheetDisplayed = False Then CBA_BasicFunctions.CBA_Running
                    CCM_CreateCCMArrays
                    If CCM_Activation = False Then Me.Show vbModeless
                    GoTo checkdatasetsagain
                Else
                    CCM_CreateCCMArrays
                End If
            End If
    End If
Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-setAlcoholMode", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
'    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub
Private Sub setStdMode(ByVal SetToStdMode As Boolean)
    Dim a As Long, CBA_COM_DEFAULT_FORM, bfound As Boolean
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    If CCM_UserDefinedState = True Then Exit Sub
    If CCM_FormState(1) = False Or CCM_Activation = True Then

        If CCM_UserDefinedState = False Then
            ck_Coles.Enabled = True: ck_WW.Enabled = True: ck_DM.Enabled = False: ck_FC.Enabled = False: ck_UD.Enabled = False
        Else
            ck_Coles.Enabled = False: ck_WW.Enabled = False: ck_DM.Enabled = False: ck_FC.Enabled = False: ck_UD.Enabled = True
        End If
    
        For a = 0 To 3
            If a = 1 Then CCM_FormState(a) = True Else CCM_FormState(a) = False
        Next
        For a = 0 To 4
            If a < 3 Then MultiPage1.Pages(a).Enabled = True Else MultiPage1.Pages(a).Enabled = False
        Next
        frme_SL.Enabled = True
checkdatasetsagain:
        CCM_DefineActiveDataSets
            For a = 1 To 2
                If CCM_IsActiveDataSet(a) = True Then
                    bfound = True
                    Exit For
                End If
            Next
            If bfound = True Then
                If a = 1 Then Me.ob_W_PL1.Value = 1: MultiPage1.Value = 1
                If a = 2 Then Me.ob_C_PL1.Value = 1: MultiPage1.Value = 0
            Else
                If CCM_UserDefinedState = False Then
                    Me.Hide
getDD:
                    Set CBA_COM_DEFAULT_FORM = New CBA_frm_SelectBaseData
                    If CCM_Runtime.CMM_getDefaultDataset > 2 Then CBA_COM_DEFAULT_FORM.Show
                    If CCM_Runtime.CMM_getDefaultDataset > 2 Then GoTo getDD Else Unload CBA_COM_DEFAULT_FORM
                    If CBA_BasicFunctions.isRunningSheetDisplayed = False Then CBA_BasicFunctions.CBA_Running
                    CCM_CreateCCMArrays
                    If CCM_Activation = False Then Me.Show vbModeless
                    GoTo checkdatasetsagain
                Else
                    CCM_CreateCCMArrays
                End If
            End If
    End If
Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-setStdMode", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    'If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub
Private Sub setUDMode()
    Dim a As Long
    On Error GoTo Err_Routine
    CBA_ErrTag = ""

    CCM_DefineActiveDataSets
    For a = 1 To 4
        If CCM_IsActiveDataSet(a) = True Then
            If a = 1 Then
                MultiPage1.Pages(1).Enabled = True
                MultiPage1.Pages(2).Enabled = True
                MultiPage1.Pages(3).Enabled = True
                Me.ob_W_PL1.Value = 1
            ElseIf a = 2 Then
                MultiPage1.Pages(0).Enabled = True
                MultiPage1.Pages(2).Enabled = True
                MultiPage1.Pages(3).Enabled = True
                Me.ob_C_PL1.Value = 1
            ElseIf a = 3 Then
                MultiPage1.Pages(4).Enabled = True
                Me.ob_DM1.Value = 1
            ElseIf a = 4 Then
                MultiPage1.Pages(4).Enabled = True
                Me.ob_FC1.Value = 1
            End If
        End If
    Next
    Me.ck_WW.Enabled = False: Me.ck_Coles.Enabled = False: Me.ck_DM.Enabled = False: Me.ck_FC.Enabled = False: Me.ck_UD.Enabled = True
Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-setUDMode", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
'    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub
Private Sub but_Export_Click()
    Dim a As Long, b As Long, wb_Export, wks_Export, lNum As Long
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    Application.ScreenUpdating = False
    CCM_DefineActiveDataSets
    Me.Hide
    
    For a = 1 To 4
        If CCM_IsActiveDataSet(a) = True Then
            Set wb_Export = Application.Workbooks.Add
            Set wks_Export = wb_Export.Sheets(1)
            wks_Export.Cells(1, 1).Value = "Competitor"
            wks_Export.Cells(1, 2).Value = "Description"
            wks_Export.Cells(1, 3).Value = "Packsize"
            wks_Export.Cells(1, 4).Value = "Price"
            wks_Export.Cells(1, 5).Value = "MetaData"
            Range(wks_Export.Cells(1, 1), wks_Export.Cells(1, 5)).Font.Bold = True
            Exit For
        End If
    Next
    
    lNum = 1
    For a = 1 To 4
    If CCM_IsActiveDataSet(a) = True Then
        With wks_Export
            If a = 1 Then
                For b = 1 To UBound(CCM_WWArr, 2)
                    lNum = lNum + 1
                    .Cells(lNum, 1).Value = "WW"
                    .Cells(lNum, 2).Value = CCM_WWArr(1, b)
                    .Cells(lNum, 3).Value = CCM_WWArr(2, b)
                    .Cells(lNum, 4).Value = CCM_WWArr(3, b)
                    .Cells(lNum, 5).Value = CCM_WWArr(4, b)
                Next
            ElseIf a = 2 Then
                For b = 1 To UBound(CCM_ColesArr, 2)
                    lNum = lNum + 1
                    .Cells(lNum, 1).Value = "Coles"
                    .Cells(lNum, 2).Value = CCM_ColesArr(1, b)
                    .Cells(lNum, 3).Value = CCM_ColesArr(2, b)
                    .Cells(lNum, 4).Value = CCM_ColesArr(3, b)
                    .Cells(lNum, 5).Value = CCM_ColesArr(4, b)
                Next
            ElseIf a = 3 Then
                For b = 1 To UBound(CCM_dmArr, 2)
                    lNum = lNum + 1
                    .Cells(lNum, 1).Value = "DM"
                    .Cells(lNum, 2).Value = CCM_dmArr(1, b)
                    .Cells(lNum, 3).Value = CCM_dmArr(2, b)
                    .Cells(lNum, 4).Value = CCM_dmArr(3, b)
                    .Cells(lNum, 5).Value = CCM_dmArr(4, b)
                Next
            ElseIf a = 4 Then
                For b = 1 To UBound(CCM_fcArr, 2)
                    lNum = lNum + 1
                    .Cells(lNum, 1).Value = "FC"
                    .Cells(lNum, 2).Value = CCM_fcArr(1, b)
                    .Cells(lNum, 3).Value = CCM_fcArr(2, b)
                    .Cells(lNum, 4).Value = CCM_fcArr(3, b)
                    .Cells(lNum, 5).Value = CCM_fcArr(4, b)
                Next
            End If
            Range(.Cells(1, 1), .Cells(1, 5)).EntireColumn.AutoFit
            Range(.Cells(1, 1), .Cells(1, 5)).AutoFilter
       End With
    End If
    Next
    Application.ScreenUpdating = True
    If CCM_Activation = False Then Me.Show vbModeless
Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-but_Export_Click", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
'    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub

Private Sub but_PPHistory_Click()
    Dim a As Long, b As Long, sel As Long
    Dim seldesc As String, selpack
    Dim selCompCode As String
    Dim compet As String, bfound As Boolean
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    
    For a = 0 To box_Options.ListCount - 1
        If box_Options.Selected(a) = True Then
            sel = a
            seldesc = box_Options.List(a, 0)
            selpack = box_Options.List(a, 1)
            Exit For
        End If
    Next
    If seldesc = "" Then Exit Sub
    CCM_DefineActiveDataSets
    
    If Me.getCCM_UserDefinedState = True Then
        For a = 1 To 4
            If CCM_IsActiveDataSet(a) = True Then
                bfound = False
                If a = 1 Then
                    For b = 1 To UBound(CCM_UDWWSKU)
                        If CCM_UDWWSKU(b).CBA_COM_SKU_CompProdName = seldesc And CCM_UDWWSKU(b).CBA_COM_SKU_CompPacksize = selpack Then _
                                        bfound = True: selCompCode = CCM_UDWWSKU(b).CBA_COM_SKU_compcode: compet = "WW": Exit For
                    Next
                ElseIf a = 2 Then
                    For b = 1 To UBound(CCM_UDColesSKU)
                        If CCM_UDColesSKU(b).CBA_COM_SKU_CompProdName = seldesc And CCM_UDColesSKU(b).CBA_COM_SKU_CompPacksize = selpack Then _
                                        bfound = True: selCompCode = CCM_UDColesSKU(b).CBA_COM_SKU_compcode: compet = "Coles": Exit For
                    Next
                ElseIf a = 3 Then
                    For b = 1 To UBound(CCM_UDDMSKU)
                        If CCM_UDDMSKU(b).CBA_COM_SKU_CompProdName = seldesc And CCM_UDDMSKU(b).CBA_COM_SKU_CompPacksize = selpack Then _
                                        bfound = True: selCompCode = CCM_UDDMSKU(b).CBA_COM_SKU_compcode: compet = "DM": Exit For
                    Next
                ElseIf a = 4 Then
                    For b = 1 To UBound(CCM_UDFCSKU)
                        If CCM_UDFCSKU(b).CBA_COM_SKU_CompProdName = seldesc And CCM_UDFCSKU(b).CBA_COM_SKU_CompPacksize = selpack Then _
                                        bfound = True: selCompCode = CCM_UDFCSKU(b).CBA_COM_SKU_compcode: compet = "FC": Exit For
                    Next
                End If
                If bfound = True Then Exit For
            End If
        Next
    Else
        For a = 1 To 4
            If CCM_IsActiveDataSet(a) = True Then
                bfound = False
                If a = 1 Then
                    For b = 1 To UBound(CCM_WWSKU)
                        If CCM_WWSKU(b).CBA_COM_SKU_CompProdName = seldesc And CCM_WWSKU(b).CBA_COM_SKU_CompPacksize = selpack Then _
                                        bfound = True: selCompCode = CCM_WWSKU(b).CBA_COM_SKU_compcode: compet = "WW": Exit For
                    Next
                ElseIf a = 2 Then
                    For b = 1 To UBound(CCM_ColesSKU)
                        If CCM_ColesSKU(b).CBA_COM_SKU_CompProdName = seldesc And CCM_ColesSKU(b).CBA_COM_SKU_CompPacksize = selpack Then _
                                        bfound = True: selCompCode = CCM_ColesSKU(b).CBA_COM_SKU_compcode: compet = "Coles": Exit For
                    Next
                ElseIf a = 3 Then
                    For b = 1 To UBound(CCM_DMSKU)
                        If CCM_DMSKU(b).CBA_COM_SKU_CompProdName = seldesc And CCM_DMSKU(b).CBA_COM_SKU_CompPacksize = selpack Then _
                                        bfound = True: selCompCode = CCM_DMSKU(b).CBA_COM_SKU_compcode: compet = "DM": Exit For
                    Next
                ElseIf a = 4 Then
                    For b = 1 To UBound(CCM_FCSKU)
                        If CCM_FCSKU(b).CBA_COM_SKU_CompProdName = seldesc And CCM_FCSKU(b).CBA_COM_SKU_CompPacksize = selpack Then _
                                        bfound = True: selCompCode = CCM_FCSKU(b).CBA_COM_SKU_compcode: compet = "FC": Exit For
                    Next
                End If
                If bfound = True Then Exit For
            End If
        Next
    End If
    
    If compet = "DM" Or compet = "FC" Then
        MsgBox "Dan Murphy's and First Choice PPH Functionality is not avaliable currently within the Data Viewer. Please use the Price and Promotion History Report in the COMRADE Reporting Tool"
        Exit Sub
    End If
    
    If selCompCode <> "" And seldesc <> "" And compet <> "" And CCM_Statelookup <> "" Then
        Me.Hide
        CBA_COM_frm_Chart.formulatePPHChart selCompCode, seldesc, CCM_Statelookup, compet
        CBA_COM_frm_Chart.Show
        If CCM_Activation = False Then Me.Show
    End If
Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-but_PPHistory_Click", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
'    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Sub
Private Sub cb_MetaData_Change()
    CCM_Filter.MetaData = cb_MetaData.Value
    CCM_CreateCCMArrays
    CCM_PopulateData
End Sub

Sub UserForm_Initialize()
    Dim a As Long
    Me.StartUpPosition = 0
    Me.Top = Application.Top + (Application.Height / 2) - (Me.Height / 2)
    Me.Left = Application.Left + (Application.Width / 2) - (Me.Width / 2)
    'CBA_MouseScrolling.HookFormScroll Me



    
    Me.Hide
    If CCM_Runtime.CMM_getDefaultDataset > 2 Then
        For a = 0 To 3
            If a = 3 Then CCM_FormState(a) = True Else CCM_FormState(a) = False
        Next
    Else
        For a = 0 To 3
            If a = 0 Then CCM_FormState(a) = True Else CCM_FormState(a) = False
        Next
    End If
    Set CCM_Filter.Matches.WW = New Collection
    Set CCM_Filter.Matches.Coles = New Collection
    Set CCM_Filter.Matches.DM = New Collection
    Set CCM_Filter.Matches.FC = New Collection
    updateMatchCollections
    Me.cb_MetaData.AddItem ""
    Me.cb_MetaData.AddItem "Promo"
    Me.cb_MetaData.AddItem "EDLP"
    Me.cb_MetaData.AddItem "New"
    Me.cb_MetaData.AddItem "Low/No Stock"
    Me.cb_MetaData.AddItem "Points"
    
    
End Sub
Private Sub UserForm_Terminate()
'28690     Call but_Stop_Click
End Sub
Private Sub but_Stop_Click()
    Dim Cancel, a As Long
    Cancel = True
    'If CCM_UserDefinedState = True Then Erase CCM_WWSKU: Erase CCM_ColesSKU: Erase CCM_DMSKU: Erase CCM_FCSKU
    CCM_UserDefinedState = False
    'CBA_MouseScrolling.UnhookFormScroll
    For a = 0 To 3
        If a = 0 Then CCM_FormState(a) = True Else CCM_FormState(a) = False
    Next
  CCM_Runtime.CCM_dropMouseClassModule

    Unload Me
End Sub
Sub LoadupProcess()
    'CCM_FormState(0) mean Closed
    'CCM_FormState(1) mean Std Mode
    'CCM_FormState(2) mean Produce Mode
    'CCM_FormState(3) mean Alcohol Mode
    Dim bfound As Boolean
    Dim a As Long
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    CCM_Activation = True
    bfound = False

    CCM_DefineActiveDataSets
    For a = 1 To 4
        If CCM_IsActiveDataSet(a) = True Then
                bfound = True
                Exit For
        End If
    Next
    If bfound = False Then
        CBA_frm_SelectBaseData.Show
        If CMM_getDefaultDataset = 0 Then Exit Sub
        If CBA_BasicFunctions.isRunningSheetDisplayed = False Then CBA_BasicFunctions.CBA_Running
    End If
    If CCM_UserDefinedState = True Then setUDMode Else If CMM_getDefaultDataset < 3 Then setStdMode True Else setAlcoholMode True
    'CCM_MatchMap = CCM_QueryMatches
    If CCM_Statelookup = "" Then setCCM_StateLookup "National"
    If CCM_PriceType = "" Then setCCM_PriceType "Mode"
    If CCM_priceonshow = "" Then setCCM_PriceOnShow "ALL"
    CCM_DefineActiveDataSets
    If CCM_UserDefinedState = False Then
        bfound = False
        For a = 1 To 4
            If CCM_IsActiveDataSet(a) = True Then
                bfound = True
                Exit For
            End If
        Next
        If bfound = False And CBA_BasicFunctions.isRunningSheetDisplayed = False Then CBA_BasicFunctions.CBA_Running
'        If CCM_isactiveDataSet(1) = True Then Me.ck_WW.Enabled = True:  Me.ck_Coles.Enabled = True: Me.ck_DM.Enabled = False: Me.ck_FC.Enabled = False: Me.ck_UD.Enabled = False: Me.ck_UD.Value = False: Me.ck_WW.Value = True
'        If CCM_isactiveDataSet(2) = True Then Me.ck_WW.Enabled = True:  Me.ck_Coles.Enabled = True: Me.ck_DM.Enabled = False: Me.ck_FC.Enabled = False: Me.ck_UD.Enabled = False: Me.ck_UD.Value = False: Me.ck_WW.Value = True
'        ElseIf CCM_isactiveDataSet(3) = True Or CCM_isactiveDataSet(4) = True Then
'            Me.ck_WW.Enabled = False:  Me.ck_Coles.Enabled = False: Me.ck_DM.Enabled = True: Me.ck_FC.Enabled = True: Me.ck_UD.Enabled = False: Me.ck_UD.Value = False
'        End If
    
    Else
'        For a = 1 To 3
'            If a = 3 Then CCM_FormState(a) = True Else CCM_FormState(a) = False
'        Next
        For a = 1 To 4
            If CCM_IsActiveDataSet(a) = True Then
                    CCM_Runtime.CCM_setDefaultDataset a
                    Exit For
            End If
        Next

        Me.ck_WW.Enabled = False: Me.ck_Coles.Enabled = False: Me.ck_DM.Enabled = False: Me.ck_FC.Enabled = False: Me.ck_UD.Enabled = True: Me.ck_UD.Value = True
    End If
    CCM_CreateCCMArrays
    CCM_PopulateData
    
'    If CCM_isactiveDataSet(1) = True Then Me.ck_WW.Value = True: b = 1: Call ob_W_PL1_Click: Me.MultiPage1.Value = 1
'    If CCM_isactiveDataSet(2) = True Then Me.ck_Coles.Value = True: b = 1: If CCM_isactiveDataSet(1) = False Then Call ob_C_PL1_Click: Me.MultiPage1.Value = 0
'    If CCM_isactiveDataSet(3) = True Then Me.ck_DM.Value = True: b = 2: Call ob_DM1_Click: Me.MultiPage1.Value = 4
'    If CCM_isactiveDataSet(4) = True Then Me.ck_FC.Value = True: b = 2: If CCM_isactiveDataSet(3) = False Then Call ob_FC1_Click: Me.MultiPage1.Value = 4
'    If b = 1 Then setStdMode True
'    If b = 2 Then setAlcoholMode True
    updateAssignedMatches
    If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
    CCM_Activation = False
Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-LoadupProcess", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    'If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub
Private Sub CCM_PopulateData()
    Dim a As Long, settopop As Long
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
     
     'Debug.Print CCM_wwArr(1, 1)
    
    If CCM_UserDefinedState = False Or CCM_Activation = True Then
    If CCM_CompData.Competitor = "WW" Then
        If CCM_IsActiveDataSet(1) = False Then Call ck_WW_Click: ck_WW.Value = True
        settopop = 1
    ElseIf CCM_CompData.Competitor = "C" Then
        If CCM_IsActiveDataSet(2) = False Then Call ck_Coles_Click: ck_Coles.Value = True
        settopop = 2
    ElseIf CCM_CompData.Competitor = "DM" Then
        If CCM_IsActiveDataSet(3) = False Then Call ck_DM_Click: ck_DM.Value = True
        settopop = 3
    ElseIf CCM_CompData.Competitor = "FC" Then
        If CCM_IsActiveDataSet(4) = False Then Call ck_FC_Click: ck_FC.Value = True
        settopop = 4
    Else
        For a = 1 To 4
            If CCM_IsActiveDataSet(a) = True Then
                settopop = a
                Exit For
            End If
        Next
    End If
    Else
        If CCM_CompData.Competitor = "WW" Then
            settopop = 1
        ElseIf CCM_CompData.Competitor = "C" Then settopop = 2
        ElseIf CCM_CompData.Competitor = "DM" Then settopop = 3
        ElseIf CCM_CompData.Competitor = "FC" Then settopop = 4
        End If
    End If
    
    
    With Me.box_Options
       .Clear
       If settopop = 1 And CCM_IsActiveDataSet(1) = True Then
           .List = CBA_BasicFunctions.CBA_TransposeArray(CCM_WWArr)
       ElseIf settopop = 2 And CCM_IsActiveDataSet(2) = True Then .List = CBA_BasicFunctions.CBA_TransposeArray(CCM_ColesArr)
       ElseIf settopop = 3 And CCM_IsActiveDataSet(3) = True Then .List = CBA_BasicFunctions.CBA_TransposeArray(CCM_dmArr)
       ElseIf settopop = 4 And CCM_IsActiveDataSet(4) = True Then .List = CBA_BasicFunctions.CBA_TransposeArray(CCM_fcArr)
       End If
       DoEvents
    End With
Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-CCM_PopulateData", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
'    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub
Private Sub setCCM_StateLookup(ByVal State As String)
    CCM_Statelookup = State
        
        If State = "National" Then
            Me.but_National.Value = True
        ElseIf State = "VIC" Then
            Me.but_VIC.Value = True
        ElseIf State = "QLD" Then
            Me.but_QLD.Value = True
        ElseIf State = "SA" Then
            Me.but_SA.Value = True
        ElseIf State = "WA" Then
            Me.but_WA.Value = True
        ElseIf State = "NSW" Then
            Me.but_NSW.Value = True
        End If

End Sub
Private Sub setCCM_PriceType(ByVal pType As String)
    CCM_PriceType = pType
    If pType = "Highest" Then
        Me.pbut_Highest.Value = True
    ElseIf pType = "Lowest" Then
        Me.pbut_Lowest.Value = True
    ElseIf pType = "Mean" Then
        Me.pbut_Mean.Value = True
    ElseIf pType = "Median" Then
        Me.pbut_Median.Value = True
    ElseIf pType = "Mode" Then
        Me.pbut_Mode.Value = True
    End If
End Sub
Private Sub setCCM_PriceOnShow(ByVal pricetoshow As String)
    CCM_priceonshow = pricetoshow
    If pricetoshow = "Recent" Then
        Me.psbut_Recent.Value = True
    ElseIf pricetoshow = "NonPromo" Then
        Me.psbut_NonPromo.Value = True
    ElseIf pricetoshow = "Promo" Then
        Me.psbut_Promo.Value = True
    ElseIf pricetoshow = "MultiBuy" Then
        Me.psbut_MultiBuy.Value = True
    ElseIf pricetoshow = "ALL" Then
        Me.psbut_ALL.Value = True
    End If
End Sub
Private Sub CCM_CreateCCMArrays()
    Dim a As Long, b As Long, lNum As Long, bfound As Boolean, bFoundData As Boolean
    Dim FilterToAdd As Boolean, CancelLine As Boolean
    Dim CP, DefaultData
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    
    If CCM_Filter.Description <> "" Or CCM_Filter.Packsize <> "" Or CCM_Filter.Price <> 0 Or CCM_Filter.Special = True _
        Or Me.cbx_OnlyMatched = True Or CCM_Filter.MetaData <> "" Then
    '        If CCM_Filter.Matches.WW.Count > 0 Then WWmatchfilter = True Else WWmatchfilter = False
    '        If CCM_Filter.Matches.Coles.Count > 0 Then Cmatchfilter = True Else Cmatchfilter = False
    '        If CCM_Filter.Matches.DM.Count > 0 Then DMmatchfilter = True Else DMmatchfilter = False
    '        If CCM_Filter.Matches.FC.Count > 0 Then FCmatchfilter = True Else FCmatchfilter = False
        FilterToAdd = True
    Else
        FilterToAdd = False
    End If
    
    
    Call CCM_DefineActiveDataSets
    If CCM_UserDefinedState = False Then
        For a = 1 To 4
            If CCM_IsActiveDataSet(a) = True Then
                

                Select Case a
                    Case 1
                        If CCM_FormState(1) = True Or CCM_FormState(0) = True Or CCM_FormState(2) = True Then
                            bFoundData = True: lNum = 0
                            If IsEmpty(CCM_WWArr) = False Then Erase CCM_WWArr
                            If IsEmpty(CCM_ColesArr) = False Then Erase CCM_ColesArr
                            If IsEmpty(CCM_dmArr) = False Then Erase CCM_dmArr
                            If IsEmpty(CCM_fcArr) = False Then Erase CCM_fcArr
                            
                            ReDim CCM_WWArr(1 To 4, 1 To 1)
                            For b = LBound(CCM_WWSKU) To UBound(CCM_WWSKU)
                                bfound = True
                                If Me.cbx_OnlyMatched = True Then
                                    If CCM_Filter.Matches.WW.Count = 0 Then Exit For
                                    bfound = False
                                    For Each CP In CCM_Filter.Matches.WW
                                        If CP = CCM_WWSKU(b).CBA_COM_SKU_compcode Then
                                            bfound = True
                                            Exit For
                                        ElseIf CP > CCM_WWSKU(b).CBA_COM_SKU_compcode Then
                                            Exit For
                                        End If
                                    Next
                                End If
                                If bfound = True Then
                                    lNum = lNum + 1: CancelLine = False
                                    ReDim Preserve CCM_WWArr(1 To 4, 1 To lNum)
                                    CCM_WWArr(1, lNum) = CCM_WWSKU(b).CBA_COM_SKU_CompProdName
                                    CCM_WWArr(2, lNum) = CCM_WWSKU(b).CBA_COM_SKU_CompPacksize
                                    CCM_WWArr(3, lNum) = CCM_WWSKU(b).getPriceData(CCM_Statelookup, CCM_PriceType, CCM_priceonshow)
                                    CCM_WWArr(4, lNum) = CCM_WWSKU(b).getMetaData(CCM_Statelookup, CCM_priceonshow)
                                    If FilterToAdd = True Then
                                        If CCM_Filter.Description <> "" And InStr(1, LCase(CCM_WWArr(1, lNum)), LCase(CCM_Filter.Description)) = 0 Then
                                            CancelLine = True
                                        ElseIf CCM_Filter.Packsize <> "" And CCM_Filter.Packsize <> CCM_WWArr(2, lNum) Then CancelLine = True
                                        ElseIf CCM_Filter.Price <> 0 And CCM_Filter.Price <> CCM_WWArr(3, lNum) Then CancelLine = True
                                        ElseIf CCM_Filter.MetaData <> "" And InStr(1, CCM_WWArr(4, lNum), CCM_Filter.MetaData) = 0 Then CancelLine = True
                                        End If
                                    End If
                                    If CancelLine = True Or CCM_WWArr(3, lNum) = 0 Then
                                        lNum = lNum - 1
                                        If lNum > 0 Then ReDim Preserve CCM_WWArr(1 To 4, 1 To lNum)
                                    End If
                                
                                End If
                            Next
                            Me.ck_WW.Value = True
                        End If
                        
                    Case 2
                        If CCM_FormState(1) = True Or CCM_FormState(0) = True Or CCM_FormState(2) Then
                        bFoundData = True: lNum = 0
                        If IsEmpty(CCM_ColesArr) = False Then Erase CCM_ColesArr
                        ReDim CCM_ColesArr(1 To 4, 1 To 1)
                        For b = LBound(CCM_ColesSKU) To UBound(CCM_ColesSKU)
                            bfound = True
                            If Me.cbx_OnlyMatched = True Then
                                If CCM_Filter.Matches.Coles.Count = 0 Then Exit For
                                bfound = False
                                For Each CP In CCM_Filter.Matches.Coles
                                    If CP = CCM_ColesSKU(b).CBA_COM_SKU_compcode Then
                                        bfound = True
                                        Exit For
                                    ElseIf Mid(CP, 1, 1) > Mid(CCM_ColesSKU(b).CBA_COM_SKU_compcode, 1, 1) Then
                                        Exit For
                                    End If
                                Next
                            End If
                            If bfound = True Then
                                lNum = lNum + 1: CancelLine = False
                                ReDim Preserve CCM_ColesArr(1 To 4, 1 To lNum)
                                CCM_ColesArr(1, lNum) = CCM_ColesSKU(b).CBA_COM_SKU_CompProdName
                                CCM_ColesArr(2, lNum) = CCM_ColesSKU(b).CBA_COM_SKU_CompPacksize
                                CCM_ColesArr(3, lNum) = CCM_ColesSKU(b).getPriceData(CCM_Statelookup, CCM_PriceType, CCM_priceonshow)
                                CCM_ColesArr(4, lNum) = CCM_ColesSKU(b).getMetaData(CCM_Statelookup, CCM_priceonshow)
                                If FilterToAdd = True Then
                                    If CCM_Filter.Description <> "" And InStr(1, LCase(CCM_ColesArr(1, lNum)), LCase(CCM_Filter.Description)) = 0 Then
                                        CancelLine = True
                                    ElseIf CCM_Filter.Packsize <> "" And CCM_Filter.Packsize <> CCM_ColesArr(2, lNum) Then CancelLine = True
                                    ElseIf CCM_Filter.Price <> 0 And CCM_Filter.Price <> CCM_ColesArr(3, lNum) Then CancelLine = True
                                    ElseIf CCM_Filter.MetaData <> "" And InStr(1, CCM_ColesArr(4, lNum), CCM_Filter.MetaData) = 0 Then CancelLine = True
                                    End If
                                End If
                                If CancelLine = True Or CCM_ColesArr(3, lNum) = 0 Then
                                    lNum = lNum - 1
                                    If lNum > 0 Then ReDim Preserve CCM_ColesArr(1 To 4, 1 To lNum)
                                End If
                            
                            End If
                            Next
                            Me.ck_Coles.Value = True
                            End If
                    Case 3
                        If CCM_FormState(3) = True Or CCM_FormState(0) = True Then
                        bFoundData = True: lNum = 0
                        If IsEmpty(CCM_dmArr) = False Then Erase CCM_dmArr
                        ReDim CCM_dmArr(1 To 4, 1 To 1)
                        For b = LBound(CCM_DMSKU) To UBound(CCM_DMSKU)
                            bfound = True
                            If Me.cbx_OnlyMatched = True Then
                                If CCM_Filter.Matches.DM.Count = 0 Then Exit For
                                bfound = False
                                For Each CP In CCM_Filter.Matches.DM
                                    If CP = CCM_DMSKU(b).CBA_COM_SKU_compcode Then
                                        bfound = True
                                        Exit For
                                    ElseIf CP > CCM_DMSKU(b).CBA_COM_SKU_compcode Then
                                        Exit For
                                    End If
                                Next
                            End If
                            If bfound = True Then
                                lNum = lNum + 1: CancelLine = False
                                ReDim Preserve CCM_dmArr(1 To 4, 1 To lNum)
                                CCM_dmArr(1, lNum) = CCM_DMSKU(b).CBA_COM_SKU_CompProdName
                                CCM_dmArr(2, lNum) = CCM_DMSKU(b).CBA_COM_SKU_CompPacksize
                                CCM_dmArr(3, lNum) = CCM_DMSKU(b).getPriceData(CCM_Statelookup, CCM_PriceType, CCM_priceonshow)
                                CCM_dmArr(4, lNum) = CCM_DMSKU(b).getMetaData(CCM_Statelookup, CCM_priceonshow)
                                If InStr(1, CCM_dmArr(4, lNum), "Promo") > 0 Then
                                    If CCM_dmArr(3, lNum) = CCM_DMSKU(b).getPriceData(CCM_Statelookup, CCM_PriceType, "Promo") Then
                                    Else
                                        CCM_dmArr(4, lNum) = Mid(CCM_dmArr(4, lNum), 1, InStr(1, CCM_dmArr(4, lNum), "Promo") - 1) & Mid(CCM_dmArr(4, lNum), InStr(1, CCM_dmArr(4, lNum), "Promo") + 5, 999)
                                    End If
                                End If
                                If FilterToAdd = True Then
                                    If CCM_Filter.Description <> "" And InStr(1, LCase(CCM_dmArr(1, lNum)), LCase(CCM_Filter.Description)) = 0 Then
                                        CancelLine = True
                                    ElseIf CCM_Filter.Packsize <> "" And CCM_Filter.Packsize <> CCM_dmArr(2, lNum) Then CancelLine = True
                                    ElseIf CCM_Filter.Price <> 0 And CCM_Filter.Price <> CCM_dmArr(3, lNum) Then CancelLine = True
                                    ElseIf CCM_Filter.MetaData <> "" And InStr(1, CCM_dmArr(4, lNum), CCM_Filter.MetaData) = 0 Then CancelLine = True
                                    End If
                                End If
                                If CancelLine = True Or CCM_dmArr(3, lNum) = 0 Then
                                    lNum = lNum - 1
                                    If lNum > 0 Then ReDim Preserve CCM_dmArr(1 To 4, 1 To lNum)
                                End If
                            
                            End If
                        Next
                            Me.ck_DM.Value = True
                        End If
                    Case 4
                        If CCM_FormState(3) = True Or CCM_FormState(0) = True Then
                        bFoundData = True: lNum = 0
                        If IsEmpty(CCM_fcArr) = False Then Erase CCM_fcArr
                        ReDim CCM_fcArr(1 To 4, 1 To 1)
                        For b = LBound(CCM_FCSKU) To UBound(CCM_FCSKU)
                            bfound = True
                            If Me.cbx_OnlyMatched = True Then
                                If CCM_Filter.Matches.FC.Count = 0 Then Exit For
                                bfound = False
                                For Each CP In CCM_Filter.Matches.FC
                                    If CP = CCM_FCSKU(b).CBA_COM_SKU_compcode Then
                                        bfound = True
                                        Exit For
                                    ElseIf CP > CCM_FCSKU(b).CBA_COM_SKU_compcode Then
                                        Exit For
                                    End If
                                Next
                            End If
                            If bfound = True Then
                                lNum = lNum + 1: CancelLine = False
                                ReDim Preserve CCM_fcArr(1 To 4, 1 To lNum)
                                CCM_fcArr(1, lNum) = CCM_FCSKU(b).CBA_COM_SKU_CompProdName
                                CCM_fcArr(2, lNum) = CCM_FCSKU(b).CBA_COM_SKU_CompPacksize
                                CCM_fcArr(3, lNum) = CCM_FCSKU(b).getPriceData(CCM_Statelookup, CCM_PriceType, CCM_priceonshow)
                                CCM_fcArr(4, lNum) = CCM_FCSKU(b).getMetaData(CCM_Statelookup, CCM_priceonshow)
                                If FilterToAdd = True Then
                                    If CCM_Filter.Description <> "" And InStr(1, LCase(CCM_fcArr(1, lNum)), LCase(CCM_Filter.Description)) = 0 Then
                                        CancelLine = True
                                    ElseIf CCM_Filter.Packsize <> "" And CCM_Filter.Packsize <> CCM_fcArr(2, lNum) Then CancelLine = True
                                    ElseIf CCM_Filter.Price <> 0 And CCM_Filter.Price <> CCM_fcArr(3, lNum) Then CancelLine = True
                                    ElseIf CCM_Filter.MetaData <> "" And InStr(1, CCM_fcArr(4, lNum), CCM_Filter.MetaData) = 0 Then CancelLine = True
                                    End If
                                End If
                                If CancelLine = True Or CCM_fcArr(3, lNum) = 0 Then
                                    lNum = lNum - 1
                                    If lNum > 0 Then ReDim Preserve CCM_fcArr(1 To 4, 1 To lNum)
                                End If
                            
                            End If
                        Next
                            Me.ck_FC.Value = True
                        End If
                End Select
            End If
        Next
    Else
        For a = 1 To 4
            If CCM_IsActiveDataSet(a) = True Then
                Select Case a
                    Case 1
                        If CCM_FormState(1) = True Or CCM_FormState(0) = True Or CCM_FormState(2) Then
                        bFoundData = True: lNum = 0
                        If IsEmpty(CCM_WWArr) = False Then Erase CCM_WWArr
                        ReDim CCM_WWArr(1 To 4, 1 To 1)
                        For b = LBound(CCM_UDWWSKU) To UBound(CCM_UDWWSKU)
                            bfound = True
                            If Me.cbx_OnlyMatched = True Then
                            If CCM_Filter.Matches.WW.Count = 0 Then Exit For
                                bfound = False
                                For Each CP In CCM_Filter.Matches.WW
                                    If CP = CCM_UDWWSKU(b).CBA_COM_SKU_compcode Then
                                        bfound = True
                                        Exit For
                                    ElseIf CP > CCM_UDWWSKU(b).CBA_COM_SKU_compcode Then
                                        Exit For
                                    End If
                                Next
                            End If
                            If bfound = True Then
                                lNum = lNum + 1: CancelLine = False
                                ReDim Preserve CCM_WWArr(1 To 4, 1 To lNum)
                                CCM_WWArr(1, lNum) = CCM_UDWWSKU(b).CBA_COM_SKU_CompProdName
                                CCM_WWArr(2, lNum) = CCM_UDWWSKU(b).CBA_COM_SKU_CompPacksize
                                CCM_WWArr(3, lNum) = CCM_UDWWSKU(b).getPriceData(CCM_Statelookup, CCM_PriceType, CCM_priceonshow)
                                CCM_WWArr(4, lNum) = CCM_UDWWSKU(b).getMetaData(CCM_Statelookup, CCM_priceonshow)
                                If FilterToAdd = True Then
                                    If CCM_Filter.Description <> "" And InStr(1, LCase(CCM_WWArr(1, lNum)), LCase(CCM_Filter.Description)) = 0 Then
                                        CancelLine = True
                                    ElseIf CCM_Filter.Packsize <> "" And CCM_Filter.Packsize <> CCM_WWArr(2, lNum) Then CancelLine = True
                                    ElseIf CCM_Filter.Price <> 0 And CCM_Filter.Price <> CCM_WWArr(3, lNum) Then CancelLine = True
                                    ElseIf CCM_Filter.MetaData <> "" And InStr(1, CCM_WWArr(4, lNum), CCM_Filter.MetaData) = 0 Then CancelLine = True
                                    End If
                                End If
                                If CancelLine = True Or CCM_WWArr(3, lNum) = 0 Then
                                    lNum = lNum - 1
                                    If lNum > 0 Then ReDim Preserve CCM_WWArr(1 To 4, 1 To lNum)
                                End If
                            
                            End If
                            
                        Next
                            Me.ck_WW.Value = True
                        End If
                    Case 2
                        If CCM_FormState(1) = True Or CCM_FormState(0) = True Or CCM_FormState(2) Then
                        bFoundData = True: lNum = 0
                        If IsEmpty(CCM_ColesArr) = False Then Erase CCM_ColesArr
                        ReDim CCM_ColesArr(1 To 4, 1 To 1)
                        For b = LBound(CCM_UDColesSKU) To UBound(CCM_UDColesSKU)
                            bfound = True
                            If Me.cbx_OnlyMatched = True Then
                            If CCM_Filter.Matches.Coles.Count = 0 Then Exit For
                                bfound = False
                                For Each CP In CCM_Filter.Matches.Coles
                                    If CP = CCM_UDColesSKU(b).CBA_COM_SKU_compcode Then
                                        bfound = True
                                        Exit For
                                    ElseIf Mid(CP, 1, 1) > Mid(CCM_UDColesSKU(b).CBA_COM_SKU_compcode, 1, 1) Then
                                        Exit For
                                    End If
                                Next
                            End If
                            If bfound = True Then
                                lNum = lNum + 1: CancelLine = False
                                ReDim Preserve CCM_ColesArr(1 To 4, 1 To lNum)
                                'Debug.Print CCM_UDColesSKU(b).CBA_COM_SKU_compcode
                                CCM_ColesArr(1, lNum) = CCM_UDColesSKU(b).CBA_COM_SKU_CompProdName
                                CCM_ColesArr(2, lNum) = CCM_UDColesSKU(b).CBA_COM_SKU_CompPacksize
                                CCM_ColesArr(3, lNum) = CCM_UDColesSKU(b).getPriceData(CCM_Statelookup, CCM_PriceType, CCM_priceonshow)
                                CCM_ColesArr(4, lNum) = CCM_UDColesSKU(b).getMetaData(CCM_Statelookup, CCM_priceonshow)
                                If FilterToAdd = True Then
                                    If CCM_Filter.Description <> "" And InStr(1, LCase(CCM_ColesArr(1, lNum)), LCase(CCM_Filter.Description)) = 0 Then
                                        CancelLine = True
                                    ElseIf CCM_Filter.Packsize <> "" And CCM_Filter.Packsize <> CCM_ColesArr(2, lNum) Then CancelLine = True
                                    ElseIf CCM_Filter.Price <> 0 And CCM_Filter.Price <> CCM_ColesArr(3, lNum) Then CancelLine = True
                                    ElseIf CCM_Filter.MetaData <> "" And InStr(1, CCM_ColesArr(4, lNum), CCM_Filter.MetaData) = 0 Then CancelLine = True
                                    End If
                                End If
                                If CancelLine = True Or CCM_ColesArr(3, lNum) = 0 Then
                                    lNum = lNum - 1
                                    If lNum > 0 Then ReDim Preserve CCM_ColesArr(1 To 4, 1 To lNum)
                                End If
                            End If
                        Next
                            Me.ck_Coles.Value = True
                        End If
                    Case 3
                        If CCM_FormState(3) = True Or CCM_FormState(0) = True Then
                        bFoundData = True: lNum = 0: lNum = 0
                        If IsEmpty(CCM_dmArr) = False Then Erase CCM_dmArr
                        ReDim CCM_dmArr(1 To 4, 1 To 1)
                        For b = LBound(CCM_UDDMSKU) To UBound(CCM_UDDMSKU)
                            bfound = True
                            If Me.cbx_OnlyMatched = True Then
                            If CCM_Filter.Matches.DM.Count = 0 Then Exit For
                                bfound = False
                                For Each CP In CCM_Filter.Matches.DM
                                    If CP = CCM_UDDMSKU(b).CBA_COM_SKU_compcode Then
                                        bfound = True
                                        Exit For
                                    ElseIf CP > CCM_UDDMSKU(b).CBA_COM_SKU_compcode Then
                                        Exit For
                                    End If
                                Next
                            End If
                            If bfound = True Then
                                lNum = lNum + 1: CancelLine = False
                                ReDim Preserve CCM_dmArr(1 To 4, 1 To lNum)
                                CCM_dmArr(1, lNum) = CCM_UDDMSKU(b).CBA_COM_SKU_CompProdName
                                CCM_dmArr(2, lNum) = CCM_UDDMSKU(b).CBA_COM_SKU_CompPacksize
                                CCM_dmArr(3, lNum) = CCM_UDDMSKU(b).getPriceData(CCM_Statelookup, CCM_PriceType, CCM_priceonshow)
                                CCM_dmArr(4, lNum) = CCM_UDDMSKU(b).getMetaData(CCM_Statelookup, CCM_priceonshow)
                                If InStr(1, CCM_dmArr(4, lNum), "Promo") > 0 Then
                                    If CCM_dmArr(3, lNum) = CCM_UDDMSKU(b).getPriceData(CCM_Statelookup, CCM_PriceType, "Promo") Then
                                    Else
                                        CCM_dmArr(4, lNum) = Mid(CCM_dmArr(4, lNum), 1, InStr(1, CCM_dmArr(4, lNum), "Promo") - 1) & Mid(CCM_dmArr(4, lNum), InStr(1, CCM_dmArr(4, lNum), "Promo") + 5, 999)
                                    End If
                                End If
                                If FilterToAdd = True Then
                                    If CCM_Filter.Description <> "" And InStr(1, LCase(CCM_dmArr(1, lNum)), LCase(CCM_Filter.Description)) = 0 Then
                                        CancelLine = True
                                    ElseIf CCM_Filter.Packsize <> "" And CCM_Filter.Packsize <> CCM_dmArr(2, lNum) Then CancelLine = True
                                    ElseIf CCM_Filter.Price <> 0 And CCM_Filter.Price <> CCM_dmArr(3, lNum) Then CancelLine = True
                                    ElseIf CCM_Filter.MetaData <> "" And InStr(1, CCM_dmArr(4, lNum), CCM_Filter.MetaData) = 0 Then CancelLine = True
                                    End If
                                End If
                                If CancelLine = True Or CCM_dmArr(3, lNum) = 0 Then
                                    lNum = lNum - 1
                                    If lNum > 0 Then ReDim Preserve CCM_dmArr(1 To 4, 1 To lNum)
                                End If
                            
                            End If
                        Next
                            Me.ck_DM.Value = True
                        End If
                    Case 4
                        If CCM_FormState(3) = True Or CCM_FormState(0) = True Then
                        bFoundData = True: lNum = 0: lNum = 0
                        If IsEmpty(CCM_fcArr) = False Then Erase CCM_fcArr
                        ReDim CCM_fcArr(1 To 4, 1 To 1)
                        For b = LBound(CCM_UDFCSKU) To UBound(CCM_UDFCSKU)
                            bfound = True
                            If Me.cbx_OnlyMatched = True Then
                                If CCM_Filter.Matches.FC.Count = 0 Then Exit For
                                bfound = False
                                For Each CP In CCM_Filter.Matches.FC
                                    If CP = CCM_UDFCSKU(b).CBA_COM_SKU_compcode Then
                                        bfound = True
                                        Exit For
                                    ElseIf CP > CCM_UDFCSKU(b).CBA_COM_SKU_compcode Then
                                        Exit For
                                    End If
                                Next
                            End If
                            If bfound = True Then
                                lNum = lNum + 1: CancelLine = False
                                ReDim Preserve CCM_fcArr(1 To 4, 1 To lNum)
                                CCM_fcArr(1, lNum) = CCM_UDFCSKU(b).CBA_COM_SKU_CompProdName
                                CCM_fcArr(2, lNum) = CCM_UDFCSKU(b).CBA_COM_SKU_CompPacksize
                                CCM_fcArr(3, lNum) = CCM_UDFCSKU(b).getPriceData(CCM_Statelookup, CCM_PriceType, CCM_priceonshow)
                                CCM_fcArr(4, lNum) = CCM_UDFCSKU(b).getMetaData(CCM_Statelookup, CCM_priceonshow)
                                If FilterToAdd = True Then
                                    If CCM_Filter.Description <> "" And InStr(1, LCase(CCM_fcArr(1, lNum)), LCase(CCM_Filter.Description)) = 0 Then
                                        CancelLine = True
                                    ElseIf CCM_Filter.Packsize <> "" And CCM_Filter.Packsize <> CCM_fcArr(2, lNum) Then CancelLine = True
                                    ElseIf CCM_Filter.Price <> 0 And CCM_Filter.Price <> CCM_fcArr(3, lNum) Then CancelLine = True
                                    ElseIf CCM_Filter.MetaData <> "" And InStr(1, CCM_fcArr(4, lNum), CCM_Filter.MetaData) = 0 Then CancelLine = True
                                    End If
                                End If
                                If CancelLine = True Or CCM_fcArr(3, lNum) = 0 Then
                                    lNum = lNum - 1
                                    If lNum > 0 Then ReDim Preserve CCM_fcArr(1 To 4, 1 To lNum)
                                End If
                            
                            End If
                        Next
                            Me.ck_FC.Value = True
                        End If
                End Select
            End If
        Next
    End If
    
    If bFoundData = False Then
        If CCM_UserDefinedState = True Then Exit Sub
        DefaultData = CMM_getDefaultDataset

'        thisbuyer = CBA_COM_Runtime.getCCMBuyer
'        If thisbuyer <> "" Then
''            If thisbuyer = "Bowyer" Or thisbuyer = "Handley" Then
''                CCM_DMSKU = CBA_COM_SetupSKUArray.CBA_SetupSKUArray("DM")
''            Else
'                CCM_WWSKU = CBA_COM_SetupSKUArray.CBA_SetupSKUArray("WW")
''            End If
'        Else
'            thisCG = CBA_COM_Runtime.getCCMCG
'            If thisCG <> 0 Then
''                If thisCG < 5 Then
''                    CCM_DMSKU = CBA_COM_SetupSKUArray.CBA_SetupSKUArray("DM")
''                Else
'                    CCM_WWSKU = CBA_COM_SetupSKUArray.CBA_SetupSKUArray("WW")
''                End If
'            End If
'        End If
        If DefaultData > 2 And CCM_FormState(3) = False And CCM_UserDefinedState = False Then setAlcoholMode True
        
'            For a = 0 To 3
'                If a = 3 Then CCM_FormState(a) = True Else CCM_FormState(a) = False
'            Next
'        End If
        If CCM_FormState(0) = True Or CCM_FormState(1) = True Or CCM_FormState(2) = True Then
            If DefaultData = 1 Then
                CCM_WWSKU = CBA_COM_SetupSKUArray.CBA_SetupSKUArray("WW")
            ElseIf DefaultData = 2 Then
                CCM_ColesSKU = CBA_COM_SetupSKUArray.CBA_SetupSKUArray("C")
            End If
        ElseIf CCM_FormState(3) = True Then
            If DefaultData = 3 Then
                CCM_DMSKU = CBA_COM_SetupSKUArray.CBA_SetupSKUArray("DM")
            ElseIf DefaultData = 4 Then
                CCM_FCSKU = CBA_COM_SetupSKUArray.CBA_SetupSKUArray("FC")
            End If
        End If


        Call CCM_CreateCCMArrays
    End If
Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-CCM_CreateCCMArrays", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    'If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Sub
Private Sub CCM_DefineActiveDataSets()
    Dim a As Long
    On Error GoTo Err_Routine
    CBA_ErrTag = ""

    For a = 1 To 4
        CCM_IsActiveDataSet(a) = True
    Next
    If CCM_UserDefinedState = False Then
        On Error Resume Next
        If UBound(CCM_WWSKU) < 1 Then
            CCM_IsActiveDataSet(1) = False
        End If
        Err.Clear
        If UBound(CCM_ColesSKU) < 1 Then
            CCM_IsActiveDataSet(2) = False
        End If
        Err.Clear
        If UBound(CCM_DMSKU) < 1 Then
            CCM_IsActiveDataSet(3) = False
        End If
        Err.Clear
        If UBound(CCM_FCSKU) < 1 Then
            CCM_IsActiveDataSet(4) = False
        End If
        Err.Clear
        On Error GoTo 0
    Else
        On Error Resume Next
        If UBound(CCM_UDWWSKU) < 1 Then
            CCM_IsActiveDataSet(1) = False
        End If
        Err.Clear
        If UBound(CCM_UDColesSKU) < 1 Then
            CCM_IsActiveDataSet(2) = False
        End If
        Err.Clear
        If UBound(CCM_UDDMSKU) < 1 Then
            CCM_IsActiveDataSet(3) = False
        End If
        Err.Clear
        If UBound(CCM_UDFCSKU) < 1 Then
            CCM_IsActiveDataSet(4) = False
        End If
        Err.Clear
        On Error GoTo 0
    End If
Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-CCM_DefineActiveDataSets", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
'    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub

Private Sub CCM_SetCurrentSelection()
    Dim lNum As Long
    For lNum = 0 To Me.box_Options.ListCount - 1
        If Me.box_Options.Selected(lNum) = True Then
            CCM_CurSel.Name = Me.box_Options.List(lNum, 0)
            CCM_CurSel.Pack = Me.box_Options.List(lNum, 1)
            Exit For
        End If
    Next
End Sub
Private Sub CCM_LoadCurrentSelection()
    Dim lNum As Long
    
    For lNum = 0 To Me.box_Options.ListCount - 1
        If CCM_CurSel.Name = Me.box_Options.List(lNum, 0) And CCM_CurSel.Pack = Me.box_Options.List(lNum, 1) Then
            Me.box_Options.Selected(lNum) = True
            Exit For
        End If
    Next
End Sub

Private Sub but_National_Click()
    If CCM_Activation = False Then CCM_SetCurrentSelection: CCM_Statelookup = "National": CCM_CreateCCMArrays: CCM_PopulateData: CCM_LoadCurrentSelection
End Sub
Private Sub but_NSW_Click()
    If CCM_Activation = False Then CCM_SetCurrentSelection: CCM_Statelookup = "NSW": CCM_CreateCCMArrays: CCM_PopulateData: CCM_LoadCurrentSelection
End Sub
Private Sub but_VIC_Click()
    If CCM_Activation = False Then CCM_SetCurrentSelection: CCM_Statelookup = "VIC": CCM_CreateCCMArrays: CCM_PopulateData: CCM_LoadCurrentSelection
End Sub
Private Sub but_WA_Click()
    If CCM_Activation = False Then CCM_SetCurrentSelection: CCM_Statelookup = "WA": CCM_CreateCCMArrays: CCM_PopulateData: CCM_LoadCurrentSelection
End Sub
Private Sub but_QLD_Click()
    If CCM_Activation = False Then CCM_SetCurrentSelection: CCM_Statelookup = "QLD": CCM_CreateCCMArrays: CCM_PopulateData: CCM_LoadCurrentSelection
End Sub
Private Sub but_SA_Click()
    If CCM_Activation = False Then CCM_SetCurrentSelection: CCM_Statelookup = "SA": CCM_CreateCCMArrays: CCM_PopulateData: CCM_LoadCurrentSelection
End Sub
Private Sub pbut_Highest_Click()
    If CCM_Activation = False Then CCM_SetCurrentSelection: CCM_PriceType = "Highest": CCM_CreateCCMArrays: CCM_PopulateData: CCM_LoadCurrentSelection
End Sub
Private Sub pbut_Lowest_Click()
    If CCM_Activation = False Then CCM_SetCurrentSelection: CCM_PriceType = "Lowest": CCM_CreateCCMArrays: CCM_PopulateData: CCM_LoadCurrentSelection
End Sub
Private Sub pbut_Mean_Click()
    If CCM_Activation = False Then CCM_SetCurrentSelection: CCM_PriceType = "Mean": CCM_CreateCCMArrays: CCM_PopulateData: CCM_LoadCurrentSelection
End Sub
Private Sub pbut_Median_Click()
    If CCM_Activation = False Then CCM_SetCurrentSelection: CCM_PriceType = "Median": CCM_CreateCCMArrays: CCM_PopulateData: CCM_LoadCurrentSelection
End Sub
Private Sub pbut_Mode_Click()
    If CCM_Activation = False Then CCM_SetCurrentSelection: CCM_PriceType = "Mode": CCM_CreateCCMArrays: CCM_PopulateData: CCM_LoadCurrentSelection
End Sub
Private Sub cbx_OnlyMatched_Click()
    If Me.cbx_OnlyMatched = False Then
        Set CCM_Filter.Matches.WW = New Collection
        Set CCM_Filter.Matches.Coles = New Collection
        Set CCM_Filter.Matches.DM = New Collection
        Set CCM_Filter.Matches.FC = New Collection
    Else
        updateMatchCollections
    End If
CCM_CreateCCMArrays
CCM_PopulateData

End Sub
Private Sub psbut_ALL_Click()
    If CCM_Activation = False Then CCM_SetCurrentSelection: CCM_priceonshow = "ALL": CCM_CreateCCMArrays: CCM_PopulateData: CCM_LoadCurrentSelection
End Sub
Private Sub psbut_NonPromo_Click()
    If CCM_Activation = False Then CCM_SetCurrentSelection: CCM_priceonshow = "NonPromo": CCM_CreateCCMArrays: CCM_PopulateData: CCM_LoadCurrentSelection
End Sub
Private Sub psbut_Promo_Click()
    If CCM_Activation = False Then CCM_SetCurrentSelection: CCM_priceonshow = "Promo": CCM_CreateCCMArrays: CCM_PopulateData: CCM_LoadCurrentSelection
End Sub
Private Sub psbut_Recent_Click()
    If CCM_Activation = False Then CCM_SetCurrentSelection: CCM_priceonshow = "Recent": CCM_CreateCCMArrays: CCM_PopulateData: CCM_LoadCurrentSelection
End Sub
Private Sub psbut_Multibuy_Click()
    If CCM_Activation = False Then CCM_SetCurrentSelection: CCM_priceonshow = "MultiBuy": CCM_CreateCCMArrays: CCM_PopulateData: CCM_LoadCurrentSelection
End Sub
Sub setCCM_UserDefinedState(ByVal isuserdefined As Boolean)
    If CCM_Activation = False Then CCM_UserDefinedState = isuserdefined
End Sub
Function getCCM_UserDefinedState() As Boolean
    getCCM_UserDefinedState = CCM_UserDefinedState
End Function
Private Sub ck_Coles_Click()
Dim lRet As Long
    If CCM_Activation = False Then
    Call CCM_DefineActiveDataSets
    If CCM_IsActiveDataSet(2) = False Then
        Me.Hide
        If CBA_BasicFunctions.isRunningSheetDisplayed = False Then CBA_BasicFunctions.CBA_Running
        CCM_ColesSKU = CBA_COM_SetupSKUArray.CBA_SetupSKUArray("C")
        If CCM_Activation = False Then Me.Show vbModeless
    Else
        If ck_Coles.Value = False Then
            lRet = MsgBox("Would you like to drop the Coles data in order to increase performance?" & Chr(10) & Chr(10) & "If you do then you will need to requery the Coles SKU Objects again", vbYesNo)
            If lRet = 6 Then
                Erase CCM_ColesSKU
            Else
                ck_Coles.Value = True
            End If
        End If
    End If
    End If
End Sub
Private Sub ck_WW_Click()
Dim lRet As Long
    If CCM_Activation = False Then
    Call CCM_DefineActiveDataSets
    If CCM_IsActiveDataSet(1) = False Then
        Me.Hide
        CCM_WWSKU = CBA_COM_SetupSKUArray.CBA_SetupSKUArray("WW")
        If CCM_Activation = False Then Me.Show vbModeless
    Else
        If ck_WW.Value = False Then
            lRet = MsgBox("Would you like to drop the Woolworths data in order to increase performance?" & Chr(10) & Chr(10) & "If you do then you will need to requery the Woolworths SKU Objects again", vbYesNo)
            If lRet = 6 Then
                Erase CCM_WWSKU
            Else
                ck_WW.Value = True
            End If
        End If
    End If
    End If
End Sub
Private Sub ck_DM_Click()
Dim lRet As Long
    If CCM_Activation = False Then
    Call CCM_DefineActiveDataSets
    If CCM_IsActiveDataSet(3) = False Then
        Me.Hide
        CCM_DMSKU = CBA_COM_SetupSKUArray.CBA_SetupSKUArray("DM")
        If CCM_Activation = False Then Me.Show vbModeless
    Else
        If ck_DM.Value = False Then
            lRet = MsgBox("Would you like to drop the Dan Murphys data in order to increase performance?" & Chr(10) & Chr(10) & "If you do then you will need to requery the Coles Dan Murphys SKU Objects again", vbYesNo)
            If lRet = 6 Then
                Erase CCM_DMSKU
            Else
                ck_DM.Value = True
            End If
        End If
    End If
    End If
End Sub
Private Sub ck_FC_Click()
Dim lRet As Long
    If CCM_Activation = False Then
    Call CCM_DefineActiveDataSets
    If CCM_IsActiveDataSet(4) = False Then
        Me.Hide
        CCM_FCSKU = CBA_COM_SetupSKUArray.CBA_SetupSKUArray("FC")
        If CCM_Activation = False Then Me.Show vbModeless
    Else
        If ck_FC.Value = False Then
            lRet = MsgBox("Would you like to drop the First Choice data in order to increase performance?" & Chr(10) & Chr(10) & "If you do then you will need to requery the First Choice SKU Objects again", vbYesNo)
            If lRet = 6 Then
                Erase CCM_FCSKU
            Else
                ck_FC.Value = True
            End If
        End If
    End If
    End If
End Sub

Private Sub btn_OK_Click()
    Dim lNum As Long, a As Long, DSettouse  As Long, bOutput As Boolean, CCM_ActivationMode
    Dim assPDesc As String, strSQL As String, strMsg As String, assPack As String
    Dim Compprod As Cprod, InfoBox
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    
    If CCM_Comp2Find = "" Then
        MsgBox "Please Select a Match Type." & Chr(10) & Chr(10) & "e.g. Coles Value or Woolworths Private Label", vbOKOnly, "Matching Error"
        Exit Sub
    Else
        CCM_CompData = CCM_Mapping.MatchType(CCM_Comp2Find)
    End If
    
    If CCM_CompData.Competitor = "" Then
        MsgBox "Please select a match type." & Chr(10) & Chr(10) & "e.g. Coles Value or Woolworths Private Label", vbOKOnly, "Matching Error"
        Exit Sub
    End If
    
    If AldiProd.PCode = 0 Then
        MsgBox "Please select an Aldi Product from the box on the left.", vbOKOnly, "Matching Error"
        Exit Sub
    End If
    
    For lNum = 0 To Me.box_Options.ListCount - 1
        If Me.box_Options.Selected(lNum) = True Then
            assPDesc = Me.box_Options.List(lNum, 0)
            assPack = Me.box_Options.List(lNum, 1)
            Exit For
        End If
    Next
    If assPDesc = "" Then
        MsgBox "Please select an competitor product to match to", vbOKOnly, "Matching Error"
        Exit Sub
    End If
    
    
    Select Case CCM_CompData.Competitor
    
        Case "WW"
            DSettouse = 1
            If CCM_UserDefinedState = False Then
                For a = LBound(CCM_WWSKU) To UBound(CCM_WWSKU)
                    If CCM_WWSKU(a).CBA_COM_SKU_CompProdName = assPDesc And CCM_WWSKU(a).CBA_COM_SKU_CompPacksize = assPack Then
                        Compprod.PCode = CCM_WWSKU(a).CBA_COM_SKU_compcode
                        Compprod.Pack = CCM_WWSKU(a).CBA_COM_SKU_CompPacksize
                        Exit For
                    End If
                Next
            Else
                For a = LBound(CCM_UDWWSKU) To UBound(CCM_UDWWSKU)
                    If CCM_UDWWSKU(a).CBA_COM_SKU_CompProdName = assPDesc And CCM_UDWWSKU(a).CBA_COM_SKU_CompPacksize = assPack Then
                        Compprod.PCode = CCM_UDWWSKU(a).CBA_COM_SKU_compcode
                        Compprod.Pack = CCM_UDWWSKU(a).CBA_COM_SKU_CompPacksize
                        Exit For
                    End If
                Next
            End If
        Case "C"
            DSettouse = 2
            If CCM_UserDefinedState = False Then
                For a = LBound(CCM_ColesSKU) To UBound(CCM_ColesSKU)
                    If CCM_ColesSKU(a).CBA_COM_SKU_CompProdName = assPDesc And CCM_ColesSKU(a).CBA_COM_SKU_CompPacksize = assPack Then
                        Compprod.PCode = CCM_ColesSKU(a).CBA_COM_SKU_compcode
                        Compprod.Pack = CCM_ColesSKU(a).CBA_COM_SKU_CompPacksize
                        Exit For
                    End If
                Next
            Else
                For a = LBound(CCM_UDColesSKU) To UBound(CCM_UDColesSKU)
                    If CCM_UDColesSKU(a).CBA_COM_SKU_CompProdName = assPDesc And CCM_UDColesSKU(a).CBA_COM_SKU_CompPacksize = assPack Then
                        Compprod.PCode = CCM_UDColesSKU(a).CBA_COM_SKU_compcode
                        Compprod.Pack = CCM_UDColesSKU(a).CBA_COM_SKU_CompPacksize
                        Exit For
                    End If
                Next
            End If
        Case "DM"
            DSettouse = 3
            If CCM_UserDefinedState = False Then
                For a = LBound(CCM_DMSKU) To UBound(CCM_DMSKU)
                    If CCM_DMSKU(a).CBA_COM_SKU_CompProdName = assPDesc And CCM_DMSKU(a).CBA_COM_SKU_CompPacksize = assPack Then
                        Compprod.PCode = CCM_DMSKU(a).CBA_COM_SKU_compcode
                        Compprod.Pack = CCM_DMSKU(a).CBA_COM_SKU_CompPacksize
                        Exit For
                    End If
                Next
            Else
                For a = LBound(CCM_UDDMSKU) To UBound(CCM_UDDMSKU)
                    If CCM_UDDMSKU(a).CBA_COM_SKU_CompProdName = assPDesc And CCM_UDDMSKU(a).CBA_COM_SKU_CompPacksize = assPack Then
                        Compprod.PCode = CCM_UDDMSKU(a).CBA_COM_SKU_compcode
                        Compprod.Pack = CCM_UDDMSKU(a).CBA_COM_SKU_CompPacksize
                        Exit For
                    End If
                Next
            End If
        Case "FC"
            DSettouse = 4
            If CCM_UserDefinedState = False Then
                For a = LBound(CCM_FCSKU) To UBound(CCM_FCSKU)
                    If CCM_FCSKU(a).CBA_COM_SKU_CompProdName = assPDesc And CCM_FCSKU(a).CBA_COM_SKU_CompPacksize = assPack Then
                        Compprod.PCode = CCM_FCSKU(a).CBA_COM_SKU_compcode
                        Compprod.Pack = CCM_FCSKU(a).CBA_COM_SKU_CompPacksize
                        Exit For
                    End If
                Next
            Else
                For a = LBound(CCM_UDFCSKU) To UBound(CCM_UDFCSKU)
                    If CCM_UDFCSKU(a).CBA_COM_SKU_CompProdName = assPDesc And CCM_UDFCSKU(a).CBA_COM_SKU_CompPacksize = assPack Then
                        Compprod.PCode = CCM_UDFCSKU(a).CBA_COM_SKU_compcode
                        Compprod.Pack = CCM_UDFCSKU(a).CBA_COM_SKU_CompPacksize
                        Exit For
                    End If
                Next
            End If
    End Select

    If Compprod.PCode <> "" Then

        If AldiProd.pCG = 0 And AldiProd.PCode <> 0 Then
            bOutput = CCM_SQLQueries.CBA_COM_MATCHGenPullSQL("isCG", , , , , , CStr(AldiProd.PCode))
            If bOutput = False Then
                GoTo notmatched
            Else
                If CBA_CBISarr(0, 0) = 0 Then
                    GoTo notmatched
                Else
                    AldiProd.pCG = CLng(CBA_CBISarr(0, 0))
                    AldiProd.PSCG = CLng(CBA_CBISarr(1, 0))
                End If
            End If
        End If

        strSQL = "DECLARE @needtoinsert bit" & Chr(10)
        strSQL = strSQL & "DECLARE @Aprod nvarchar(9) = " & AldiProd.PCode & Chr(10)
        strSQL = strSQL & "DECLARE @Cprod nvarchar(50) = '" & Compprod.PCode & "'" & Chr(10)
        strSQL = strSQL & "SET @needtoinsert = isnull((select case when A_Code is null then 0 else 1 end  from  tools.dbo.Com_ProdMap where A_Code = @Aprod),0)" & Chr(10)
        strSQL = strSQL & "if @needtoinsert = 0" & Chr(10)
'        If Aldiprod.PCG > 0 And Aldiprod.PCG < 5 Then
'        strSQL = strSQL & "    INSERT INTO tools.dbo.Com_ProdMap (A_Code, " & CCM_CompData.DBFieldName & ", " & CCM_CompData.AlcoholPackDBField & ")" & Chr(10)
'        strSQL = strSQL & "    VALUES (@Aprod,@Cprod,'" & Compprod.Pack & "')" & Chr(10)
'        Else
        strSQL = strSQL & "    INSERT INTO tools.dbo.Com_ProdMap (A_Code, " & CCM_CompData.DbFieldName & ")" & Chr(10)
        strSQL = strSQL & "    VALUES (@Aprod,@Cprod)" & Chr(10)
'        End If
        strSQL = strSQL & "Else" & Chr(10)
        strSQL = strSQL & "    Update tools.dbo.Com_ProdMap" & Chr(10)
        strSQL = strSQL & "    SET " & CCM_CompData.DbFieldName & " = @Cprod" & Chr(10)
        'If Aldiprod.PCG > 0 And Aldiprod.PCG < 5 Then strSQL = strSQL & ", " & CCM_CompData.AlcoholPackDBField & " = '" & Compprod.Pack & "'"
        strSQL = strSQL & "    where A_Code = @Aprod" & Chr(10)
        strSQL = strSQL & "insert into tools.dbo.Com_MapChange (AldiUser, DateChanged, AldiProd, CompPCode,  CompType)" & Chr(10)
        strSQL = strSQL & "Values('" & Application.UserName & "', getdate(), '" & AldiProd.PCode & "', '" & Compprod.PCode & "'" & ", '" & CCM_CompData.Comp2Find & "')" & Chr(10)
'      '  Debug.Print strSQL
        bOutput = CBA_DB_Connect.CBA_DB_CC_NonC("UPDATE", "COMRADE", CBA_BasicFunctions.TranslateServerName("599DBL12", Date), "SQLOLEDB.1", strSQL, 60, "Tools")
    End If
    strMsg = CCM_CompData.ApplyMatchComment
    Set InfoBox = CreateObject("WScript.Shell")
    InfoBox.Popup strMsg, 1, "Matched Product"
    CCM_Runtime.CCM_updateMatches
    CCM_Matches = getMatches
    If CCM_ActivationMode = False Then updateMatchCollections
Exit Sub

notmatched:
        strMsg = "No Match has been added as there is no CG assigned to " & AldiProd.PCode & " in CBIS"
        Set InfoBox = CreateObject("WScript.Shell")
        InfoBox.Popup strMsg, 1, "Matched Product"
Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-btn_OK_Click", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
'    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Sub
Private Function getMatches() As MatchMap
Dim tmm() As MatchMap, a As Long

    If AldiProd.PCode > 0 Then
        tmm = CCM_Runtime.CCM_getMatches
        For a = LBound(tmm) To UBound(tmm)
            If CStr(tmm(a).AldiPCode) = CStr(AldiProd.PCode) Then
                getMatches = tmm(a)
                Exit For
            End If
        Next
    Else
        'No Aldi product selected
    End If

End Function
Function updateMatchCollections(Optional ByVal Acode As Long)
    Dim thisCD As CompDics, thisCC As CompCols, T
    On Error GoTo Err_Routine
    CBA_ErrTag = ""

    thisCD = getMatchesdic(Acode)
    Set thisCC.WW = New Collection
    Set thisCC.Coles = New Collection
    Set thisCC.DM = New Collection
    Set thisCC.FC = New Collection
    For Each T In thisCD.WW.Keys
        'Debug.Print t
        thisCC.WW.Add T
    Next
    If thisCC.WW.Count > 0 Then CBA_BasicFunctions.CBA_sortCollection thisCC.WW
    For Each T In thisCD.Coles.Keys
        thisCC.Coles.Add T
    Next
    If thisCC.Coles.Count > 0 Then CBA_BasicFunctions.CBA_sortCollection thisCC.Coles
    For Each T In thisCD.DM.Keys
        thisCC.DM.Add T
    Next
    If thisCC.DM.Count > 0 Then CBA_BasicFunctions.CBA_sortCollection thisCC.DM
    For Each T In thisCD.FC.Keys
        thisCC.FC.Add T
    Next
    If thisCC.FC.Count > 0 Then CBA_BasicFunctions.CBA_sortCollection thisCC.FC
    
    Set CCM_Filter.Matches.WW = thisCC.WW
    Set CCM_Filter.Matches.Coles = thisCC.Coles
    Set CCM_Filter.Matches.DM = thisCC.DM
    Set CCM_Filter.Matches.FC = thisCC.FC
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-updateMatchCollections", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    'If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Function
Private Function getMatchesdic(Optional ByVal Acode As Long) As CompDics
    Dim tmm() As MatchMap
    Dim thisAP As Long, WWcnt As Long, Ccnt As Long, DMcnt As Long, FCcnt As Long
    Dim tc As CompDics, aP, a As Long
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    Set tc.WW = New Scripting.Dictionary
    Set tc.Coles = New Scripting.Dictionary
    Set tc.DM = New Scripting.Dictionary
    Set tc.FC = New Scripting.Dictionary
    
    If Me.Box_ProdList.ListCount > 0 Then
    
        tmm = CCM_Runtime.CCM_getMatches
            If UBound(tmm, 1) < 20 Then CCM_updateMatches
        For Each aP In Me.Box_ProdList.List
            If IsNull(aP) = False Then
                thisAP = Mid(aP, 1, InStr(1, aP, "-") - 1)
                If Acode = 0 Or thisAP = Acode Then
                    For a = LBound(tmm) To UBound(tmm)
                        If CLng(tmm(a).AldiPCode) = CStr(thisAP) Then
                            If tmm(a).ColesCB1 <> "" Then If tc.Coles.Exists(tmm(a).ColesCB1) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesCB1, Ccnt
                            If tmm(a).ColesWNAT1 <> "" Then If tc.Coles.Exists(tmm(a).ColesWNAT1) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesWNAT1, Ccnt
                            If tmm(a).ColesWNAT2 <> "" Then If tc.Coles.Exists(tmm(a).ColesWNAT2) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesWNAT2, Ccnt
                            If tmm(a).ColesWNAT3 <> "" Then If tc.Coles.Exists(tmm(a).ColesWNAT3) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesWNAT3, Ccnt
                            If tmm(a).ColesWNAT4 <> "" Then If tc.Coles.Exists(tmm(a).ColesWNAT4) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesWNAT4, Ccnt
                            If tmm(a).ColesWNSW <> "" Then If tc.Coles.Exists(tmm(a).ColesWNSW) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesWNSW, Ccnt
                            If tmm(a).ColesWVIC <> "" Then If tc.Coles.Exists(tmm(a).ColesWVIC) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesWVIC, Ccnt
                            If tmm(a).ColesWQLD <> "" Then If tc.Coles.Exists(tmm(a).ColesWQLD) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesWQLD, Ccnt
                            If tmm(a).ColesWSA <> "" Then If tc.Coles.Exists(tmm(a).ColesWSA) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesWSA, Ccnt
                            If tmm(a).ColesWWA <> "" Then If tc.Coles.Exists(tmm(a).ColesWWA) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesWWA, Ccnt
                            If tmm(a).WWWeb <> "" Then If tc.WW.Exists(tmm(a).WWWeb) = False Then WWcnt = WWcnt + 1: tc.WW.Add tmm(a).WWWeb, WWcnt
                            If tmm(a).WWWNAT1 <> "" Then If tc.WW.Exists(tmm(a).WWWNAT1) = False Then WWcnt = WWcnt + 1: tc.WW.Add tmm(a).WWWNAT1, WWcnt
                            If tmm(a).WWWNAT2 <> "" Then If tc.WW.Exists(tmm(a).WWWNAT2) = False Then WWcnt = WWcnt + 1: tc.WW.Add tmm(a).WWWNAT2, WWcnt
                            If tmm(a).WWWNAT3 <> "" Then If tc.WW.Exists(tmm(a).WWWNAT3) = False Then WWcnt = WWcnt + 1: tc.WW.Add tmm(a).WWWNAT3, WWcnt
                            If tmm(a).WWWNAT4 <> "" Then If tc.WW.Exists(tmm(a).WWWNAT4) = False Then WWcnt = WWcnt + 1: tc.WW.Add tmm(a).WWWNAT4, WWcnt
                            If tmm(a).WWWNSW <> "" Then If tc.WW.Exists(tmm(a).WWWNSW) = False Then WWcnt = WWcnt + 1: tc.WW.Add tmm(a).WWWNSW, WWcnt
                            If tmm(a).WWWVIC <> "" Then If tc.WW.Exists(tmm(a).WWWVIC) = False Then WWcnt = WWcnt + 1: tc.WW.Add tmm(a).WWWVIC, WWcnt
                            If tmm(a).WWWQLD <> "" Then If tc.WW.Exists(tmm(a).WWWQLD) = False Then WWcnt = WWcnt + 1: tc.WW.Add tmm(a).WWWQLD, WWcnt
                            If tmm(a).WWWSA <> "" Then If tc.WW.Exists(tmm(a).WWWSA) = False Then WWcnt = WWcnt + 1: tc.WW.Add tmm(a).WWWSA, WWcnt
                            If tmm(a).WWWWA <> "" Then If tc.WW.Exists(tmm(a).WWWWA) = False Then WWcnt = WWcnt + 1: tc.WW.Add tmm(a).WWWWA, WWcnt
                            If tmm(a).ColesWeb <> "" Then If tc.Coles.Exists(tmm(a).ColesWeb) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesWeb, Ccnt
                            If tmm(a).ColesSB1 <> "" Then If tc.Coles.Exists(tmm(a).ColesSB1) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesSB1, Ccnt
                            If tmm(a).ColesSB2 <> "" Then If tc.Coles.Exists(tmm(a).ColesSB2) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesSB2, Ccnt
                            If tmm(a).ColesSB3 <> "" Then If tc.Coles.Exists(tmm(a).ColesSB3) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesSB3, Ccnt
                            If tmm(a).ColesPL1 <> "" Then If tc.Coles.Exists(tmm(a).ColesPL1) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesPL1, Ccnt
                            If tmm(a).ColesPL2 <> "" Then If tc.Coles.Exists(tmm(a).ColesPL2) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesPL2, Ccnt
                            If tmm(a).ColesPL3 <> "" Then If tc.Coles.Exists(tmm(a).ColesPL3) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesPL3, Ccnt
                            If tmm(a).ColesPL4 <> "" Then If tc.Coles.Exists(tmm(a).ColesPL4) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesPL4, Ccnt
                            If tmm(a).ColesVal1 <> "" Then If tc.Coles.Exists(tmm(a).ColesVal1) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesVal1, Ccnt
                            If tmm(a).ColesVal2 <> "" Then If tc.Coles.Exists(tmm(a).ColesVal2) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesVal2, Ccnt
                            If tmm(a).ColesVal3 <> "" Then If tc.Coles.Exists(tmm(a).ColesVal3) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesVal3, Ccnt
                            If tmm(a).ColesCB1 <> "" Then If tc.Coles.Exists(tmm(a).ColesCB1) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesCB1, Ccnt
                            If tmm(a).ColesPB1 <> "" Then If tc.Coles.Exists(tmm(a).ColesPB1) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesPB1, Ccnt
                            If tmm(a).ColesPB2 <> "" Then If tc.Coles.Exists(tmm(a).ColesPB2) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesPB2, Ccnt
                            If tmm(a).ColesML1 <> "" Then If tc.Coles.Exists(tmm(a).ColesML1) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesML1, Ccnt
                            If tmm(a).ColesML2 <> "" Then If tc.Coles.Exists(tmm(a).ColesML2) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesML2, Ccnt
                            If tmm(a).ColesML3 <> "" Then If tc.Coles.Exists(tmm(a).ColesML3) = False Then Ccnt = Ccnt + 1: tc.Coles.Add tmm(a).ColesML3, Ccnt
                            If tmm(a).WWWW1 <> "" Then If tc.WW.Exists(tmm(a).WWWW1) = False Then WWcnt = WWcnt + 1: tc.WW.Add tmm(a).WWWW1, WWcnt
                            If tmm(a).WWWW2 <> "" Then If tc.WW.Exists(tmm(a).WWWW2) = False Then WWcnt = WWcnt + 1: tc.WW.Add tmm(a).WWWW2, WWcnt
                            If tmm(a).WWWW3 <> "" Then If tc.WW.Exists(tmm(a).WWWW3) = False Then WWcnt = WWcnt + 1: tc.WW.Add tmm(a).WWWW3, WWcnt
                            If tmm(a).WWWW4 <> "" Then If tc.WW.Exists(tmm(a).WWWW4) = False Then WWcnt = WWcnt + 1: tc.WW.Add tmm(a).WWWW4, WWcnt
                            If tmm(a).WWWW5 <> "" Then If tc.WW.Exists(tmm(a).WWWW5) = False Then WWcnt = WWcnt + 1: tc.WW.Add tmm(a).WWWW5, WWcnt
                            If tmm(a).WWHB1 <> "" Then If tc.WW.Exists(tmm(a).WWHB1) = False Then WWcnt = WWcnt + 1: tc.WW.Add tmm(a).WWHB1, WWcnt
                            If tmm(a).WWHB2 <> "" Then If tc.WW.Exists(tmm(a).WWHB2) = False Then WWcnt = WWcnt + 1: tc.WW.Add tmm(a).WWHB2, WWcnt
                            If tmm(a).WWHB3 <> "" Then If tc.WW.Exists(tmm(a).WWHB3) = False Then WWcnt = WWcnt + 1: tc.WW.Add tmm(a).WWHB3, WWcnt
                            If tmm(a).WWCB1 <> "" Then If tc.WW.Exists(tmm(a).WWCB1) = False Then WWcnt = WWcnt + 1: tc.WW.Add tmm(a).WWCB1, WWcnt
                            If tmm(a).WWPB1 <> "" Then If tc.WW.Exists(tmm(a).WWPB1) = False Then WWcnt = WWcnt + 1: tc.WW.Add tmm(a).WWPB1, WWcnt
                            If tmm(a).WWPB2 <> "" Then If tc.WW.Exists(tmm(a).WWPB2) = False Then WWcnt = WWcnt + 1: tc.WW.Add tmm(a).WWPB2, WWcnt
                            If tmm(a).WWML1 <> "" Then If tc.WW.Exists(tmm(a).WWML1) = False Then WWcnt = WWcnt + 1: tc.WW.Add tmm(a).WWML1, WWcnt
                            If tmm(a).WWML2 <> "" Then If tc.WW.Exists(tmm(a).WWML2) = False Then WWcnt = WWcnt + 1: tc.WW.Add tmm(a).WWML2, WWcnt
                            If tmm(a).WWML3 <> "" Then If tc.WW.Exists(tmm(a).WWML3) = False Then WWcnt = WWcnt + 1: tc.WW.Add tmm(a).WWML3, WWcnt
                            If tmm(a).DM1 <> "" Then If tc.DM.Exists(tmm(a).DM1) = False Then DMcnt = DMcnt + 1: tc.DM.Add tmm(a).DM1, DMcnt
                            If tmm(a).DM1Pack <> "" Then If tc.DM.Exists(tmm(a).DM1Pack) = False Then DMcnt = DMcnt + 1:  tc.DM.Add tmm(a).DM1Pack, DMcnt
                            If tmm(a).DM2 <> "" Then If tc.DM.Exists(tmm(a).DM2) = False Then DMcnt = DMcnt + 1:  tc.DM.Add tmm(a).DM2, DMcnt
                            If tmm(a).DM2Pack <> "" Then If tc.DM.Exists(tmm(a).DM2Pack) = False Then DMcnt = DMcnt + 1:  tc.DM.Add tmm(a).DM2Pack, DMcnt
                            If tmm(a).DMQ <> "" Then If tc.DM.Exists(tmm(a).DMQ) = False Then DMcnt = DMcnt + 1:  tc.DM.Add tmm(a).DMQ, DMcnt
                            If tmm(a).DMQPack <> "" Then If tc.DM.Exists(tmm(a).DMQPack) = False Then DMcnt = DMcnt + 1:  tc.DM.Add tmm(a).DMQPack, DMcnt
                            If tmm(a).FC1 <> "" Then If tc.FC.Exists(tmm(a).FC1) = False Then FCcnt = FCcnt + 1: tc.FC.Add tmm(a).FC1, FCcnt
                            If tmm(a).FC1Pack <> "" Then If tc.FC.Exists(tmm(a).FC1Pack) = False Then FCcnt = FCcnt + 1: tc.FC.Add tmm(a).FC1Pack, FCcnt
                            If tmm(a).FC2 <> "" Then If tc.FC.Exists(tmm(a).FC2) = False Then FCcnt = FCcnt + 1: tc.FC.Add tmm(a).FC2, FCcnt
                            If tmm(a).FC2Pack <> "" Then If tc.FC.Exists(tmm(a).FC2Pack) = False Then FCcnt = FCcnt + 1: tc.FC.Add tmm(a).FC2Pack, FCcnt
                            If tmm(a).FCQ <> "" Then If tc.FC.Exists(tmm(a).FCQ) = False Then FCcnt = FCcnt + 1: tc.FC.Add tmm(a).FCQ, FCcnt
                            If tmm(a).FCQPack <> "" Then If tc.FC.Exists(tmm(a).FCQPack) = False Then FCcnt = FCcnt + 1: tc.FC.Add tmm(a).FCQPack, FCcnt
                            Exit For
                        'ElseIf CLng(tmm(a).AldiPCode) > thisAP Then
                            'Exit For
                        End If
                    Next
                End If
            End If
        Next
     End If
     getMatchesdic = tc
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-getMatchesdic", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    'If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function
Private Sub but_Commentadd_Click()
    Dim a As Long
    Dim seldesc As String
    
    For a = 0 To Me.Box_ProdList.ListCount - 1
        If Box_ProdList.Selected(a) = True Then
            seldesc = Trim(Mid(Box_ProdList.List(a, 0), 1, InStr(1, Box_ProdList.List(a, 0), "-") - 1))
            Exit For
        End If
    Next
    If seldesc = "" Then Exit Sub
    
    
    Me.Hide
    CBA_COM_frm_Commentadd.CommentsFormFormulate seldesc
    CBA_COM_frm_Commentadd.Show vbModal
    If CCM_Activation = False Then Me.Show vbModeless


End Sub
Private Function getComp2Find(ByVal btntosave As String) As String
    Dim prevcompetitor As String, ctrl
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    
    Application.EnableEvents = False
    prevcompetitor = CCM_CompData.Competitor
    For Each ctrl In Me.Controls
        If LCase(ctrl.Name) = LCase(btntosave) Then
            ctrl.Value = True
        ElseIf LCase(Mid(ctrl.Name, 1, 2)) = "ob" Then
            If TypeName(ctrl) = "OptionButton" Then ctrl.Value = False
        End If
    Next ctrl
    CCM_Comp2Find = CCM_Mapping.ob_FindComp(btntosave)
    CCM_CompData = CCM_Mapping.MatchType(CCM_Comp2Find)
    If CCM_Activation = False Then
        If CCM_CompData.Competitor <> prevcompetitor Then
            If CCM_UserDefinedState = False Then
                CCM_DefineActiveDataSets
                If CCM_CompData.Competitor = "WW" Then If CCM_IsActiveDataSet(1) = False Then Call ck_WW_Click: ck_WW.Value = True
                If CCM_CompData.Competitor = "C" Then If CCM_IsActiveDataSet(2) = False Then Call ck_Coles_Click: ck_Coles.Value = True
                If CCM_CompData.Competitor = "DM" Then If CCM_IsActiveDataSet(3) = False Then Call ck_DM_Click: ck_DM.Value = True
                If CCM_CompData.Competitor = "FC" Then If CCM_IsActiveDataSet(4) = False Then Call ck_FC_Click: ck_FC.Value = True
            End If
            CCM_CreateCCMArrays
            CCM_PopulateData
        End If
        
        updateAssignedMatches

    End If
    getComp2Find = CCM_Comp2Find
    Application.EnableEvents = True
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-getComp2Find", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
'    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function
Sub updateAssignedMatches()
    Dim pp As String, ctrl
    Dim ThisComp As Cprod
    CCM_Matches = getMatches
    If CCM_Comp2Find = "" Then
        For Each ctrl In Me.Controls
            If LCase(Mid(ctrl.Name, 1, 2)) = "ob" Then
                If TypeName(ctrl) = "OptionButton" And ctrl.Value = True Then
                    If CCM_Comp2Find = "" Then getComp2Find ctrl.Name
                    Exit For
                End If
            End If
        Next ctrl
    End If
    pp = getSpecificMatch(CCM_Comp2Find)
    If pp <> "" Then
        ThisComp = getCompCodeandPack(pp)
        Me.box_Assdesc = ThisComp.PDescription
        Me.box_AssPack = ThisComp.Pack
    Else
        Me.box_Assdesc = ""
        Me.box_AssPack = ""
    End If
    If CCM_Activation = False Then updateMatchCollections AldiProd.PCode
    
End Sub
Private Function getCompCodeandPack(ByVal CompCode As String) As Cprod
    Dim DSettouse, a As Long
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    
    Select Case CCM_CompData.Competitor
        Case "WW"
            DSettouse = 1
            If CCM_UserDefinedState = False Then
                For a = LBound(CCM_WWSKU) To UBound(CCM_WWSKU)
                    If CCM_WWSKU(a).CBA_COM_SKU_compcode = CompCode Then
                        getCompCodeandPack.PDescription = CCM_WWSKU(a).CBA_COM_SKU_CompProdName
                        getCompCodeandPack.Pack = CCM_WWSKU(a).CBA_COM_SKU_CompPacksize
                        Exit For
                    End If
                Next
            Else
                For a = LBound(CCM_UDWWSKU) To UBound(CCM_UDWWSKU)
                    If CCM_UDWWSKU(a).CBA_COM_SKU_compcode = CompCode Then
                        getCompCodeandPack.PDescription = CCM_UDWWSKU(a).CBA_COM_SKU_CompProdName
                        getCompCodeandPack.Pack = CCM_UDWWSKU(a).CBA_COM_SKU_CompPacksize
                        Exit For
                    End If
                Next
            End If
        Case "C"
            DSettouse = 2
            If CCM_UserDefinedState = False Then
                For a = LBound(CCM_ColesSKU) To UBound(CCM_ColesSKU)
                    If CCM_ColesSKU(a).CBA_COM_SKU_compcode = CompCode Then
                        getCompCodeandPack.PDescription = CCM_ColesSKU(a).CBA_COM_SKU_CompProdName
                        getCompCodeandPack.Pack = CCM_ColesSKU(a).CBA_COM_SKU_CompPacksize
                        Exit For
                    End If
                Next
            Else
                For a = LBound(CCM_UDColesSKU) To UBound(CCM_UDColesSKU)
                    If CCM_UDColesSKU(a).CBA_COM_SKU_compcode = CompCode Then
                        getCompCodeandPack.PDescription = CCM_UDColesSKU(a).CBA_COM_SKU_CompProdName
                        getCompCodeandPack.Pack = CCM_UDColesSKU(a).CBA_COM_SKU_CompPacksize
                        Exit For
                    End If
                Next
            End If
        Case "DM"
            DSettouse = 3
            If CCM_UserDefinedState = False Then
                For a = LBound(CCM_DMSKU) To UBound(CCM_DMSKU)
                    If CCM_DMSKU(a).CBA_COM_SKU_compcode = CompCode Then
                        getCompCodeandPack.PDescription = CCM_DMSKU(a).CBA_COM_SKU_CompProdName
                        getCompCodeandPack.Pack = CCM_DMSKU(a).CBA_COM_SKU_CompPacksize
                        Exit For
                    End If
                Next
            Else
                For a = LBound(CCM_UDDMSKU) To UBound(CCM_UDDMSKU)
                    If CCM_UDDMSKU(a).CBA_COM_SKU_compcode = CompCode Then
                        getCompCodeandPack.PDescription = CCM_UDDMSKU(a).CBA_COM_SKU_CompProdName
                        getCompCodeandPack.Pack = CCM_UDDMSKU(a).CBA_COM_SKU_CompPacksize
                        Exit For
                    End If
                Next
            End If
        Case "FC"
            DSettouse = 4
            If CCM_UserDefinedState = False Then
                For a = LBound(CCM_FCSKU) To UBound(CCM_FCSKU)
                    If CCM_FCSKU(a).CBA_COM_SKU_compcode = CompCode Then
                        getCompCodeandPack.PDescription = CCM_FCSKU(a).CBA_COM_SKU_CompProdName
                        getCompCodeandPack.Pack = CCM_FCSKU(a).CBA_COM_SKU_CompPacksize
                        Exit For
                    End If
                Next
            Else
                For a = LBound(CCM_UDFCSKU) To UBound(CCM_UDFCSKU)
                    If CCM_UDFCSKU(a).CBA_COM_SKU_compcode = CompCode Then
                        getCompCodeandPack.PDescription = CCM_UDFCSKU(a).CBA_COM_SKU_CompProdName
                        getCompCodeandPack.Pack = CCM_UDFCSKU(a).CBA_COM_SKU_CompPacksize
                        Exit For
                    End If
                Next
            End If
    End Select
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-getCompCodeandPack", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
'    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function
Private Sub ob_C_NAT_Produce1_Click()
    CCM_Comp2Find = getComp2Find("ob_C_NAT_Produce1")
    Call but_National_Click
End Sub
Private Sub ob_C_NAT_Produce2_Click()
    CCM_Comp2Find = getComp2Find("ob_C_NAT_Produce2")
End Sub
Private Sub ob_C_NAT_Produce3_Click()
    CCM_Comp2Find = getComp2Find("ob_C_NAT_Produce3")
End Sub
Private Sub ob_C_NAT_Produce4_Click()
    CCM_Comp2Find = getComp2Find("ob_C_NAT_Produce4")
End Sub
Private Sub ob_C_NSW_Produce_Click()
    CCM_Comp2Find = getComp2Find("ob_C_NSW_Produce")
    Call but_NSW_Click
End Sub
Private Sub ob_C_VIC_Produce_Click()
    CCM_Comp2Find = getComp2Find("ob_C_VIC_Produce")
End Sub
Private Sub ob_C_QLD_Produce_Click()
    CCM_Comp2Find = getComp2Find("ob_C_QLD_Produce")
End Sub
Private Sub ob_C_SA_Produce_Click()
    CCM_Comp2Find = getComp2Find("ob_C_SA_Produce")
End Sub
Private Sub ob_C_WA_Produce_Click()
    CCM_Comp2Find = getComp2Find("ob_C_WA_Produce")
End Sub
Private Sub ob_DM1_Click()
    CCM_Comp2Find = getComp2Find("ob_DM1")
End Sub
Private Sub ob_DM2_Click()
    CCM_Comp2Find = getComp2Find("ob_DM2")
End Sub
Private Sub ob_DMQ_Click()
    CCM_Comp2Find = getComp2Find("ob_DMQ")
End Sub
Private Sub ob_FC1_Click()
    CCM_Comp2Find = getComp2Find("ob_FC1")
End Sub
Private Sub ob_FC2_Click()
    CCM_Comp2Find = getComp2Find("ob_FC2")
End Sub
Private Sub ob_FCQ_Click()
    CCM_Comp2Find = getComp2Find("ob_FCQ")
End Sub
Private Sub ob_W_NAT_Produce1_Click()
    CCM_Comp2Find = getComp2Find("ob_W_NAT_Produce1")
End Sub
Private Sub ob_W_NAT_Produce2_Click()
    CCM_Comp2Find = getComp2Find("ob_W_NAT_Produce2")
End Sub
Private Sub ob_W_NAT_Produce3_Click()
    CCM_Comp2Find = getComp2Find("ob_W_NAT_Produce3")
End Sub
Private Sub ob_W_NAT_Produce4_Click()
    CCM_Comp2Find = getComp2Find("ob_W_NAT_Produce4")
End Sub
Private Sub ob_W_NSW_Produce_Click()
CCM_Comp2Find = getComp2Find("ob_W_NSW_Produce")
End Sub
Private Sub ob_W_VIC_Produce_Click()
    CCM_Comp2Find = getComp2Find("ob_W_VIC_Produce")
End Sub
Private Sub ob_W_QLD_Produce_Click()
    CCM_Comp2Find = getComp2Find("ob_W_QLD_Produce")
End Sub
Private Sub ob_W_SA_Produce_Click()
    CCM_Comp2Find = getComp2Find("ob_W_SA_Produce")
End Sub
Private Sub ob_W_WA_Produce_Click()
    CCM_Comp2Find = getComp2Find("ob_W_WA_Produce")
End Sub
Private Sub ob_C_Web_Click()
    CCM_Comp2Find = getComp2Find("ob_C_Web")
End Sub
Private Sub ob_W_Web_Click()
    CCM_Comp2Find = getComp2Find("ob_W_Web")
End Sub
Private Sub ob_C_SB1_Click()
    CCM_Comp2Find = getComp2Find("ob_C_SB1")
End Sub
Private Sub ob_C_SB2_Click()
    CCM_Comp2Find = getComp2Find("ob_C_SB2")
End Sub
Private Sub ob_C_SB3_Click()
    CCM_Comp2Find = getComp2Find("ob_C_SB3")
End Sub
Private Sub ob_C_PL1_Click()
    CCM_Comp2Find = getComp2Find("ob_C_PL1")
End Sub
Private Sub ob_C_PL2_Click()
    CCM_Comp2Find = getComp2Find("ob_C_PL2")
End Sub
Private Sub ob_C_PL3_Click()
    CCM_Comp2Find = getComp2Find("ob_C_PL3")
End Sub
Private Sub ob_C_PL4_Click()
    CCM_Comp2Find = getComp2Find("ob_C_PL4")
End Sub
Private Sub ob_C_CV1_Click()
    CCM_Comp2Find = getComp2Find("ob_C_CV1")
End Sub
Private Sub ob_C_CV2_Click()
    CCM_Comp2Find = getComp2Find("ob_C_CV2")
End Sub
Private Sub ob_C_CV3_Click()
    CCM_Comp2Find = getComp2Find("ob_C_CV3")
End Sub
Private Sub ob_W_PL1_Click()
    CCM_Comp2Find = getComp2Find("ob_W_PL1")
End Sub
Private Sub ob_W_PL2_Click()
    CCM_Comp2Find = getComp2Find("ob_W_PL2")
End Sub
Private Sub ob_W_PL3_Click()
    CCM_Comp2Find = getComp2Find("ob_W_PL3")
End Sub
Private Sub ob_W_PL4_Click()
    CCM_Comp2Find = getComp2Find("ob_W_PL4")
End Sub
Private Sub ob_W_PL5_Click()
    CCM_Comp2Find = getComp2Find("ob_W_PL5")
End Sub
Private Sub ob_W_HB1_Click()
    CCM_Comp2Find = getComp2Find("ob_W_HB1")
End Sub
Private Sub ob_W_HB2_Click()
    CCM_Comp2Find = getComp2Find("ob_W_HB2")
End Sub
Private Sub ob_W_HB3_Click()
    CCM_Comp2Find = getComp2Find("ob_W_HB3")
End Sub
Private Sub ob_C_CB_Click()
    CCM_Comp2Find = getComp2Find("ob_C_CB")
End Sub
Private Sub ob_W_CB_Click()
    CCM_Comp2Find = getComp2Find("ob_W_CB")
End Sub
Private Sub ob_C_PB1_Click()
    CCM_Comp2Find = getComp2Find("ob_C_PB1")
End Sub
Private Sub ob_C_PB2_Click()
    CCM_Comp2Find = getComp2Find("ob_C_PB2")
End Sub
Private Sub ob_W_PB1_Click()
    CCM_Comp2Find = getComp2Find("ob_W_PB1")
End Sub
Private Sub ob_W_PB2_Click()
    CCM_Comp2Find = getComp2Find("ob_W_PB2")
End Sub
Private Sub ob_C_ML1_Click()
    CCM_Comp2Find = getComp2Find("ob_C_ML1")
End Sub
Private Sub ob_C_ML2_Click()
    CCM_Comp2Find = getComp2Find("ob_C_ML2")
End Sub
Private Sub ob_C_ML3_Click()
    CCM_Comp2Find = getComp2Find("ob_C_ML3")
End Sub
Private Sub ob_W_ML1_Click()
    CCM_Comp2Find = getComp2Find("ob_W_ML1")
End Sub
Private Sub ob_W_ML2_Click()
    CCM_Comp2Find = getComp2Find("ob_W_ML2")
End Sub
Private Sub ob_W_ML3_Click()
    CCM_Comp2Find = getComp2Find("ob_W_ML3")
End Sub
Private Sub box_SPrice_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then
    If box_SPrice.Value = "" Then CCM_Filter.Price = 0 Else CCM_Filter.Price = box_SPrice.Value
    CCM_CreateCCMArrays
    CCM_PopulateData
End If
End Sub
Private Sub box_SPack_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then
    CCM_Filter.Packsize = box_SPack.Value
    CCM_CreateCCMArrays
    CCM_PopulateData
End If
End Sub
Private Sub box_SDesc_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then
    CCM_Filter.Description = box_SDesc.Value
    CCM_CreateCCMArrays
    CCM_PopulateData
End If
End Sub

Private Sub lbl_pack_Click()
    Call Sorting(3)
End Sub
Private Sub lbl_desc_Click()
    Call Sorting(2)
End Sub
Private Sub lbl_brand_Click()
    Call Sorting(1)
End Sub
Function getSpecificMatch(ByVal Comp2Find As String) As String
    Dim MM As MatchMap
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    MM = getMatches
    Select Case Comp2Find
'CORE RANGE COLES CODES
        Case "ColesWeb"
            getSpecificMatch = MM.ColesWeb
        Case "ColesSB1"
            getSpecificMatch = MM.ColesSB1
        Case "ColesSB2"
            getSpecificMatch = MM.ColesSB2
        Case "ColesSB3"
            getSpecificMatch = MM.ColesSB3
        Case "ColesVal1"
            getSpecificMatch = MM.ColesVal1
        Case "ColesVal2"
            getSpecificMatch = MM.ColesVal2
        Case "ColesVal3"
            getSpecificMatch = MM.ColesVal3
        Case "ColesCB1"
            getSpecificMatch = MM.ColesCB1
        Case "ColesPL1"
            getSpecificMatch = MM.ColesPL1
        Case "ColesPL2"
            getSpecificMatch = MM.ColesPL2
        Case "ColesPL3"
            getSpecificMatch = MM.ColesPL3
        Case "ColesPL4"
            getSpecificMatch = MM.ColesPL4
        Case "ColesPB1"
            getSpecificMatch = MM.ColesPB1
        Case "ColesPB2"
            getSpecificMatch = MM.ColesPB2
        Case "ColesML1"
            getSpecificMatch = MM.ColesML1
        Case "ColesML2"
            getSpecificMatch = MM.ColesML2
        Case "ColesML3"
            getSpecificMatch = MM.ColesML3
        Case "WWWeb"
            getSpecificMatch = MM.WWWeb
        Case "WWHB1"
            getSpecificMatch = MM.WWHB1
        Case "WWHB2"
            getSpecificMatch = MM.WWHB2
        Case "WWHB3"
            getSpecificMatch = MM.WWHB3
        Case "WWWW1"
            getSpecificMatch = MM.WWWW1
        Case "WWWW2"
            getSpecificMatch = MM.WWWW2
        Case "WWWW3"
            getSpecificMatch = MM.WWWW3
        Case "WWWW4"
            getSpecificMatch = MM.WWWW4
        Case "WWWW5"
            getSpecificMatch = MM.WWWW5
        Case "WWCB1"
            getSpecificMatch = MM.WWCB1
        Case "WWPB1"
            getSpecificMatch = MM.WWPB1
        Case "WWPB2"
            getSpecificMatch = MM.WWPB2
        Case "WWML1"
            getSpecificMatch = MM.WWML1
        Case "WWML2"
            getSpecificMatch = MM.WWML2
        Case "WWML3"
            getSpecificMatch = MM.WWML3
        Case "DM1"
            getSpecificMatch = MM.DM1
        Case "DM2"
            getSpecificMatch = MM.DM2
        Case "DMQ"
            getSpecificMatch = MM.DMQ
        Case "FC1"
            getSpecificMatch = MM.FC1
        Case "FC2"
            getSpecificMatch = MM.FC2
        Case "FCQ"
            getSpecificMatch = MM.FCQ
        Case "ColesWNAT1"
            getSpecificMatch = MM.ColesWNAT1
        Case "ColesWNAT2"
            getSpecificMatch = MM.ColesWNAT2
        Case "ColesWNAT3"
            getSpecificMatch = MM.ColesWNAT3
        Case "ColesWNAT4"
            getSpecificMatch = MM.ColesWNAT4
        Case "ColesWNSW"
            getSpecificMatch = MM.ColesWNSW
        Case "ColesWVIC"
            getSpecificMatch = MM.ColesWVIC
        Case "ColesWQLD"
            getSpecificMatch = MM.ColesWQLD
        Case "ColesWSA"
            getSpecificMatch = MM.ColesWSA
        Case "ColesWWA"
            getSpecificMatch = MM.ColesWWA
        Case "WWWNAT1"
            getSpecificMatch = MM.WWWNAT1
        Case "WWWNAT2"
            getSpecificMatch = MM.WWWNAT2
        Case "WWWNAT3"
            getSpecificMatch = MM.WWWNAT3
        Case "WWWNAT4"
            getSpecificMatch = MM.WWWNAT4
        Case "WWWNSW"
            getSpecificMatch = MM.WWWNSW
        Case "WWWVIC"
            getSpecificMatch = MM.WWWVIC
        Case "WWWQLD"
            getSpecificMatch = MM.WWWQLD
        Case "WWWSA"
            getSpecificMatch = MM.WWWSA
        Case "WWWWA"
            getSpecificMatch = MM.WWWWA
    End Select
Exit_Routine:
    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-getSpecificMatch", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
'    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & CBA_strSQL_TBLNAME
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Function
Private Sub Sorting(ByVal Sortcol As Long)
'Dim inarray() As Variant
'Dim loadform As Long
''---------INDEX----------
''       0 = Productcode
''       1 = Brand
''       2 = Description
''       3 = packsize
''       4 = retail
''       6 = per unit retail
'On Error Resume Next
'isittrue = IsEmpty(Calcarray(sortcol, a))
'
'If Err > 0 Or isittrue = True Then
'    Err.Clear
'    On Error GoTo 0
'    Exit Sub
'Else
'    Err.Clear
'    On Error GoTo 0
'End If
'loadform = UBound(Calcarray, 2)
'If loadform > 600 Then
'    strAldimess = "You have requested a large volume of data to be sorted" & Chr(10) & "Please wait while the information is sorted, this should not take more than a minute"
'    Me.Hide
'    frm_AldiWarn.Show (vbModeless)
'    DoEvents
'End If
'
'For a = 0 To UBound(Calcarray, 2)
'    ReDim Preserve inarray(0 To a)
'    inarray(a) = Calcarray(sortcol, a)
'Next
'
'Call CBA_COM_Functions01.QuickSort(inarray, 0, UBound(inarray, 1))
'
'c = -1
'ReDim temparr(0 To 6, 0 To UBound(Calcarray, 2))
'For a = 0 To UBound(inarray, 1)
'    If a = 0 Then
'        For b = 0 To UBound(Calcarray, 2)
'            If Calcarray(sortcol, b) = inarray(a) Then
'                c = c + 1
'                temparr(0, c) = Calcarray(0, b)
'                temparr(1, c) = Calcarray(1, b)
'                temparr(2, c) = Calcarray(2, b)
'                temparr(3, c) = Calcarray(3, b)
'                temparr(4, c) = Calcarray(4, b)
'                temparr(5, c) = Calcarray(5, b)
'                temparr(6, c) = Calcarray(6, b)
'            End If
'        Next
'    ElseIf Not inarray(a - 1) = inarray(a) Then
'        For b = 0 To UBound(Calcarray, 2)
'            If Calcarray(sortcol, b) = inarray(a) Then
'                c = c + 1
'                temparr(0, c) = Calcarray(0, b)
'                temparr(1, c) = Calcarray(1, b)
'                temparr(2, c) = Calcarray(2, b)
'                temparr(3, c) = Calcarray(3, b)
'                temparr(4, c) = Calcarray(4, b)
'                temparr(5, c) = Calcarray(5, b)
'                temparr(6, c) = Calcarray(6, b)
'            End If
'        Next
'    End If
'                Comparer = inarray(a)
'                For d = 0 To UBound(inarray, 1)
'                    If inarray(d) = Comparer Then inarray(d) = Empty
'                Next
'Next
'
'Calcarray = temparr
'
'        ReDim inarray(0 To c, 0 To 5)
'        For a = 0 To UBound(Calcarray, 2)
'            For b = 1 To 5
'                inarray(a, b - 1) = CStr(Calcarray(b, a))
'            Next
'        Next
'        Me.box_Options.Clear
'        Me.box_Options.List = inarray
'If loadform > 600 Then
'    Unload frm_AldiWarn
'    Me.Show vbModeless
'    DoEvents
'End If
End Sub

Private Sub but_Unassign_Click()
    Dim InfoBox As Object, bOutput As Boolean, a, strSQL As String, ThisMatch, CCM_ActivationMode
    On Error GoTo Err_Routine
    CBA_ErrTag = "SQL"
    
    If AldiProd.PCode = 0 Then
        Set InfoBox = CreateObject("WScript.Shell")
        MsgBox "Please select an Aldi Product from the product list on the left.", vbOKOnly, "Matching Error"
        Exit Sub
    End If
    
    If IsEmpty(CCM_Comp2Find) Then
        Set InfoBox = CreateObject("WScript.Shell")
        a = MsgBox("Please Select a Match Type." & Chr(10) & Chr(10) & "e.g. Coles Value or Woolworths Select", vbOKOnly, "Matching Error")
        Exit Sub
    End If
    strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
    strSQL = strSQL & "update tools.dbo.com_prodmap SET " & CCM_CompData.DbFieldName & " = NULL where A_Code = " & AldiProd.PCode & Chr(10)
    strSQL = strSQL & "insert into tools.dbo.Com_MapChange (AldiUser, DateChanged, AldiProd, CompPCode,  CompType)" & Chr(10)
    strSQL = strSQL & "Values('" & Application.UserName & "', getdate(), '" & AldiProd.PCode & "', 'Unassigned', '" & CCM_Comp2Find & "')" & Chr(10)
    'Debug.Print strSQL
    bOutput = CBA_DB_Connect.CBA_DB_CC_NonC("UPDATE", "COMRADE", CBA_BasicFunctions.TranslateServerName("599DBL12", Date), "SQLOLEDB.1", strSQL, 60, "Tools")
    If bOutput = True Then
        Me.box_Assdesc = ""
        Me.box_AssPack = ""
        Set InfoBox = CreateObject("WScript.Shell")
        ThisMatch = InfoBox.Popup(CCM_CompData.DeactivateMatchComment, 1, "UnMatched Product")
    End If
    If CCM_ActivationMode = False Then
        CCM_Runtime.CCM_updateMatches
        updateMatchCollections
    End If
Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-but_Unassign_Click", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & strSQL
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Gen", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub




















