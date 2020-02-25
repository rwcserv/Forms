VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fTEN_1_Ideas 
   ClientHeight    =   10695
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   18630
   OleObjectBlob   =   "fTEN_1_Ideas.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "fTEN_1_Ideas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit     ' @fTEN_1_Ideas 190930

Private psSQL As String, pbPassOk As Boolean, pbSetUp As Boolean, bNoChg As Boolean, pbAdd_IP_ID As Boolean, pbUpdIP As Boolean, pbRfrshIP As Boolean
Private plID_ID As Long, plLV_ID As Long, plAdded_ID_ID As Long, psngWidth As Single, psngHeight As Single, psAct As String, pbCancelExit As Boolean
Private psngLeftPos As Single, psngLeftOrig As Single, plMouseMove As Long
Private Const SMALLWIDTH As Long = 110, SMALLHEIGHT = 100

Private Sub cboBD_Change()
    If pbSetUp Then Exit Sub
    Call mw_RemoveBoxHook
    pbSetUp = True
    pbAdd_IP_ID = False
    plID_ID = 0: plLV_ID = 0
    plAdded_ID_ID = 0
''    Call mTEN_Runtime.p_FillListBox_BDBA(Me.cboBD, Me.cboBA, IIf(Me.cboBD.ListIndex > -1, Me.cboBD.Value, ""), , True)
    pbSetUp = False
    Call p_FillListBox_ID
End Sub
Private Sub cboBD_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If pbSetUp Then Exit Sub
    If Me.cboBD.ListCount > 8 Then
        Call mw_SetBoxHook(cboBD) '%%RW
    End If
End Sub

Private Sub cboBA_Change()
    If pbSetUp Then Exit Sub
    Call mw_RemoveBoxHook
    pbAdd_IP_ID = False
    plID_ID = 0: plLV_ID = 0
    plAdded_ID_ID = 0
    Call p_FillListBox_ID
End Sub
Private Sub cboBA_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If pbSetUp Then Exit Sub
    If Me.cboBA.ListCount > 8 Then
        Call mw_SetBoxHook(cboBA)  '%%RW
    End If
End Sub

Private Sub chkActive_Click()
    If pbSetUp Then Exit Sub
    Call mw_RemoveBoxHook
    pbAdd_IP_ID = False
    plAdded_ID_ID = 0
    Call p_FillListBox_ID
End Sub
Private Sub chkActive2_Click()
    If pbSetUp Then Exit Sub
    Call mw_RemoveBoxHook
    Call p_FillListBox_LV(False)
    pbRfrshIP = False
End Sub

Private Sub cmdAdd_ID_Click()
    Dim sBA As String, sBD As String
    Call mw_RemoveBoxHook
    If p_TestWSDelete = False Then Exit Sub

    varFldVars.sHdg = "Idea Maint"
    varFldVars.sSQL = "SELECT Sts_ID,Sts_Desc FROM C0_Statuses WHERE Sts_IdeaValid='Y' ORDER BY Sts_Seq"
    varFldVars.sDB = "TEN"
    varFldVars.sField2 = ""       ' Desc
    varFldVars.sField3 = ""       ' BA
    varFldVars.sField4 = ""       ' BD
    sBA = varFldVars.sField3
    sBD = varFldVars.sField4
    varFldVars.sField1 = Null     ' Sts
    varFldVars.lID2 = 0                                 ' Add Sts Allowed
    plID_ID = 0: plLV_ID = 0
    varFldVars.lCols = 2
    fTEN_0_Maint.Show vbModal
    If varFldVars.sField1 > "" Then
        Me.cmdMapping.Enabled = False
        pbAdd_IP_ID = True
        ' Apply changes
        Call p_Upd_ID(Me.cboBA, Me.cboBD)
        plLV_ID = mTEN_Runtime.Add_TD_LV(plID_ID, 0, "", "", 1, "", "", "", "", "", "", "", 0, 0, "", "", "A")
        Call p_FillListBox_LV(True)
        pbAdd_IP_ID = False
        If mTEN_Runtime.SaveIP_ID Then Call p_SetVis("Idea", True)
        Call cmdMapping_Click
    End If
End Sub

Private Sub cmdAdd_LV_Click()
    Dim sPFsNo As String, sPFsDesc As String, lLVID As Long
    Call mw_RemoveBoxHook
    If Me.lstIdeas.ListIndex = -1 Then Exit Sub
    If p_TestWSDelete = False Then Exit Sub

    ' Copy the promo details from the last PF entered
    If Me.lstPF.ListCount > 0 Then
        lLVID = Me.lstPF.Column(0, Me.lstPF.ListCount - 1)
        sPFsNo = mTEN_Runtime.Get_TD_LV(lLVID, "LV_Portfolio_No")
        sPFsDesc = mTEN_Runtime.Get_TD_LV(lLVID, "LV_Portfolio_Desc")
    End If
    ' Add a record
    plLV_ID = 0
    plID_ID = Me.lstIdeas.Column(0)
    Call mTEN_Runtime.Add_TD_LV(plID_ID, plLV_ID, sPFsNo, sPFsDesc, 1, "", "", "", "", "", "", "", 0, 0, "", "", "A")
    ' Fill in the rest of the values
    Call cmdMapping_Click
End Sub

Private Sub cmdImport_Click()
    Dim sPath As String, sFile As String, sPathFile As String, lColsAdded As Long
    Static bRecursed As Boolean
    Call mw_RemoveBoxHook
    Me.lstIdeas.Enabled = False
    Me.lstPF.Enabled = False
    If Me.lstPF.ListIndex = -1 Then
        MsgBox "Please select a Portfolio Line before importing a comparison Document."
        GoTo Exit_Routine
    ElseIf Val(mTEN_Runtime.Get_TD_LV(plLV_ID, "LV_AH_ID")) > 0 Then
        MsgBox "Tender Document already has an attached ATP import. "
        GoTo Exit_Routine
    ElseIf NZ(Me.lstPF.Column(3), "") = "" Then
        MsgBox "Please Map a Portfolio Version before importing a comparison Document."
        GoTo Exit_Routine
    End If
    If pbSetUp Then Exit Sub
    If bRecursed Then Exit Sub
    ' Delete the old worksheet, if it exists
    If p_TestWSDelete = False Then Exit Sub
    bRecursed = True
    sPathFile = g_GetFile(sPath, sFile)
    If sPathFile = "" Then GoTo Exit_Routine
    CBA_BasicFunctions.CBA_Running "Loading Comparison Data "
    If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.RunningSheetAddComment 6, 4, "File = " & sFile
    lColsAdded = mTEN_Runtime.Get_XLS_File(sPath, sFile)
    If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running

    If lColsAdded > -1 Then
        If lColsAdded > 0 Then
''            Call TEN_Set_Values(, , , , , , , plLV_ID)      ' Save the LV_ID under which the ATP Doc was created
            Call p_FillListBox_LV
            MsgBox sFile & " comparison document has loaded with " & IIf(lColsAdded > CBA_MAX_SUPPS, "the first " & CBA_MAX_SUPPS, lColsAdded) & " suppliers", vbOKOnly
        Else
            MsgBox "No valid supplier columns were found in the " & sFile & " comparison document", vbOKOnly
        End If
    End If
Exit_Routine:
    Me.lstIdeas.Enabled = True
    Me.lstPF.Enabled = True
    bRecursed = False
End Sub

Private Sub cmdMapping_Click()
    Dim lSts As Long, bChanged As Boolean ',sPortf As String,
    Call mw_RemoveBoxHook
    If pbRfrshIP Then Call chkActive2_Click
    If p_TestWSDelete = False Then Exit Sub
    ' Lock the screen
    Call p_SetVis("Lock", True)
    varFldVars.sField4 = CStr(plLV_ID): bChanged = False
    ' Enter the Portfolio
    varFldVars.sType = "PF_ID"
    If mTEN_Runtime.Get_TD_LV(plLV_ID, "LV_Portfolio_Desc") = "" Then
        varFldVars.sField1 = ""
        varFldVars.sField2 = ""
        fTEN_0_Selector.Show vbModal
        If varFldVars.sField1 > "" Then
            Call mTEN_Runtime.Upd_TD_LV(plLV_ID, "LV_Portfolio_No", varFldVars.sField1, "U")
            Call mTEN_Runtime.Upd_TD_LV(plLV_ID, "LV_Portfolio_Desc", varFldVars.sField2, "U")
            bChanged = True
''            Call p_FillListBox_LV(True)
''            ' Assume a change so display the save buttons
''            Call p_SetVis("Idea", True)
        Else
            GoTo Exit_Routine
        End If
    End If
    ' Enter the Version No
    varFldVars.sType = "PV_ID"
    If mTEN_Runtime.Get_TD_LV(plLV_ID, "LV_Version_Desc") = "" Then
        varFldVars.sField1 = ""
        varFldVars.sField2 = ""
        fTEN_0_Selector.Show vbModal
        If varFldVars.sField1 > "" Then
            Call mTEN_Runtime.Upd_TD_LV(plLV_ID, "LV_Version_No", varFldVars.sField1, "U")
            Call mTEN_Runtime.Upd_TD_LV(plLV_ID, "LV_Version_Desc", varFldVars.sField2, "U")
            bChanged = True
''            Call p_FillListBox_LV(True)
''            ' Assume a change so display the save buttons
''            Call p_SetVis("Idea", True)
        Else
            GoTo Exit_Routine
        End If
    End If
    ' Enter the Prior Product
    varFldVars.sType = "PPD_ID"
    If mTEN_Runtime.Get_TD_LV(plLV_ID, "LV_Prior_Product_Code") = "" Then
        varFldVars.sField1 = ""
        varFldVars.sField2 = ""
        fTEN_0_Selector.Show vbModal
        If varFldVars.sField1 > "" Then
            Call mTEN_Runtime.Upd_TD_LV(plLV_ID, "LV_Prior_Product_Code", varFldVars.sField1, "U")
            bChanged = True
''            Call p_FillListBox_LV(True)
''            ' Assume a change so display the save buttons
''            Call p_SetVis("Idea", True)
        Else
            GoTo Exit_Routine
        End If
    End If
    ' Enter the Products
    varFldVars.sType = "PD_ID"
    If mTEN_Runtime.Get_TD_LV(plLV_ID, "LV_Product_Code") = "" Then
        varFldVars.sField1 = ""
        varFldVars.sField2 = ""
        fTEN_0_Selector.Show vbModal
        If varFldVars.sField1 > "" Then
            Call mTEN_Runtime.Upd_TD_LV(plLV_ID, "LV_Product_Code", varFldVars.sField1, "U")
            bChanged = True
''            Call p_FillListBox_LV(True)
''            ' Assume a change so display the save buttons
''            Call p_SetVis("Idea", True)
        Else
            GoTo Exit_Routine
        End If
    End If
    ' Enter the Contracts
    varFldVars.sType = "PC_ID"
    If mTEN_Runtime.Get_TD_LV(plLV_ID, "LV_Contract_No") = "" Then
        varFldVars.sField1 = ""
        varFldVars.sField2 = ""
        fTEN_0_Selector.Show vbModal
        If varFldVars.sField1 > "" Then
            Call mTEN_Runtime.Upd_TD_LV(plLV_ID, "LV_Contract_No", varFldVars.sField1, "U")
            bChanged = True
''            Call p_FillListBox_LV(True)
''            ' Assume a change so display the save buttons
''            Call p_SetVis("Idea", True)
        Else
            GoTo Exit_Routine
        End If
    End If
Exit_Routine:
    If bChanged = True Then
        Call p_FillListBox_LV(True)
        ' Assume a change so display the save buttons
        Call p_SetVis("Idea", True)
    Else
        ' Unlock the screen
        Call p_SetVis("Lock", False)
    End If
End Sub

Public Sub p_Dummy()
    ' Pull focus....
End Sub

Private Sub cmdSave_Click()
    Dim lID_ID As Long, lLV_ID As Long
    Static bRecursed As Boolean
    Call mw_RemoveBoxHook
    If pbSetUp Then Exit Sub
    If bRecursed Then Exit Sub
    bRecursed = True
    Call mTEN_Runtime.SaveIP_ID("Set", False)
    pbSetUp = True
    Call mTEN_Runtime.Write_db_Class_ID(lID_ID, lLV_ID)
    If lID_ID > 0 Then plID_ID = lID_ID: plAdded_ID_ID = lID_ID
    Call p_FillListBox_ID(True)
    If lLV_ID > 0 Then plLV_ID = lLV_ID
    Call p_FillListBox_LV(True)
    Call p_SetOptButtons ' Set the right cmd keys and status controls
    pbSetUp = False
    Call p_SetVis("Idea", False)
    bRecursed = False
End Sub

Private Sub cmdReset_Click()
    Call mw_RemoveBoxHook
    Call mTEN_Runtime.SaveIP_ID("Set", False)
    Call p_FillListBox_ID(True)
    Call p_SetVis("Idea", False)
End Sub

Private Sub cmdTenderDoc_Click()
    Static bRecursed As Boolean
    Call mw_RemoveBoxHook
    If Me.lstPF.ListIndex = -1 Then
        MsgBox "Please select a Portfolio Line before maintaining a Tender Document"
        Exit Sub
    ElseIf NZ(Me.lstPF.Column(3), "") = "" Then
        MsgBox "Please Map a Portfolio Version before creating a Tender Document"
        Exit Sub
    End If
    If pbSetUp Then Exit Sub
    If bRecursed Then Exit Sub
    bRecursed = True
    ' Delete the old worksheet, if it exists
''    Call TenWSDelete
    Call p_SetVis("Lock", True)
    varFldVars.lFrmLeft = g_PosForm(0, (varFldVars.lFldWidth * 4), 0, "Left")
    varFldVars.lFrmTop = g_PosForm(0, 0, 0, "Top")
    varFldVars.lID1 = plLV_ID
    varFldVars.sField3 = mTEN_Runtime.Get_TD_LV(plLV_ID, "LV_TH_Docs")
    varFldVars.sField2 = Me.lblVersion
    varFldVars.sField4 = ""
    fTEN_2_Templates.Show vbModal
    DoEvents
    If varFldVars.sField4 = "Done" Then
        MsgBox "Tender Doc has been produced"
    Else
        Call p_SetVis("Lock", False)
    End If
    bRecursed = False
    
End Sub

Private Sub lblPF_Click()
    Call mw_RemoveBoxHook
    If pbRfrshIP Then Call chkActive2_Click
End Sub

Private Sub lstIdeas_Click()
    Static bRecursed As Boolean
    If pbSetUp Then Exit Sub
    If bRecursed Then Exit Sub
    bRecursed = True
    plLV_ID = 0
    plID_ID = Me.lstIdeas.Column(0)
    Me.frmeAll.Visible = False
    Me.chkActive2.Value = True
    Call p_FillListBox_LV(True)
    bRecursed = False
End Sub

Private Sub lstIdeas_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim sBD As String, sBA As String
    Static bRecursed As Boolean
    If pbSetUp Then Exit Sub
    If bRecursed Then Exit Sub
    bRecursed = True

    varFldVars.lFrmLeft = g_PosForm(0, (varFldVars.lFldWidth * 4), 0, "Left")
    varFldVars.lFrmTop = g_PosForm(0, 0, 0, "Top")
    varFldVars.sHdg = "Idea Maint"
    varFldVars.sSQL = "SELECT Sts_ID,Sts_Desc FROM C0_Statuses WHERE Sts_IdeaValid='Y' ORDER BY Sts_Seq"
    varFldVars.sDB = "TEN"
    varFldVars.lID1 = Me.lstIdeas.Column(0)              ' ID
    varFldVars.lID2 = 1                                 ' All Stses Allowed
    varFldVars.sField2 = Me.lstIdeas.Column(1)          ' Desc
    varFldVars.sField3 = mTEN_Runtime.Get_TD_ID(varFldVars.lID1, "ID_BA_Emp_No")         ' BA
    sBA = varFldVars.sField3
    varFldVars.sField4 = mTEN_Runtime.Get_TD_ID(varFldVars.lID1, "ID_BD_Emp_No")         ' BD
    sBD = varFldVars.sField4
    varFldVars.sField1 = Me.lstIdeas.Column(5)          ' Sts
    varFldVars.lCols = 2
    varFldVars.bAllowNullOfField = False
    fTEN_0_Maint.Show vbModal
    If CBA_bDataChg = False Then GoTo Exit_Routine
    ' Apply changes to the Ideas listbox
    Call p_Upd_ID(Me.cboBA.Value, Me.cboBD.Value)
    If varFldVars.sField1 = "1" Or Me.chkActive.Value = False Then
        On Error Resume Next
        Me.lstIdeas.Value = Null
        Me.lstIdeas.Value = plID_ID
        Me.lstIdeas.Value = Null
        Me.lstIdeas.Value = plID_ID
    Else
        Me.lstPF.Clear
    End If
Exit_Routine:
    bRecursed = False
    Exit Sub
End Sub

Private Sub lstIdeas_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Me.lstIdeas.ListCount > 12 Then
        Call mw_SetBoxHook(lstIdeas)  '%%RW
    Else
        Call mw_RemoveBoxHook
    End If
End Sub

Private Sub lstPF_Click()
    Dim lSts As Long, lTH_ID As Long
    If Me.lstPF.ListIndex = -1 Then
        plLV_ID = 0
        Exit Sub
    Else
        ' Get the Level ID
        plLV_ID = lstPF.Column(0)
    End If
    ' Set Values
    lTH_ID = mTEN_Runtime.Get_TD_LV(plLV_ID, "LV_TH_ID")
    Call mTEN_Runtime.TEN_Set_Values(-1)
    Call mTEN_Runtime.TEN_Set_Values(, plLV_ID, , lTH_ID, , , , , plID_ID)
    ' Get the status
    lSts = mTEN_Runtime.Get_TD_LV(plLV_ID, "LV_Sts_ID")                             'Val(Ten_FldGetSet("Get", "", "LV_Sts_ID", "ID"))
    ' Depending on the sts...
    If lSts = 1 Then
        Me.optActive.Value = True
    ElseIf lSts = 2 Then
        Me.optSuspended.Value = True
    ElseIf lSts = 3 Then
        Me.optDeleted.Value = True
    Else
        MsgBox "sts not found"
    End If
    Call p_SetOptButtons ' Set the right cmd keys and status controls
        
End Sub

Private Sub lstPF_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Me.lstPF.ListCount > 15 Then
        Call mw_SetBoxHook(lstPF)   '%%RW
    Else
        Call mw_RemoveBoxHook
    End If
End Sub

Private Sub optActive_Click()
    Dim lSts As Long
    If pbSetUp Then Exit Sub
    Call mw_RemoveBoxHook
    If Me.lstPF.ListIndex = -1 Then Exit Sub
    lSts = mTEN_Runtime.Get_TD_LV(plLV_ID, "LV_Sts_ID")
    If lSts <> 1 Then
        lSts = 1
        Call mTEN_Runtime.Upd_TD_LV(plLV_ID, "LV_Sts_ID", lSts, "U")
        Call p_SetOptButtons
        Call mTEN_Runtime.Write_db_Class_ID(plID_ID, plLV_ID)
        Call p_FillListBox_LV(False)
    End If
End Sub

Private Sub optSuspended_Click()
    Dim lSts As Long
    If pbSetUp Then Exit Sub
    Call mw_RemoveBoxHook
    If Me.lstPF.ListIndex = -1 Then Exit Sub
    lSts = mTEN_Runtime.Get_TD_LV(plLV_ID, "LV_Sts_ID")
    If lSts <> 2 Then
        lSts = 2
        Call mTEN_Runtime.Upd_TD_LV(plLV_ID, "LV_Sts_ID", lSts, "U")
        Call p_SetOptButtons
        Call mTEN_Runtime.Write_db_Class_ID(plID_ID, plLV_ID)
        Call p_FillListBox_LV(False)
    End If
End Sub

Private Sub optDeleted_Click()
    Dim lSts As Long
    If pbSetUp Then Exit Sub
    Call mw_RemoveBoxHook
    If Me.lstPF.ListIndex = -1 Then Exit Sub
    lSts = mTEN_Runtime.Get_TD_LV(plLV_ID, "LV_Sts_ID")
    If lSts <> 3 Then
        lSts = 3
        Call mTEN_Runtime.Upd_TD_LV(plLV_ID, "LV_Sts_ID", lSts, "U")
        Call p_SetOptButtons
        Call mTEN_Runtime.Write_db_Class_ID(plID_ID, plLV_ID)
        Call p_FillListBox_LV(False)
    End If
End Sub

Private Sub txtFromDate_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Select the date from
    If pbSetUp Then Exit Sub
    If txtSearch.Locked = True Then Exit Sub
    Call mw_RemoveBoxHook
    If p_TestWSDelete = False Then Exit Sub
    pbAdd_IP_ID = False
    plAdded_ID_ID = 0
    varCal.sDate = g_FixDate(Me.txtFromDate)
    Call CBA_BT_CalendarShow(Me.Top, Me.Left, Me.txtFromDate, CBA_D3DMY, False)
    plID_ID = 0: plLV_ID = 0
    Call p_FillListBox_ID
End Sub

Private Sub txtSearch_AfterUpdate()
    If pbSetUp Then Exit Sub
    Call mw_RemoveBoxHook
    plID_ID = 0: plLV_ID = 0
    Call p_FillListBox_ID
End Sub

Private Sub txtSearch_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If p_TestWSDelete = False Then
        Cancel = True
    End If
End Sub

Private Sub p_FillListBox_ID(Optional bReset As Boolean = False)
    ' Fill the Ideas List Box
    Dim bChanged As Boolean, llID_BA_Emp_No As Long, llID_BD_Emp_No As Long
    Static stclID_BA_Emp_No As Long, stclID_BD_Emp_No As Long, stcActive As Boolean, stcSearch As String, stcFromDate As String
    Me.MousePointer = fmMousePointerHourGlass
    bChanged = False
    ' If a BD has been entered or omitted
    If Me.cboBD.ListIndex > -1 Then llID_BD_Emp_No = cboBD
    If stclID_BD_Emp_No <> llID_BD_Emp_No Then bChanged = True
    ' If a BA has been entered or omitted
    If Me.cboBA.ListIndex > -1 Then llID_BA_Emp_No = cboBA
    If bReset Then bChanged = True
    If stclID_BA_Emp_No <> llID_BA_Emp_No Then bChanged = True
    ' If a ststus has been entered
    If Not stcActive = Me.chkActive.Value Then bChanged = True
    If Not stcFromDate = Me.txtFromDate.Value Then bChanged = True
    ' If a search has been entered
    If stcSearch <> NZ(Me.txtSearch, "") Then bChanged = True
    ' If a change then reload the data...
    If bChanged = True Then
        stclID_BA_Emp_No = llID_BA_Emp_No
        stclID_BD_Emp_No = llID_BD_Emp_No
        stcActive = Me.chkActive.Value
        stcSearch = NZ(Me.txtSearch, "")
        stcFromDate = Me.txtFromDate
        Call mTEN_Runtime.Get_db_TD_ID(stclID_BD_Emp_No, stclID_BA_Emp_No, stcSearch, stcFromDate, CLng(IIf(stcActive = True, 1, 0)), plAdded_ID_ID)
    End If
    Call mTEN_Runtime.p_FillListBox_ID(Me.lstIdeas, 9, Me.chkActive.Value, plID_ID)
    Me.lstPF.Clear
    bNoChg = False
    Me.MousePointer = fmMousePointerDefault

End Sub

Public Sub p_FillListBox_LV(Optional bHideCtls As Boolean = False)
    ' Fill the Ideas List Box
    Dim sWhere As String, sAnd As String, lIdx As Long
    Me.MousePointer = fmMousePointerHourGlass
    
    Call mTEN_Runtime.p_FillListBox_LV(Me.lstPF, 16, Me.chkActive2.Value, plID_ID, plLV_ID)
    If Me.lstPF.ListIndex = -1 Then bHideCtls = True
    bNoChg = False
    If bHideCtls Then Me.frmeAll.Visible = False
    On Error Resume Next
    Me.MousePointer = fmMousePointerDefault

End Sub

Private Sub p_FillBDDDBox()
    ' Fill the BD / BA DDB Boxes
    Call mTEN_Runtime.p_FillListBox_BDBA(Me.cboBD, Me.cboBA, IIf(Me.cboBD.ListIndex > -1, Me.cboBD.Value, ""), , True)
End Sub

Private Sub p_SetVis(sType As String, bSts As Boolean)
    If sType = "Idea" Then
        Me.cmdSave.Visible = bSts
        Me.cmdReset.Visible = bSts
    End If
    If sType = "Idea" Or sType = "Lock" Then
        Me.chkActive.Locked = bSts
        Me.chkActive2.Locked = bSts
        Me.cboBA.Locked = bSts
        Me.cboBD.Locked = bSts
        Me.txtSearch.Locked = bSts
        Me.txtFromDate.Locked = bSts
        Me.lstIdeas.Locked = bSts
        Me.lstPF.Locked = bSts
    End If
End Sub

Private Sub p_SetOptButtons()
    ' Set up the option buttons
    If Me.frmeAll.Visible = False Then Me.frmeAll.Visible = True
    If Me.frmeAll.Visible = True Then
        If Me.optActive.Value = True Then
            Me.cmdAdd_LV.Enabled = True
            Me.optActive.Caption = "Active"
            If mTEN_Runtime.Get_TD_LV(plLV_ID, "LV_Contract_No") > "" Then
                Me.cmdMapping.Enabled = False
            Else
                Me.cmdMapping.Enabled = True
            End If
            If mTEN_Runtime.SaveIP_ID Or mTEN_Runtime.SaveIP_LV Then
                Me.cmdImport.Enabled = False
                Me.cmdTenderDoc.Enabled = False
            Else
                Me.cmdImport.Enabled = True
                Me.cmdTenderDoc.Enabled = True
            End If
        Else
            Me.optActive.Caption = "Set Active"
            Me.cmdMapping.Enabled = False
            Me.cmdImport.Enabled = False
            Me.cmdTenderDoc.Enabled = False
        End If
        If Me.optSuspended.Value = True Then Me.optSuspended.Caption = "Suspended" Else Me.optSuspended.Caption = "Set Suspended"
        If Me.optDeleted.Value = True Then Me.optDeleted.Caption = "Deleted" Else Me.optDeleted.Caption = "Set Deleted"
    End If
End Sub

Private Function p_Upd_ID(ByVal sBA, ByVal sBD) As Boolean
    ' Will get and re / display the Ideas listbox
    Dim sGrp_Emp_No As String
    Call mTEN_Runtime.LoadIP_ID("Set", True)
    pbSetUp = True
    p_Upd_ID = True
    sBA = NZ(sBA, ""): sBD = NZ(sBD, "")
    If CBA_bDataChg Then
        sGrp_Emp_No = CStr(mTEN_Runtime.Get_BDBA_Value(CLng(varFldVars.sField4), "Grp_Emp_No"))
        If pbAdd_IP_ID Then
            Call mTEN_Runtime.SaveIP_ID("Set", True)
            Call p_SetVis("Idea", True)
            plID_ID = mTEN_Runtime.Add_TD_ID(0, CStr(varFldVars.sField2), 1, "", CLng(varFldVars.sField3), CLng(varFldVars.sField4), CLng(sGrp_Emp_No), "", "", "", "A")
        Else
            Call mTEN_Runtime.Upd_TD_ID(plID_ID, "ID_Desc", varFldVars.sField2, "U")
            Call mTEN_Runtime.Upd_TD_ID(plID_ID, "ID_BA_Emp_No", varFldVars.sField3, "U")
            Call mTEN_Runtime.Upd_TD_ID(plID_ID, "ID_BD_Emp_No", varFldVars.sField4, "U")
            Call mTEN_Runtime.Upd_TD_ID(plID_ID, "ID_GD_Emp_No", sGrp_Emp_No, "U")
            Call mTEN_Runtime.Upd_TD_ID(plID_ID, "ID_Sts_ID", varFldVars.sField1, "U")
            ' Save to the database if a change
            Call mTEN_Runtime.Write_db_Class_ID(plID_ID)
        End If
        Me.frmeAll.Visible = False
        Call p_FillListBox_ID
        On Error Resume Next
        Me.lstIdeas = plID_ID
    End If

    pbSetUp = False
    Call mTEN_Runtime.LoadIP_ID("Set", False)
End Function

Private Function p_ActTime(Optional sUpd As String = "")
    ' Will ensure that the Activate / Deactivate routines are not called too close together and effectively cancel each other out
    Static lWasTime As Long, lTimes As Long
    Dim lNowTime As Long
    lTimes = lTimes + 1
    lNowTime = CLng(Val(Format(Now(), "hhnnss") & Right(Format(Timer, "#0.00"), 2)))
    If lWasTime < lNowTime - 10 Then
        lWasTime = lNowTime
        p_ActTime = False
    Else
        p_ActTime = True
        If sUpd <> "NoUpd" Then lWasTime = lNowTime
    End If
    If lTimes > 3 Then lTimes = 0:

End Function

Private Function p_TestWSDelete() As Boolean
    ' Test to see if the worksheet should be deleted - In Ideas Form
    p_TestWSDelete = True: pbCancelExit = False
    If mTEN_Runtime.TenTestWS() = True Then
        If MsgBox("Delete without saving? Click 'Yes'" & vbCrLf & vbCrLf & "Continue to edit? Click 'No'", vbYesNo, "WARNING: Tender Document not saved!") = vbYes Then
            ' Delete workbook if was created
            Call mTEN_Runtime.TenWSDelete
            Call p_SetVis("Lock", False)
        Else
            p_TestWSDelete = False: pbCancelExit = True
        End If
    End If
End Function

''Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
''    Debug.Print Button & "," & X & "-" & Y
''    DoEvents
''End Sub
''Private Sub UserForm_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Control As MSForms.Control, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal State As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
''    Debug.Print "drag," & X & "-" & Y
''    DoEvents
''
''End Sub
Private Sub UserForm_Activate()
    ' Debug.Print "(Act)" & psAct;
    Call mw_RemoveBoxHook
    If plMouseMove = g_GetTickCount() Then
        If psngWidth > Me.Width Then
            If mTEN_Runtime.ShowIP_ID("IsShowing", True) = True And mTEN_Runtime.ShowIP_ID("IsLoading") = False Then
                If p_TestWSDelete = False Then Exit Sub
                If p_ActTime = True Then Exit Sub
                Call p_SetVis("Lock", False)
                Me.Width = psngWidth: Me.Height = psngHeight
                Me.Left = psngLeftOrig
            End If
        End If
        If psAct = "" Then psAct = "Act"
    End If
End Sub

Private Sub Image2_Click()
    Call mw_RemoveBoxHook
    If psngWidth > Me.Width Then
        plMouseMove = g_GetTickCount()
        Call UserForm_Activate
''    ElseIf psngWidth = Me.Width Then
''        Call UserForm_Deactivate
    End If
End Sub

Private Sub lblVersion_Click()
    Call mw_RemoveBoxHook
    If psngWidth > Me.Width Then
        plMouseMove = g_GetTickCount()
        Call UserForm_Activate
''    ElseIf psngWidth = Me.Width Then
''        Call UserForm_Deactivate
    End If
End Sub

Private Sub UserForm_Click()
    Call mw_RemoveBoxHook
    If psngWidth > Me.Width Then
        plMouseMove = g_GetTickCount()
        Call UserForm_Activate
''        ElseIf psngWidth = Me.Width Then
''            Call UserForm_Deactivate
    End If
End Sub

Private Sub UserForm_Deactivate()
    If psngWidth = Me.Width Then
        If mTEN_Runtime.ShowIP_ID("IsShowing", False) = False Then
            If p_ActTime = True Then Exit Sub
            Me.Width = SMALLWIDTH: Me.Height = SMALLHEIGHT
            Me.Left = psngLeftPos
        End If
    End If
    psAct = "DeAct"
End Sub

Private Sub UserForm_Initialize()
    pbSetUp = True
    Me.StartUpPosition = 0
    Me.Top = Application.Top + (Application.Height / 2) - (Me.Height / 2)
    Me.Left = Application.Left + (Application.Width / 2) - (Me.Width / 2)
    psngLeftOrig = Me.Left
    psngLeftPos = Me.Left + Me.Width - SMALLWIDTH
    mTEN_Runtime.ShowIP_ID ("SetShow")
    Me.lblVersion.Caption = CBA_getVersionStatus(g_GetDB("Ten"), CBA_Ten_Ver, "Tender Tool", "Ten")  ' Test to see if it is the latest version
    psngWidth = Me.Width: psngHeight = Me.Height
    Me.frmeAll.Visible = False
    plAdded_ID_ID = 0: plID_ID = 0: plLV_ID = 0
    Me.txtFromDate = Format(DateSerial(Year(Date) - 1, Month(Date), Day(Date)), CBA_D3DMY)
    Call p_FillListBox_ID
    Call p_FillBDDDBox
    pbSetUp = False
End Sub
''Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
''    If Button = 0 Then
''        plMouseMove = g_GetTickCount()
''    End If
''''    Debug.Print Button & "," & X & "-" & Y
''''    DoEvents
''End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error Resume Next
    Call mw_RemoveBoxHook
    If CBA_BasicFunctions.isRunningSheetDisplayed = False Then
        Call p_TestWSDelete
    Else
        pbCancelExit = True
    End If
    If pbCancelExit Then
        Cancel = vbCancel
        Exit Sub
    End If
End Sub

''Private Sub UserForm_Scroll(ByVal ActionX As MSForms.fmScrollAction, ByVal ActionY As MSForms.fmScrollAction, ByVal RequestDx As Single, ByVal RequestDy As Single, ByVal ActualDx As MSForms.ReturnSingle, ByVal ActualDy As MSForms.ReturnSingle)
''    Debug.Print "scroll," & ActionX & "-" & ActionY
''    DoEvents
''
''End Sub
