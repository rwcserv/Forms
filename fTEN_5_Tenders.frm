VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fTEN_5_Tenders 
   Caption         =   "Ten Portfolios"
   ClientHeight    =   10575
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11865
   OleObjectBlob   =   "fTEN_5_Tenders.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "fTEN_5_Tenders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit     ' @fTEN_5_Tenders 29/07/2019

Private psSQL As String, pbPassOk As Boolean, pbSetUp As Boolean, bNoChg As Boolean, pbAddIP As Boolean, pbUpdIP As Boolean, pbRfrshIP As Boolean
Private plSH_ID As Long, plSL_ID As Long, plLV_ID As Long, psSH_Desc As String, psNPD_R_MSO As String
Private clsTemplate As cTEN3_Templates
Private clsTender As cTEN6_Tenders

Private Sub cmdTenderDoc_Click()
''    Unload fTEN_5_Tenders
    If InStr(1, psSH_Desc, "NPD") > 0 Then
        psNPD_R_MSO = "NPD"
    ElseIf InStr(1, psSH_Desc, "MSO") > 0 Then
        psNPD_R_MSO = "MSO"
    Else
        psNPD_R_MSO = "R"
    End If
    ' Set Values
    Call mTEN_Runtime.TEN_Set_Values(plSH_ID, plLV_ID, , , psNPD_R_MSO, psSH_Desc)
    ' Close the form here
    Unload fTEN_5_Tenders
    ' Get the status
    Call mTEN_Runtime.CrtTenderDoc(0, "", True, "T")
    varFldVars.sField4 = "Done"
''    call mTEN_Runtime.FilePrint( "Exit frm5_Tenders-cmdTenderDoc_Click", , , True)
End Sub

Private Sub lblToA_Click()
    If Me.lstU.ListIndex = -1 Then
        MsgBox "Please select a Used Segment to transfer"
        Exit Sub
    End If
    If mTEN_Runtime.Get_TD_SL(Me.lstU.Column(0), "SL_Mandatory") = "Y" Then
        MsgBox "We can't remove a mandatory segment"
        Exit Sub
    End If
    Call mTEN_Runtime.Upd_TD_SL(Me.lstU.Column(0), "SL_Mandatory", "A", "U")
    Call mTEN_Runtime.p_FillListBoxes_TD(Me.lstU, Me.lstA, plSH_ID)
End Sub

Private Sub lblToU_Click()
    If Me.lstA.ListIndex = -1 Then
        MsgBox "Please select an Available Segment to transfer"
        Exit Sub
    End If

    Call mTEN_Runtime.Upd_TD_SL(Me.lstA.Column(0), "SL_Mandatory", "U", "U")
    Call mTEN_Runtime.p_FillListBoxes_TD(Me.lstU, Me.lstA, plSH_ID)
End Sub

Private Sub lstA_Click()
    Dim lSts As Long
    If Me.lstA.ListIndex = -1 Then Exit Sub
'    ' Get the Level ID
'    plSL_ID = lstA.Column(0)
'    psNPD_R_MSO = lstA.Column(1)
''    Call p_SetOptButtons ' Set the right cmd keys and status controls
End Sub

Private Sub lstA_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call mw_SetBoxHook(lstA)
End Sub
Private Sub lstU_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call mw_SetBoxHook(lstU)
End Sub

Private Sub lstA_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.lstA.ListIndex = -1 Then Exit Sub
    Call lblToU_Click
End Sub

Private Sub lstU_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.lstU.ListIndex = -1 Then Exit Sub
    Call lblToA_Click
End Sub

Private Sub txtSearch_Change()
    If pbSetUp Then Exit Sub
    Call p_FillTendersListBox
End Sub

Private Sub UserForm_Initialize()
    pbSetUp = True
    Me.StartUpPosition = 0
    Me.Top = Application.Top + (Application.Height / 2) - (Me.Height / 2)
    Me.Left = Application.Left + (Application.Width / 2) - (Me.Width / 2)
    
    Set clsTemplate = New cTEN3_Templates
    Set clsTender = New cTEN6_Tenders

    plLV_ID = varFldVars.lID1
    psSH_Desc = varFldVars.sField4
    plSH_ID = varFldVars.lID2
    Call mTEN_Runtime.p_FillListBoxes_TD(Me.lstU, Me.lstA, plSH_ID)
   
    Me.lblVersion.Caption = varFldVars.sField2
    Me.lblExtra.Caption = psSH_Desc
    Call g_PosForm(Me.Top, Me.Width, Me.Left, , True)
    cmdTenderDoc.Enabled = True
    varFldVars.sField4 = "Inited"
    pbSetUp = False
End Sub

Sub p_FillTendersListBox()
    ' Fill the Tenders List Box
    Dim sWhere As String, sAnd As String, lIdx As Long
    Me.MousePointer = fmMousePointerHourGlass
    sWhere = "": sAnd = ""
    psSQL = "SELECT  SH_ID,SH_Desc,SH_BA,SH_BD,Sts_Desc,SH_Sts_ID,SH_UpdDate,SH_UpdUser,SH_CrtUser FROM qry_AllTenders1 "
    ' If a search has been entered
    If NZ(Me.txtSearch, "") > "" Then
        sWhere = sWhere & sAnd & " SH_Desc LIKE '%" & Replace(Me.txtSearch, "'", "`") & "%'"
        sAnd = " AND "
    End If
    If sWhere > "" Then psSQL = psSQL & " WHERE " & sWhere
    psSQL = psSQL & " GROUP BY SH_ID,SH_Desc,SH_BA,SH_BD,Sts_Desc,SH_Sts_ID,SH_UpdDate,SH_UpdUser,SH_CrtUser ORDER BY SH_ID DESC;"
    CBA_DBtoQuery = 1: Erase CBA_ABIarr
    
    pbPassOk = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", "Ten_Query", g_GetDB("Ten"), CBA_MSAccess, psSQL, 120, , , False)
    Me.lstA.Clear
    If pbPassOk = True Then
        For lIdx = 0 To UBound(CBA_ABIarr, 2)
            CBA_ABIarr(6, lIdx) = g_FixDate(CBA_ABIarr(6, lIdx), CBA_D3DMY)
        Next
    End If
    CBA_DBtoQuery = 3
    bNoChg = False
    Me.MousePointer = fmMousePointerDefault

End Sub

Sub p_FillListBox_LV(Optional bHideCtls As Boolean = False)
    ' Fill the Tenders List Box
    Dim sWhere As String, sAnd As String, lIdx As Long
    Me.MousePointer = fmMousePointerHourGlass
    
    Set clsTender = New cTEN6_Tenders
    Call clsTender.Generate(plSH_ID, "", "")
    bNoChg = False
    If bHideCtls Then Me.frmeAll.Visible = False
    On Error Resume Next
    Me.MousePointer = fmMousePointerDefault

End Sub

Private Function p_UpdTender(sBA As String, sBD As String) As Boolean
    pbSetUp = True
    p_UpdTender = True
    If CBA_bDataChg Then
        Me.frmeAll.Visible = False
''        plSH_ID = mTEN_Runtime.Get_TD_LV(plLV_ID, "SH_ID")
        Call p_FillTendersListBox
        On Error Resume Next
    End If
    pbSetUp = False
End Function

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error Resume Next
    Set clsTender = Nothing
End Sub
