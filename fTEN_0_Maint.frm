VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fTEN_0_Maint 
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10635
   OleObjectBlob   =   "fTEN_0_Maint.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fTEN_0_Maint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit   ' @CBA_Ten-frm0_Maint 190930
Private pbHasBeenSaved As Boolean, pbSetUp As Boolean

Private Sub cboStatus_Change()
    If pbSetUp Then Exit Sub
    If Me.cboStatus.ListIndex > -1 Then
        Me.cmdSave.Visible = True
    End If
End Sub

Private Sub txtIdeaDesc_Change()
    If pbSetUp Then Exit Sub
    Me.cmdSave.Visible = True
End Sub

Private Sub cboBD_Change()
    If pbSetUp Then Exit Sub
    Call mTEN_Runtime.p_FillListBox_BDBA(Me.cboBD, Me.cboBA, IIf(Me.cboBD.ListIndex > -1, Me.cboBD.Value, ""))
    Me.cmdSave.Visible = True
End Sub

Private Sub cboBD_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    '%%RWCall mw_SetBoxHook(cboBD)
End Sub

Private Sub cboBA_Change()
    If pbSetUp Then Exit Sub
    Me.cmdSave.Visible = True
End Sub

Private Sub cmdSave_Click()
    ' Check out the save
    If Me.cboStatus.ListIndex > -1 Then
        varFldVars.sField1 = Me.cboStatus.Column(0, Me.cboStatus.ListIndex)
    Else
        MsgBox "Status must be selected", vbOKOnly
        Me.cboStatus.SetFocus
        Exit Sub
    End If
    If NZ(Me.txtIdeaDesc, "") = "" Then
        MsgBox "A Description must be entered", vbOKOnly
        Me.txtIdeaDesc.SetFocus
        Exit Sub
    Else
        varFldVars.sField2 = Replace(Me.txtIdeaDesc, "'", "`")
    End If
    If Me.cboBA.ListIndex = -1 Then
        MsgBox "BA must be entered", vbOKOnly
        Me.cboBA.SetFocus
        Exit Sub
    Else
        varFldVars.sField3 = Me.cboBA
    End If
    If Me.cboBD.ListIndex = -1 Then
        MsgBox "BD must be entered", vbOKOnly
        Me.cboBD.SetFocus
        Exit Sub
    Else
        varFldVars.sField4 = Me.cboBD
    End If
    pbHasBeenSaved = True
    CBA_bDataChg = True
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    Dim bPassOK As Boolean
    pbSetUp = True
   ' Pre-format the input parms
    Me.StartUpPosition = 0
    Me.Top = Application.Top + (Application.Height / 2) - (Me.Height / 2)
    Me.Left = Application.Left + (Application.Width / 2) - (Me.Width / 2)
''    Me.StartUpPosition = 0
    Me.cmdSave.Visible = False: pbHasBeenSaved = False: CBA_bDataChg = False
    Me.lblVersion.Caption = CBA_getVersionStatus(g_GetDB("Ten"), CBA_Ten_Ver, "Tender Tool", "Ten")  ' Get the latest version
''    Me.Top = varFldVars.lFrmTop
''    Me.Left = varFldVars.lFrmLeft
    Me.cboStatus.Visible = True
    Me.txtIdeaDesc.Visible = True
    Me.cboBA.Visible = True
    Me.cboBD.Visible = True
    CBA_DBtoQuery = 1: Call g_EraseAry(CBA_ABIarr)
    bPassOK = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", "cbo_Fields", g_GetDB(varFldVars.sDB), CBA_MSAccess, varFldVars.sSQL, 120, , , False)
    If bPassOK = True Then
        Call AST_FillDDBox(Me.cboStatus, varFldVars.lCols)
    End If
    Me.txtIdeaDesc = varFldVars.sField2
    Call mTEN_Runtime.p_FillListBox_BDBA(Me.cboBD, Me.cboBA, varFldVars.sField4, varFldVars.sField3)
    Me.cboStatus = IIf(NZ(varFldVars.sField1, 0) = 0, 1, varFldVars.sField1)
    Me.cboStatus.Enabled = IIf(varFldVars.lID2 = 0, False, True)
    pbSetUp = False
    Exit Sub
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Dim lReturn As Long
    If pbHasBeenSaved = False And Me.cmdSave.Visible = True Then
        lReturn = MsgBox("Exit without saving?", vbYesNo + vbDefaultButton2, "Exit warning")
        If lReturn <> vbYes Then
            Cancel = True
            Exit Sub
        End If
        CBA_strAldiMsg = ""
    ElseIf pbHasBeenSaved = False Then
        CBA_strAldiMsg = ""
    End If
End Sub
