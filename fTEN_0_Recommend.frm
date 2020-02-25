VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fTEN_0_Recommend 
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7530
   OleObjectBlob   =   "fTEN_0_Recommend.frx":0000
End
Attribute VB_Name = "fTEN_0_Recommend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit   ' @fTEN_0_Recommend 191116
Const CREGS As String = "Min,Der,Stp,Pre,Dan,Bre,Rgy,Jkt", pCNO_REGS = 8
Private psdSupp As Scripting.Dictionary
Private psdSuppReg(1 To pCNO_REGS) As Scripting.Dictionary
Private pSet_UpIP As Boolean
Private plTfrMergeHeight As Long, plTfrHdgIdx As Long '',plTfrRowDiff As Long,  ', plTfrGrpRows As Long
Private plAddMergeHeight As Long, plAddGrpRows As Long, plAddHdgIdx As Long '', plAddRowDiff As Long '', plAddLastGrpRow01 As Long
Private plStartRow As Long, plEndRow As Long, plTSDiff As Long, plDDP As Long, plExW As Long, plFOB As Long, plTrial As Long, plFlgIdx As Long

Private Sub cboSupp1_Change()
    Dim lReg As Long, sValue_Ext As String
    If pSet_UpIP Then Exit Sub
    If Me.cboSupp1.ListIndex = -1 Then
        sValue_Ext = "??vvvv???"
        Me.cmdConfirm.Visible = False
    Else
        sValue_Ext = Me.cboSupp1.Column(1)
        Me.cmdConfirm.Visible = True
    End If
    For lReg = 1 To pCNO_REGS
        If psdSuppReg(lReg).Exists(sValue_Ext) Then
            Me("chk" & lReg & "Regions").Value = True
            Me("chk" & lReg & "Regions").Enabled = True
            Me("optFOB" & lReg).Visible = True
            Me("optFG" & lReg).Visible = True
            Me("optDDP" & lReg).Visible = True
        Else
            Me("chk" & lReg & "Regions").Value = False
            Me("chk" & lReg & "Regions").Enabled = False
            Me("optFOB" & lReg).Visible = False
            Me("optFG" & lReg).Visible = False
            Me("optDDP" & lReg).Visible = False
        End If
    Next

End Sub

Private Sub cmdConfirm_Click()
    Dim lIdx As Long, lIdxG As Long, lIdxR As Long, sAddit As String, lReg As Long, lGrp_No As Long, lMult_No As Long
    Dim UHAry() As String, sSupp As String, sSupp_Ext As String
    
    On Error GoTo Err_Routine
    If Me.cboSupp1.ListIndex = -1 Then GoTo Exit_Routine
    CBA_Error = ""
    pSet_UpIP = True
    sSupp = Me.cboSupp1.Column(1)
    
    For lReg = 1 To pCNO_REGS
        If psdSuppReg(lReg).Exists(sSupp) Then
            If Me("chk" & lReg & "Regions").Value = True Then
                lIdxR = 0
                If Me("optDDP" & lReg).Value = True Then
                    lIdxR = plDDP
                ElseIf Me("optFG" & lReg).Value = True Then
                    lIdxR = plExW
                ElseIf Me("optFOB" & lReg).Value = True Then
                    lIdxR = plFOB
                End If
                If lIdxR > 0 Then
                    lIdxG = psdSuppReg(lReg).Item(sSupp)
                    Call mTEN_Runtime.Upd_TD_UT(lIdxG, "UT_Addit" & plFlgIdx, lIdxR)
                End If
            End If
        End If
    Next
    pSet_UpIP = False
''    Call cboSupp1_Change
    CBA_bDataChg = True
Exit_Routine:
    On Error Resume Next
    Unload fTEN_0_Recommend
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-TEN.frm0_Recommend.UserForm_Initialize", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Ten", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Sub

Private Sub UserForm_Initialize()
    Dim lIdx As Long, lIdxG As Long, lIdxR As Long, sAddit As String, lReg As Long, lGrp_No As Long, lMult_No As Long
    Dim UHAry() As String, sSupp As String, sSupp_Ext As String
    
    On Error GoTo Err_Routine
    CBA_Error = ""
    pSet_UpIP = True
    CBA_bDataChg = False
    Me.StartUpPosition = 0
    Me.Top = Application.Top + (Application.Height / 2) - (Me.Height / 2)
    Me.Left = Application.Left + (Application.Width / 2) - (Me.Width / 2)
    
''    varFldVars.lFrmLeft = g_PosForm(0, (varFldVars.lFldWidth * 4), 0, "Left")
''    varFldVars.lFrmTop = g_PosForm(0, 0, 0, "Top")
''    varFldVars.sHdg = "Recommendation Selection"
    Set psdSupp = New Scripting.Dictionary
    For lIdx = 1 To pCNO_REGS
        Set psdSuppReg(lIdx) = New Scripting.Dictionary
    Next
    sAddit = varFldVars.sField1
    UHAry = Split(sAddit, ",")                                   ' The following ---Tfr---- fields come from the the Tfr Button, but reflect the lines in the Tfr region
    plTfrMergeHeight = Val(UHAry(0))                             ' Merge Height of the to be transferred lines
'    plTfrGrpRows = Val(UHAry(1))                                 ' Number of 'to be transferred' lines (i.e. 8)
''    plTfrRowDiff = Val(UHAry(2))                                 ' Number of rows to the first top position
    plAddHdgIdx = Val(UHAry(3))                                  ' Add Button Index
    plTfrHdgIdx = Val(UHAry(4))                                  ' Tfr Button Index
    sAddit = mTEN_Runtime.Get_TD_UT(plAddHdgIdx, "UT_Addit")  ' The following ---Add---- fields come from the the Add Button, but reflect the lines in the Add region
    UHAry = Split(sAddit, ",")                                   ' Merge Height of the transferred to lines
    plAddMergeHeight = Val(UHAry(0))                             ' Number of transferred lines (i.e. 8)
    plAddGrpRows = Val(UHAry(1))                                 ' How Many Rows are in the Add Group
''    plAddRowDiff = Val(UHAry(2))                                 ' Difference between the button header Top_Row and the FIRST line Top_Row
''    plAddLastGrpRow01 = Val(UHAry(3))                            ' Difference between the button header Top_Row and the LAST line Top_Row
    UHAry = Split(varFldVars.sField2, ",")
    plStartRow = UHAry(0)                                        ' Tfr Start Line
    plEndRow = UHAry(1)                                          ' Tfr End Line
    plTSDiff = UHAry(2)
    plDDP = UHAry(3)
    plExW = UHAry(4)
    plFOB = UHAry(5)
''    plTrial = UHAry(6)
    plFlgIdx = UHAry(6)
    
    For lIdxG = plStartRow To plEndRow Step plTSDiff '   (plAddGrpRows * plAddMergeHeight)
        lGrp_No = Val(mTEN_Runtime.Get_TD_UT(lIdxG, "UT_Grp_No"))
        lMult_No = g_Get_Mid_Fmt(lGrp_No, 4, 2)
        lReg = g_Get_Mid_Fmt(lGrp_No, 2, 2)
        sSupp = mTEN_Runtime.Get_TD_UT(lIdxG, "UT_Default_Value")
''        If sSupp Like "Fred*" Then
''            sSupp = sSupp
''        End If
        sSupp_Ext = sSupp & Format(lMult_No, "00")
        If sSupp > "" Then
            If Not psdSupp.Exists(sSupp_Ext) Then
                psdSupp.Add sSupp_Ext, lMult_No
                Me("cboSupp1").AddItem lIdxG
                Me("cboSupp1").List(Me("cboSupp1").ListCount - 1, 1) = sSupp_Ext
                Me("cboSupp1").List(Me("cboSupp1").ListCount - 1, 2) = sSupp
            End If
            psdSuppReg(lReg).Add sSupp_Ext, lIdxG
        End If
    Next
    pSet_UpIP = False
    Call cboSupp1_Change

Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-TEN.frm0_Recommend.UserForm_Initialize", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Ten", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
''    Dim lReturn As Long
''    If pbHasBeenSaved = False And Me.cmdSave.Visible = True Then
''        lReturn = MsgBox("Exit without saving?", vbYesNo + vbDefaultButton2, "Exit warning")
''        If lReturn <> vbYes Then
''            Cancel = True
''            Exit Sub
''        End If
''        CBA_strAldiMsg = ""
''    ElseIf pbHasBeenSaved = False Then
''        CBA_strAldiMsg = ""
''    End If
End Sub

