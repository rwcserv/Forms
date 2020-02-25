VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CBA_frmEntryField 
   Caption         =   "Entry Field"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10770
   OleObjectBlob   =   "CBA_frmEntryField.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "CBA_frmEntryField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit   ' @CBA_ASyst 190401
Private pbHasBeenSaved As Boolean, pbSetUp As Boolean

Private Sub cboField_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call mw_SetBoxHook(cboField)
End Sub

Private Sub cboField_Change()
    If pbSetUp Then Exit Sub
    If Me.cboField.ListIndex > -1 Then
        Me.cmdSave.Visible = True
    End If
End Sub

Private Sub cmdNA_Click()
    If varFldVars.sType = "ComboBox" Then
        Me.cboField = Null
    Else
        Me.txtField = ""
    End If
End Sub

Private Sub cmdSave_Click()
    ' Check out the save
    If varFldVars.sType = "ComboBox" Then
        If Me.cboField.ListIndex > -1 Then
            Select Case varFldVars.lCols
            Case 1
                varFldVars.sField1 = Me.cboField.Column(0, Me.cboField.ListIndex)
            Case 2
                varFldVars.sField1 = Me.cboField.Column(0, Me.cboField.ListIndex)
                varFldVars.sField2 = Me.cboField.Column(1, Me.cboField.ListIndex)
            Case 3
                varFldVars.sField1 = Me.cboField.Column(0, Me.cboField.ListIndex)
                varFldVars.sField2 = Me.cboField.Column(1, Me.cboField.ListIndex)
                varFldVars.sField3 = Me.cboField.Column(3, Me.cboField.ListIndex)
            End Select
        ElseIf varFldVars.bAllowNullOfField = False Then
            MsgBox "Must make a selection for this field", vbOKOnly
            Me.cboField.SetFocus
            Exit Sub
        End If
    Else
        If varFldVars.bAllowNullOfField = False And NZ(Me.txtField, "") = "" Then
            MsgBox "Blank field not allowed", vbOKOnly
            Me.txtField.SetFocus
            Exit Sub
        End If
        varFldVars.sField1 = Me.txtField
    End If
    pbHasBeenSaved = True
    CBA_bDataChg = True
    Call mw_RemoveBoxHook
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    Dim bPassOK As Boolean
    pbSetUp = True
   ' Pre-format the input parms
    Me.StartUpPosition = 0
    Me.cmdSave.Visible = False: pbHasBeenSaved = False: CBA_bDataChg = False
    Me.lblHdg.Caption = varFldVars.sHdg
    Me.Top = varFldVars.lFrmTop
    Me.Left = varFldVars.lFrmLeft
    Me.Width = varFldVars.lFldWidth * 4
    If varFldVars.lFldHeight > 0 Then Me.Height = varFldVars.lFldHeight + 300
    Me.cmdNA.Enabled = varFldVars.bAllowNullOfField
    If varFldVars.sType = "ComboBox" Then
        Me.cboField.Visible = True
        Me.cboField.Left = (varFldVars.lFldWidth / 2)
        Me.cboField.Width = varFldVars.lFldWidth
        Me.cboField.ColumnCount = varFldVars.lCols
''        Me.cboField.ColumnWidths = "0 pt;" & (varFldVars.lFldWidth * 4) & " pt"
        Me.cboField.ListWidth = varFldVars.lFldWidth
        If varFldVars.sSQL = "AST_WeeksOfSale" Then
            Me.cboField.ColumnCount = 1
            Me.cboField.ColumnWidths = varFldVars.lFldWidth
            varFldVars.lCols = 1
            Call AST_WeeksOfSale(Me.cboField)
        Else
            Me.cboField.ColumnWidths = "0;" & (varFldVars.lFldWidth)
            CBA_DBtoQuery = 1: Call g_EraseAry(CBA_ABIarr)
            bPassOK = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", "cbo_Fields", g_GetDB(varFldVars.sDB), CBA_MSAccess, varFldVars.sSQL, 120, , , False)
            If bPassOK = True Then
                Call AST_FillDDBox(Me.cboField, varFldVars.lCols)
            End If
        End If
    Else
        Me.txtField.Visible = True
        Me.txtField.Left = (varFldVars.lFldWidth / 2)
        Me.txtField.Width = varFldVars.lFldWidth
        If varFldVars.lFldHeight > 0 Then Me.txtField.Height = varFldVars.lFldHeight
    End If
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
    ' Remove any remnants of the mousewheel
    Call mw_RemoveBoxHook
End Sub
