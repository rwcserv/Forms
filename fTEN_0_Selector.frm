VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fTEN_0_Selector 
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7200
   OleObjectBlob   =   "fTEN_0_Selector.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fTEN_0_Selector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit         ' fTEN_0_Selector  Changed 03/07/2019

Private plLV_ID As Long, psPf As String, psPfVer As String, psPfProd As String, psPfCont As String, pbLinked As Boolean

Private Sub cmdLink_Click()
    pbLinked = True
    varFldVars.sField1 = Me.lstSelect.Column(0)
    varFldVars.sField2 = Me.lstSelect.Column(1)
    varFldVars.sField3 = Me.lstSelect.Column(2)
    Unload Me
End Sub

Private Sub cmdSearch_Click()
    If varFldVars.sType = "PF_ID" Then
        Call Ten_GetListbox(Me.lstSelect, varFldVars.sType, Me.txtSearch, plLV_ID)
    ElseIf varFldVars.sType = "PV_ID" Then
        Call Ten_GetListbox(Me.lstSelect, varFldVars.sType, Me.txtSearch, plLV_ID)
    ElseIf varFldVars.sType = "PD_ID" Then
        Call Ten_GetListbox(Me.lstSelect, varFldVars.sType, Me.txtSearch, plLV_ID)
    ElseIf varFldVars.sType = "PPD_ID" Then
        Call Ten_GetListbox(Me.lstSelect, varFldVars.sType, Me.txtSearch, plLV_ID)
    ElseIf varFldVars.sType = "PC_ID" Then
        Call Ten_GetListbox(Me.lstSelect, varFldVars.sType, Me.txtSearch, plLV_ID)
    End If
End Sub

Private Sub lstSelect_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.lstSelect.ListIndex > -1 Then
        Me.cmdLink.Enabled = True
        Call cmdLink_Click
    End If
End Sub

Private Sub lstSelect_Change()
    If Me.lstSelect.ListIndex > -1 Then
        Me.cmdLink.Enabled = True
    Else
        Me.cmdLink.Enabled = False
    End If
End Sub

Private Sub lstSelect_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    '%%RWCall mw_SetBoxHook(lstSelect)
End Sub


Private Sub UserForm_Initialize()
    pbLinked = False
    Me.StartUpPosition = 0
    Me.Top = Application.Top + (Application.Height / 2) - (Me.Height / 2)
    Me.Left = Application.Left + (Application.Width / 2) - (Me.Width / 2)

    plLV_ID = Val(varFldVars.sField4)
    If CStr(varFldVars.sType) = "PF_ID" Then
        Me.lblSelector.Caption = "Portfolio Selection"
        ''Me.lblName.Caption = "Portfolio Name"
        Me.lblDesc.Caption = "Enter text to search for the related portfolio" & vbCrLf & _
        "Select the desired record and click on Link"
    ElseIf CStr(varFldVars.sType) = "PV_ID" Then
        Me.lblSelector.Caption = "Version Selection"
        ''Me.lblName.Caption = "Version Name"
        Me.lblDesc.Caption = "Select the desired Portfolio version" & vbCrLf & _
        "Version is minimum requirement for tender document creation"
        Call cmdSearch_Click
    ElseIf CStr(varFldVars.sType) = "PPD_ID" Then
        Me.lblSelector.Caption = "Prior Product Selection"
        ''Me.lblName.Caption = "Prior Product Name"
        Me.lblDesc.Caption = "Select a product from previous version or search for a best-fit product" & vbCrLf & _
        "Search on description or product code"
        Call cmdSearch_Click
    ElseIf CStr(varFldVars.sType) = "PD_ID" Then
        Me.lblSelector.Caption = "Product Selection"
        ''Me.lblName.Caption = "Product Name"
        Me.lblDesc.Caption = "Search for a product to link to the tender document" & vbCrLf & _
        "Search on description or product code"
        Call cmdSearch_Click
    ElseIf CStr(varFldVars.sType) = "PC_ID" Then
        Me.lblSelector.Caption = "Contract Selection"
        ''Me.lblName.Caption = "Contract Name"
        Me.lblDesc.Caption = "Search for a contract to link to the tender document" & vbCrLf & _
        "Search on description or contract number"
        Call cmdSearch_Click
    End If
    Call lstSelect_Change
End Sub

Private Sub UserForm_Terminate()
    If pbLinked = False Then
        varFldVars.sField1 = ""
        varFldVars.sField2 = ""
        varFldVars.sField3 = ""
        varFldVars.sField4 = ""
    End If
End Sub
