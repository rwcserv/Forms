VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CBA_BTF_frm_Reporting 
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8055
   OleObjectBlob   =   "CBA_BTF_frm_Reporting.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "CBA_BTF_frm_Reporting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit       ' @CBA_BTF
Private RP As CBA_BTF_ReportParamaters
Private CGList As Variant
Private SCGList As Variant

Private Sub cbx_DBU_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call mw_SetBoxHook(cbx_DBU)
End Sub
Private Sub cbx_GBD_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call mw_SetBoxHook(cbx_GBD)
End Sub
Private Sub cbx_CG_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call mw_SetBoxHook(cbx_CG)
End Sub
Private Sub cbx_SCG_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call mw_SetBoxHook(cbx_SCG)
End Sub

Sub UserForm_Initialize()
    Dim a As Long, BDList, GBDList
    Me.lblVersion.Caption = CBA_getVersionStatus(g_GetDB("ForeCast"), CBA_FCAST_Ver, "Forcasting Tool", "FCast")  ' Get the latest version
    Me.StartUpPosition = 0
    Me.Top = Application.Top + (Application.Height / 2) - (Me.Height / 2)
    Me.Left = Application.Left + (Application.Width / 2) - (Me.Width / 2)
    If CBAR_getAdminUsers = True Then Me.cbx_ReportSelect.AddItem "Admin S&M Report"
    Me.cbx_ReportSelect.AddItem "CG/SCG Sales and Margin Forecast Report"
    Me.cbx_ReportSelect.AddItem "DBU Period Forecast Report"
    CBA_COM_SQLQueries.CBA_COM_GenPullSQL "CGList"
    CGList = CBA_CBISarr
    CBA_COM_SQLQueries.CBA_COM_GenPullSQL "SCGList"
    SCGList = CBA_CBISarr
    Erase CBA_CBISarr
    For a = LBound(CGList, 2) To UBound(CGList, 2)
        cbx_CG.AddItem CGList(0, a)
    Next
    CBA_COM_SQLQueries.CBA_COM_GenPullSQL "BDList"
    BDList = CBA_CBISarr
    Erase CBA_CBISarr
    For a = LBound(BDList, 2) To UBound(BDList, 2)
        If IsNull(BDList(0, a)) = False Then Me.cbx_BD.AddItem BDList(0, a)
    Next
    CBA_COM_SQLQueries.CBA_COM_GenPullSQL "GBDList"
    GBDList = CBA_CBISarr
    Erase CBA_CBISarr
    For a = LBound(GBDList, 2) To UBound(GBDList, 2)
        If IsNull(GBDList(0, a)) = False Then Me.cbx_GBD.AddItem GBDList(0, a)
    Next
    
    
    lbl_DBU.Visible = False
    cbx_DBU.Visible = False
    CBA_SQL_Queries.CBA_GenPullSQL "DBU_List"
    For a = LBound(CBA_ABIarr, 2) To UBound(CBA_ABIarr, 2)
        If Len(CBA_ABIarr(2, a)) <> 0 Then
            If Left(CBA_ABIarr(2, a), 1) = Chr(10) Then
                cbx_DBU.AddItem Mid(CBA_ABIarr(2, a), 2, Len(CBA_ABIarr(2, a)) - 1)
            ElseIf Right(CBA_ABIarr(2, a), 1) = Chr(10) Then
                cbx_DBU.AddItem Left(CBA_ABIarr(2, a), Len(CBA_ABIarr(2, a)) - 1)
            Else
                cbx_DBU.AddItem CBA_ABIarr(2, a)
            End If
        End If
    Next
    
    For a = 1 To 12
        cbx_PE_Month.AddItem a
        cbx_PS_Month.AddItem a
    Next
    For a = Year(Date) - 1 To Year(Date) + 5
        cbx_PS_Year.AddItem a
        cbx_PE_Year.AddItem a
    Next
    cbx_ProductClass.AddItem "1 - Core Range"
    cbx_ProductClass.AddItem "2 - Food Special"
    cbx_ProductClass.AddItem "3 - Non-Food Special"
    cbx_ProductClass.AddItem "4 - Seasonal"
End Sub
Private Sub cbx_ReportSelect_Change()
    RP.ReportName = cbx_ReportSelect.Value
    If RP.ReportName = "CG/SCG Sales and Margin Forecast Report" Then
        cbx_PS_Month.Visible = False: cbx_PE_Month.Visible = False: cbx_PE_Year.Visible = False: cbx_PS_Year.Visible = True
        Me.cbx_BD.Visible = True: Me.cbx_CG.Visible = True: Me.cbx_GBD.Visible = True: Me.cbx_ProductClass.Visible = True: Me.cbx_SCG.Visible = True: Me.cbx_DBU.Visible = False
        Me.Label3.Visible = True: Me.Label4.Visible = True: Me.Label5.Visible = True: Me.Label6.Visible = True: Me.Label11.Visible = True: Me.lbl_DBU.Visible = False
    ElseIf RP.ReportName = "DBU Period Forecast Report" Then
        cbx_PS_Month.Visible = True: cbx_PE_Month.Visible = True: cbx_PE_Year.Visible = True: cbx_PS_Year.Visible = True
        Me.cbx_BD.Visible = False: Me.cbx_CG.Visible = False: Me.cbx_GBD.Visible = False: Me.cbx_ProductClass.Visible = False: Me.cbx_SCG.Visible = False: Me.cbx_DBU.Visible = True
        Me.Label3.Visible = False: Me.Label4.Visible = False: Me.Label5.Visible = False: Me.Label6.Visible = False: Me.Label11.Visible = False: Me.lbl_DBU.Visible = True
    Else
        cbx_PS_Month.Visible = True: cbx_PE_Month.Visible = True: cbx_PE_Year.Visible = True: cbx_PS_Year.Visible = True
        Me.cbx_BD.Visible = True: Me.cbx_CG.Visible = True: Me.cbx_GBD.Visible = True: Me.cbx_ProductClass.Visible = True: Me.cbx_SCG.Visible = True: Me.cbx_DBU.Visible = False
        Me.Label3.Visible = True: Me.Label4.Visible = True: Me.Label5.Visible = True: Me.Label6.Visible = True: Me.Label11.Visible = True: Me.lbl_DBU.Visible = False
    End If
End Sub
Private Sub cbx_PS_Month_Change()
    RP.PSMonth = cbx_PS_Month.Value
End Sub
Private Sub cbx_PS_Year_Change()
    RP.PSYear = cbx_PS_Year.Value
End Sub
Private Sub cbx_PE_Month_Change()
    If cbx_PE_Month.Value = "" Then RP.PEMonth = 0 Else RP.PEMonth = cbx_PE_Month.Value
End Sub
Private Sub cbx_PE_Year_Change()
    If IsNumeric(cbx_PE_Year.Value) Then
        RP.PEYear = cbx_PE_Year.Value
    Else
        cbx_PE_Year.Value = ""
    End If
End Sub
Private Sub cbx_BD_Change()
    RP.BD = cbx_BD.Value
    If cbx_BD.Value <> "" Then
        cbx_GBD.Visible = False
        cbx_CG.Visible = False
        cbx_SCG.Visible = False
    Else
        cbx_GBD.Visible = True
        cbx_CG.Visible = True
        cbx_SCG.Visible = True
    End If
    
End Sub
Private Sub cbx_GBD_Change()
    RP.GBD = cbx_GBD.Value
    If cbx_GBD.Value <> "" Then
        cbx_BD.Visible = False
        cbx_CG.Visible = False
        cbx_SCG.Visible = False
    Else
        cbx_BD.Visible = True
        cbx_CG.Visible = True
        cbx_SCG.Visible = True
    End If
End Sub
Private Sub cbx_CG_Change()
    CGBoxChange
End Sub
Private Sub cbx_CG_AfterUpdate()
    CGBoxChange
End Sub
Private Sub CGBoxChange()
    Dim a As Long
    On Error GoTo Err_Routine
    If cbx_CG.Value = "" Then
        RP.CG = 0
    Else
        RP.CG = Trim(Mid(cbx_CG.Value, 1, InStr(1, cbx_CG.Value, " - ")))
    End If
    If InStr(1, cbx_CG.Value, " - ") = 0 Then
        cbx_SCG.Clear
        cbx_CG.Value = ""
        cbx_BD.Visible = True
        cbx_GBD.Visible = True
        Exit Sub
    End If
    cbx_SCG.Clear
    For a = LBound(SCGList, 2) To UBound(SCGList, 2)
        If SCGList(0, a) = Trim(Mid(cbx_CG.Value, 1, InStr(1, cbx_CG.Value, " - "))) Then cbx_SCG.AddItem SCGList(1, a)
    Next
    If cbx_CG.Value = "" Then
        cbx_BD.Visible = True
        cbx_GBD.Visible = True
    Else
        cbx_BD.Visible = False
        cbx_GBD.Visible = False
    End If
Err_Routine:
    Err.Clear
    On Error GoTo 0

End Sub
Private Sub cbx_SCG_Change()
    If InStr(1, cbx_SCG.Value, " - ") = 0 Then
        cbx_SCG.Value = ""
        Exit Sub
    End If
    RP.scg = Trim(Mid(cbx_SCG.Value, 1, InStr(1, cbx_SCG.Value, " - ")))
End Sub
Private Sub cbx_ProductClass_Change()
    RP.ProductClass = Trim(Mid(cbx_ProductClass.Value, 1, InStr(1, cbx_ProductClass.Value, " - ")))
End Sub

Private Sub but_Run_Click()
    CBA_BTF_Runtime.RunReport RP
End Sub
