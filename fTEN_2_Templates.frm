VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fTEN_2_Templates 
   ClientHeight    =   10695
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5820
   OleObjectBlob   =   "fTEN_2_Templates.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "fTEN_2_Templates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit     ' @fTEN_2_Templates 29/07/2019

Private psSQL As String, pbPassOk As Boolean, pbSetUp As Boolean, bNoChg As Boolean, pbAddIP As Boolean, pbUpdIP As Boolean, pbRfrshIP As Boolean
Private plSH_ID As Long, psSH_Desc As String, psNPD_R_MSO As String

Private Sub lstTemplates_Click()
    ' Get the template selected
    Static bRecursed As Boolean
    If pbSetUp Then Exit Sub
    If bRecursed Then Exit Sub
    bRecursed = True
    If Me.lstTemplates.ListIndex > -1 Then
        plSH_ID = Me.lstTemplates.Column(0)
        psSH_Desc = Me.lstTemplates.Column(1)
        varFldVars.lFrmLeft = g_PosForm(0, (varFldVars.lFldWidth * 4), 0, "Left")
        varFldVars.lFrmTop = g_PosForm(0, 0, 0, "Top")
        varFldVars.lID2 = plSH_ID
        varFldVars.sField4 = psSH_Desc
        Unload fTEN_2_Templates
        fTEN_5_Tenders.Show vbModal
    Else
        plSH_ID = 0
    End If
    bRecursed = False
''    call mTEN_Runtime.FilePrint( "Exit frm2_Templates-lstTemplates_Click", , , True
''    DoEvents
End Sub

Private Sub UserForm_Initialize()
    pbSetUp = True
    Me.StartUpPosition = 0
    Me.Top = Application.Top + (Application.Height / 2) - (Me.Height / 2)
    Me.Left = Application.Left + (Application.Width / 2) - (Me.Width / 2)
    
    Me.lblVersion.Caption = varFldVars.sField2
    ''Call g_PosForm(Me.Top, Me.Width, Me.Left, , True)
    Call p_FillTemplatesListBox
    pbSetUp = False
End Sub

Sub p_FillTemplatesListBox()
    ' Fill the Templates List Box
    Dim lIdx As Long, sTDocs As String, sSelDocs As String, sSep As String      '', sWhere As String, sAnd As String
    Me.MousePointer = fmMousePointerHourGlass
    sTDocs = "-" & varFldVars.sField3 & "-"
    If InStr(1, sTDocs, "-NPD-") > 0 Or (InStr(1, sTDocs, "-R-") = 0 And InStr(1, sTDocs, "-MSO-") = 0) Then
        sSelDocs = "1"
        sSep = ","
    End If
    If InStr(1, sTDocs, "-R-") > 0 Or InStr(1, sTDocs, "-MSO-") = 0 Then
        sSelDocs = sSelDocs & sSep & "2"
        sSep = ","
    End If
''    If InStr(1, sTDocs, "-MSO-") > 0 Or InStr(1, sTDocs, "-MSO-") = 0 Then
    sSelDocs = sSelDocs & sSep & "3"
''    End If
    
    psSQL = "SELECT * FROM C1_Seg_Template_Hdrs WHERE SH_Sts_ID = 1 AND SH_SysType='T' and SH_ID IN (" & sSelDocs & ");"
    CBA_DBtoQuery = 1: Call g_EraseAry(CBA_ABIarr)
    pbPassOk = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", "Ten_Query", g_GetDB("Ten"), CBA_MSAccess, psSQL, 120, , , False)
    Me.lstTemplates.Clear
    If pbPassOk = True Then
        Call AST_FillListBox(Me.lstTemplates, 2)
    End If
    CBA_DBtoQuery = 3
    bNoChg = False
    Me.MousePointer = fmMousePointerDefault
End Sub
