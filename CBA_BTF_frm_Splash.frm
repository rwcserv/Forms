VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CBA_BTF_frm_Splash 
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5880
   OleObjectBlob   =   "CBA_BTF_frm_Splash.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "CBA_BTF_frm_Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit       ' @CBA_BTF Changed 181224

Private Sub but_Run_Click()
    If Me.optCG.Value = True Then
        Call CrtForecastingWorkbook("optCG")
    ElseIf Me.optSCG.Value = True Then
        Call CrtForecastingWorkbook("optSCG")
    'ElseIf Me.optProduct.Value = True Then
        'Call CrtForecastingWorkbook("optProduct")
    Else
        MsgBox "No Reporting Method has been selected", vbOKOnly, "Method Warning"
        Exit Sub
    End If
End Sub
Private Sub UserForm_Initialize()
    Me.lblVersion.Caption = CBA_getVersionStatus(g_GetDB("ForeCast", , , True), CBA_FCAST_Ver, "Forcasting Tool", "FCast")
End Sub
