VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CBA_frm_SelectBaseData 
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4065
   OleObjectBlob   =   "CBA_frm_SelectBaseData.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "CBA_frm_SelectBaseData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub pic_WW_Click()
    CCM_Runtime.CCM_setDefaultDataset 1
    Unload Me
End Sub
Private Sub pic_Coles_Click()
    CCM_Runtime.CCM_setDefaultDataset 2
    Unload Me
End Sub
Private Sub pic_DM_Click()
    CCM_Runtime.CCM_setDefaultDataset 3
    Unload Me
End Sub
Private Sub pic_FC_Click()
    CCM_Runtime.CCM_setDefaultDataset 4
    Unload Me
End Sub

Sub UserForm_Initialize()
    Me.StartUpPosition = 0
    Me.Top = Application.Top + (Application.Height / 2) - (Me.Height / 2)
    Me.Left = Application.Left + (Application.Width / 2) - (Me.Width / 2)
End Sub
Private Sub UserForm_Terminate()
'    If CCM_Runtime.CMM_getDefaultDataset = 0 Then
'        CCM_Runtime.CCM_setDefaultDataset 1
'    End If
End Sub
