VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CBA_AADD_frm_ActivityPlan 
   ClientHeight    =   11370
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8655
   OleObjectBlob   =   "CBA_AADD_frm_ActivityPlan.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "CBA_AADD_frm_ActivityPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'#RW Added new mousewheel routines 190701
Private Sub lbx_Prods_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call mw_SetBoxHook(lbx_Prods)
End Sub
Private Sub cbx_Promo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call mw_SetBoxHook(cbx_Promo)
End Sub

Private Sub UserForm_Terminate()
    Call but_Stop_Click
End Sub
Private Sub but_Stop_Click()
    Unload Me
End Sub
