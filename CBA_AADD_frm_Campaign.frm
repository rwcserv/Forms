VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CBA_AADD_frm_Campaign 
   ClientHeight    =   10470
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10320
   OleObjectBlob   =   "CBA_AADD_frm_Campaign.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "CBA_AADD_frm_Campaign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'#RW Added new mousewheel routines 190701
Private Sub lbx_Promos_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call mw_SetBoxHook(lbx_Promos)
End Sub
Private Sub lbx_Camps_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call mw_SetBoxHook(lbx_Camps)
End Sub
Sub UserForm_Initialize()
    Me.StartUpPosition = 0
    Me.Top = Application.Top + (Application.Height / 2) - (Me.Height / 2)
    Me.Left = Application.Left + (Application.Width / 2) - (Me.Width / 2)
    
    cbx_CampType.AddItem "Price"
    cbx_CampType.AddItem "MasterBrand"
    cbx_CampType.AddItem "Special Buys"
    cbx_CampType.AddItem "Campaigns"
    cbx_CampType.AddItem "Mobile"
    cbx_CampType.AddItem "AlwaysOn Search"
    cbx_CampType.AddItem "AlwaysOn Social"
    cbx_CampType.AddItem "Holidays"
    


End Sub
Private Sub UserForm_Terminate()
    Call but_Stop_Click
End Sub
Private Sub but_Stop_Click()
    Unload Me
End Sub



