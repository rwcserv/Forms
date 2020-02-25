VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fTEN_ADMIN 
   ClientHeight    =   12300
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11385
   OleObjectBlob   =   "fTEN_ADMIN.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "fTEN_ADMIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub but_SelectAll_Click()
'Dim a As Long
'    For a = 0 To Me.box_Details.ListCount - 1
'        Me.box_Details.Selected(a) = True
'    Next
'End Sub
Sub UserForm_Terminate()

    

End Sub
Private Sub btn_OK_Click()
Unload Me
End Sub
Sub UserForm_Initialize()
    Dim lTop, lLeft, lRow, lcol As Long
    Me.StartUpPosition = 0
    Me.Top = Application.Top + (Application.Height / 2) - (Me.Height / 2)
    Me.Left = Application.Left + (Application.Width / 2) - (Me.Width / 2)
    Me.Hide
End Sub

