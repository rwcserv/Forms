VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CBA_POS_StoreSalesForm 
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9015
   OleObjectBlob   =   "CBA_POS_StoreSalesForm.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "CBA_POS_StoreSalesForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub UserForm_Terminate()
    If Me.bx501.Value = True Then CBA_POSQuery.CBA_POS_addtoRegionCol 501
    If Me.bx502.Value = True Then CBA_POSQuery.CBA_POS_addtoRegionCol 502
    If Me.bx503.Value = True Then CBA_POSQuery.CBA_POS_addtoRegionCol 503
    If Me.bx504.Value = True Then CBA_POSQuery.CBA_POS_addtoRegionCol 504
    If Me.bx505.Value = True Then CBA_POSQuery.CBA_POS_addtoRegionCol 505
    If Me.bx506.Value = True Then CBA_POSQuery.CBA_POS_addtoRegionCol 506
    If Me.bx507.Value = True Then CBA_POSQuery.CBA_POS_addtoRegionCol 507
    If Me.bx509.Value = True Then CBA_POSQuery.CBA_POS_addtoRegionCol 509

End Sub
Private Sub btn_OK_Click()
Unload Me
End Sub
Sub UserForm_Initialize()
Dim Found501 As Boolean, Found502 As Boolean, Found503 As Boolean, Found504 As Boolean, Found505 As Boolean, Found506 As Boolean, Found507 As Boolean, Found509 As Boolean
    
    Me.StartUpPosition = 0
    Me.Top = Application.Top + (Application.Height / 2) - (Me.Height / 2)
    Me.Left = Application.Left + (Application.Width / 2) - (Me.Width / 2)
' Not sure that the following are doing anything????
Found501 = False
Found502 = False
Found503 = False
Found504 = False
Found505 = False
Found506 = False
Found507 = False
Found509 = False


    
End Sub




