VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CBA_BT_ReorderMAA 
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6825
   OleObjectBlob   =   "CBA_BT_ReorderMAA.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "CBA_BT_ReorderMAA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btn_OK_Click()
    Dim SellT As Single
    
    If isDate(Me.tbx_OSD) = False Or Me.tbx_OSD = "" Then GoTo EntryError
    If IsNumeric(Me.tbx_PCode) = False Or Len(Me.tbx_PCode) < 4 Or Len(Me.tbx_PCode) > 10 Or Me.tbx_PCode = "" Then GoTo EntryError
    If InStr(1, Me.tbx_SellT, "%") > 0 Then Me.tbx_SellT = Mid(Me.tbx_SellT, 1, InStr(1, Me.tbx_SellT, "%") - 1)
    If IsNumeric(Me.tbx_SellT) = False Or Me.tbx_SellT = "" Then GoTo EntryError
    
    If Me.tbx_SellT > 30 And Me.tbx_SellT < 110 Then
        SellT = Me.tbx_SellT / 100
    Else
        SellT = Me.tbx_SellT
    End If
    If IsNumeric(Me.tbx_MU) = False Then GoTo EntryError
    
    
    
    Unload Me
    CBA_BT_ReorderRuntime.ReorderRuntime Me.tbx_OSD, Me.tbx_PCode, SellT, Me.tbx_MU, Me.cbx_MultiDropCalc
    
    
    
    Exit Sub
EntryError:
        MsgBox "One or more entered values are invalid"
End Sub

Sub UserForm_Initialize()
'    Dim lTop, lLeft, lRow, lCol As Long
     
    Me.StartUpPosition = 0
    Me.Top = Application.Top + (Application.Height / 2) - (Me.Height / 2)
    Me.Left = Application.Left + (Application.Width / 2) - (Me.Width / 2)
        
End Sub




