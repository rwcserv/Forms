VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CBA_BT_ICFix 
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7005
   OleObjectBlob   =   "CBA_BT_ICFix.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "CBA_BT_ICFix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private valtobtn(1 To 3) As Long
Private Sub btn_OK_Click()
    Dim this
    Err.Clear
    On Error Resume Next
    valtobtn(1) = CLng(box_Header)
    valtobtn(2) = CLng(box_Columns)
    valtobtn(3) = CLng(box_Rows)
    If Err <> 0 Then
        MsgBox "An invalid value has been entered. Please try again", vbOKOnly
        Err.Clear
        On Error GoTo 0
    Else
        On Error GoTo 0
        this = CBA_BT_BasicTools.CBA_ICFix(valtobtn())
    End If
    
    Unload Me

End Sub
Sub UserForm_Initialize()
    Dim lTop, lLeft, lRow, lcol As Long
     
    With ActiveWindow.VisibleRange
        lRow = .Rows.Count / 1.8
        lcol = .Columns.Count / 1.8
    End With
    With Cells(lRow, lcol)
        lTop = .Top
        lLeft = .Left
    End With
    With Me
        .Top = lTop
        .Left = lLeft
    End With
    

    
End Sub




