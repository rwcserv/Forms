VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CBA_BT_frm_ALDT 
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12585
   OleObjectBlob   =   "CBA_BT_frm_ALDT.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "CBA_BT_frm_ALDT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'#RW Added new mousewheel routines 190701
Private Sub lbx_ProdsToRun_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call mw_SetBoxHook(lbx_ProdsToRun)
End Sub

Private Sub btn_OK_Click()
    CBA_BT_ALDT.pullData
End Sub
Private Sub but_Add_Click()
    If isDate(Me.tbx_DFrom) = False Or Me.tbx_DFrom = "" Then GoTo EntryError
    If isDate(Me.tbx_DTo) = False Or Me.tbx_DTo = "" Then GoTo EntryError
    If IsNumeric(Me.tbx_Prod) = False Or Len(Me.tbx_Prod) < 4 Or Len(Me.tbx_Prod) > 10 Or Me.tbx_Prod = "" Then GoTo EntryError
    If DateDiff("D", Me.tbx_DTo, Me.tbx_DFrom) > 0 Then GoTo EntryError
    CBA_SQL_Queries.CBA_GenPullSQL "CBIS_ProductDesc", , , CLng(Me.tbx_Prod)
    Me.lbx_ProdsToRun.AddItem Me.tbx_Prod
    Me.lbx_ProdsToRun.List(Me.lbx_ProdsToRun.ListCount - 1, 1) = CBA_CBISarr(0, 0)
    Me.lbx_ProdsToRun.List(Me.lbx_ProdsToRun.ListCount - 1, 2) = Me.tbx_DFrom
    Me.lbx_ProdsToRun.List(Me.lbx_ProdsToRun.ListCount - 1, 3) = Me.tbx_DTo
    Me.tbx_Prod = ""
    Exit Sub
EntryError:
        MsgBox "One or more entered values are invalid"
End Sub
Private Sub lbx_ProdsToRun_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim a As Long
    For a = 0 To Me.lbx_ProdsToRun.ListCount - 1
        If Me.lbx_ProdsToRun.Selected(a) = True Then
            Me.lbx_ProdsToRun.RemoveItem (a)
            Exit For
        End If
    Next
End Sub

Sub UserForm_Initialize()
    Me.StartUpPosition = 0
    Me.Top = Application.Top + (Application.Height / 2) - (Me.Height / 2)
    Me.Left = Application.Left + (Application.Width / 2) - (Me.Width / 2)
End Sub




