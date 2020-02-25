VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CBA_CDS_frm 
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8340
   OleObjectBlob   =   "CBA_CDS_frm.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "CBA_CDS_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ReadyToRun As Boolean
Private Sub btn_OK_Click()
Dim bfound As Boolean
    If isDate(Me.tbx_DFrom) = False Or Me.tbx_DFrom = "" Then GoTo EntryError
    If isDate(Me.tbx_DTo) = False Or Me.tbx_DTo = "" Then GoTo EntryError
    bfound = False
    If CBA_CDS_Runtime.CBA_CDS_SetParamaters(Me.tbx_DFrom, Me.tbx_DTo, Me.ListBox1.Selected(0), _
        Me.ListBox1.Selected(1), Me.ListBox1.Selected(3), Me.ListBox1.Selected(4), Me.ListBox1.Selected(2)) = True Then
        ReadyToRun = True
        Unload Me
        'CBA_CDS_Runtime.ProceedToCDSReport
    Else
        MsgBox "Non Schemes Selected"
        
    End If
    
    Exit Sub
EntryError:
        MsgBox "Non Date Value Entered"
End Sub

Sub UserForm_Initialize()
'    Dim lTop, lLeft, lRow, lCol As Long
    ReadyToRun = False
    Me.StartUpPosition = 0
    Me.Top = Application.Top + (Application.Height / 2) - (Me.Height / 2)
    Me.Left = Application.Left + (Application.Width / 2) - (Me.Width / 2)
    Me.ListBox1.AddItem "NSW"
    Me.ListBox1.AddItem "QLD"
    Me.ListBox1.AddItem "EXP_QLD"
    Me.ListBox1.AddItem "ACT"
    Me.ListBox1.AddItem "EXP_ACT"
End Sub
Private Sub UserForm_Terminate()
    If ReadyToRun = False Then CBA_CDS_Runtime.CBA_CDS_SetParamaters 0, 0, False, False, False, False, False
End Sub
