VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CBA_frm_CCS 
   ClientHeight    =   8820
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11700
   OleObjectBlob   =   "CBA_frm_CCS.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "CBA_frm_CCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'#RW Added new mousewheel routines 190701
Private Sub box_CCS_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call mw_SetBoxHook(box_CCS)
End Sub

Sub UserForm_Initialize()
    Dim LB
     
    Me.StartUpPosition = 0
    Me.Top = Application.Top + (Application.Height / 2) - (Me.Height / 2)
    Me.Left = Application.Left + (Application.Width / 2) - (Me.Width / 2)

    Set LB = Me.box_CCS
    With LB
        .Clear
        .List = CBA_BasicFunctions.CBA_TransposeArray(CBA_COM_Runtime.getCCS_Results)
    End With

End Sub
Private Sub but_DSelAll_Click()
    Dim i As Long
    For i = 0 To Me.box_CCS.ListCount - 1
        Me.box_CCS.Selected(i) = False
    Next i
End Sub
Private Sub but_SelAll_Click()
    Dim i As Long
    For i = 0 To Me.box_CCS.ListCount - 1
        Me.box_CCS.Selected(i) = True
    Next i
End Sub
Private Sub but_ViewMT_Click()
Dim i As Long
Dim colWW As Collection, colColes As Collection, colDM As Collection, colFC As Collection
Dim WW As Boolean, Coles As Boolean, DM As Boolean, FC As Boolean
    
Set colWW = New Collection
Set colColes = New Collection
Set colDM = New Collection
Set colFC = New Collection
    
    

    
    For i = 0 To Me.box_CCS.ListCount - 1
        If Me.box_CCS.Selected(i) = True Then
            Select Case Me.box_CCS.List(i, 0)
                Case "Woolworths"
                    WW = True
                    colWW.Add Me.box_CCS.List(i, 1)
                Case "Coles"
                    Coles = True
                    colColes.Add Me.box_CCS.List(i, 1)
                Case "Dan Murphys"
                    DM = True
                    colDM.Add Me.box_CCS.List(i, 1)
                Case "First Choice"
                    FC = True
                    colFC.Add Me.box_CCS.List(i, 1)
            End Select
        End If
    Next

If Me.box_CCS.ListCount - 1 > 1300 Then If CBA_BasicFunctions.isRunningSheetDisplayed = False Then CBA_BasicFunctions.CBA_Running
Unload Me


If WW = True Then CCM_UDWWSKU = CBA_COM_SetupSKUArray.CBA_SetupSKUArray("WW", colWW)
If Coles = True Then CCM_UDColesSKU = CBA_COM_SetupSKUArray.CBA_SetupSKUArray("C", colColes)
If DM = True Then CCM_UDDMSKU = CBA_COM_SetupSKUArray.CBA_SetupSKUArray("DM", colDM)
If FC = True Then CCM_UDFCSKU = CBA_COM_SetupSKUArray.CBA_SetupSKUArray("FC", colFC)


CBA_COM_frm_MatchingTool.setCCM_UserDefinedState (True)
CCM_Runtime.CCM_MatchingSelectorActivate
CBA_BasicFunctions.CBA_Close_Running
End Sub




