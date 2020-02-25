VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Splash 
   Caption         =   "Please wait..."
   ClientHeight    =   2145
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17655
   OleObjectBlob   =   "frm_Splash.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''Private Sub UserForm_Activate()
''    Call SplashForm("Init")
''End Sub


Public Sub SplashForm(Optional sWhen As String = "2nd")
    Static sLast As String
    If CBA_strAldiMsg = "" Then
''        Application.OnTime Now + TimeValue("00:00:01"), "Quitform"
    ElseIf sLast <> CBA_strAldiMsg Then
''        Application.Wait (Now + TimeValue("00:00:02"))
        If sWhen = "Init" Then
            frm_Splash.lblSplash.Caption = CBA_strAldiMsg
        Else
            frm_Splash.lblSplash2.Caption = CBA_strAldiMsg
        End If
        frm_Splash.Repaint
        sLast = CBA_strAldiMsg
    Else
''        Application.Wait (Now + TimeValue("00:00:02"))
    End If
End Sub

Private Sub UserForm_Initialize()
    Me.Top = Application.Top ''+ ((Application.Height / 4) * 2.75) '- (Me.Height / 4)
    Me.Left = Application.Left ''+ (Application.Width / 2) - (Me.Width / 2)
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Stop any exit if not complete
    If CBA_strAldiMsg > "" Then Cancel = True
End Sub
