VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fCAM_FrontForm 
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9495
   OleObjectBlob   =   "fCam_FrontForm.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "fCam_FrontForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit             'fCam_FrontForm
Private pbIsActive As Boolean
Private pvOrigCatRevData() As Variant
Private psdEmptoNameUserList As Scripting.Dictionary
Private psdCatRevDataMap As Scripting.Dictionary
Private psAppUser As String

' Used to manage which individual category reviews exist for the user
Private Sub cmd_New_Click() ' @RWCam More to do?
    ' Hide and Create new Document object in CBA_Document Holder (using correct docutype).
    ' As part of the document creation for CATREV it must run the fCam_Parameters (but this is handled by the document holder)
    Me.Hide
    fCAM_Parameters.Show
End Sub

Private Sub cmd_Open_Click()    ' @RWCam More to do?
    ' Hide and Ask the document holder to render the document
    Me.Hide

End Sub

Private Sub cmd_Permissions_Click()
    ' If listbox element selected, hide, and ask Runtime to open permissions userform
    Dim a As Long
    For a = 0 To Me.lst_CatRev.ListCount - 1
        If lst_CatRev.Selected(a) = True Then
            fCam_Permissions.Show
        End If
    Next
End Sub

Private Sub UserForm_Initialize()
    Dim lTop, lLeft, lRow, lcol As Long
    Me.StartUpPosition = 0
    Me.Top = Application.Top + (Application.Height / 2) - (Me.Height / 2)
    Me.Left = Application.Left + (Application.Width / 2) - (Me.Width / 2)
    Me.lblVersion.Caption = CBA_getVersionStatus(g_GetDB("Cam"), CBA_Cam_Ver, "Camera Tool", "Camera")  ' Get the latest version
    ' @RWCam-How to Init
''    Me.lst_CatRev.AddItem "Fish 21/09/2019"
''    Me.lst_CatRev.AddItem "ShellFish 15/10/2019"
    IsActive = True
End Sub

Private Sub lst_CatRev_Click()
    Dim a As Long, bfound As Boolean, U
    For a = 0 To lst_CatRev.ListCount - 1
        If lst_CatRev.Selected(a) = True Then
            For Each U In sdCatRevDataMap.Keys
                If Me.lst_CatRev.List(a) = pvOrigCatRevData(5, U) & " - " & pvOrigCatRevData(4, U) Then
                    Call mCAM_Runtime.setCurID(pvOrigCatRevData(0, U), sdEmptoNameUserList(CStr(pvOrigCatRevData(2, U))))
                    bfound = True
                    Exit Sub
                End If
            Next
        End If
    Next
    If bfound = False Then mCAM_Runtime.setCurID -1, ""
End Sub

Public Sub GenerateAndDisplay(ByRef CatRevListData() As Variant, ByVal UserList As Scripting.Dictionary, ByVal EmpNotoNameUserList As Scripting.Dictionary, ByVal AUser As String)
    ' Puts values into the listbox
    Dim a As Long, b As Long, U As Variant, lTestEmpno As Long, bIsOwner As Boolean
    On Error GoTo Err_Routine
    CBA_Error = ""
    pvOrigCatRevData = CatRevListData
    Set sdEmptoNameUserList = EmpNotoNameUserList

    Set sdCatRevDataMap = New Scripting.Dictionary
    sAppUser = AUser
    If g_isArrayWData(pvOrigCatRevData) Then
        For a = LBound(pvOrigCatRevData, 2) To UBound(pvOrigCatRevData, 2)
            For b = LBound(pvOrigCatRevData, 1) To UBound(pvOrigCatRevData, 1)
    
                If b = 1 Then
                    bIsOwner = False
''                    If CStr(pvOrigCatRevData(b, a)) = UserList(sAppUser) Then sdCatRevDataMap.Add CStr(a), True: Exit For
                    If EmpNotoNameUserList.Exists(CStr(pvOrigCatRevData(b, a))) Then bIsOwner = True
                        If bIsOwner = True And AUser = EmpNotoNameUserList(CStr(pvOrigCatRevData(b, a))) Then
                        sdCatRevDataMap.Add CStr(a), True '': Exit For
                    ElseIf CBA_TestIP = "Y" Then                        '@RWCam - Added this so can be tested???? Will it work
                        If lTestEmpno < 10 Then
                            lTestEmpno = lTestEmpno + 1
                            sdCatRevDataMap.Add CStr(a), True '': Exit For
                        End If
                    Else
                        Exit For
                    End If
                ElseIf b > 7 Then
'                    If IsNull(pvOrigCatRevData(b, a)) Then Exit For
'                    If CStr(pvOrigCatRevData(b, a)) = UserList(sAppUser) Then sdCatRevDataMap.Add CStr(a), False
                    If NZ((pvOrigCatRevData(b, a)), 0) = 0 Then Exit For            ' @RWCam Added a formatted unique key by including the EmpNo but it could be the var 'b'
                    If EmpNotoNameUserList.Exists(CStr(pvOrigCatRevData(b, a))) Then sdCatRevDataMap.Add CStr(g_Fmt_2_IDs(a, pvOrigCatRevData(b, a), 7, 3)), False
                End If
            Next
        Next
        a = -1
        For Each U In sdCatRevDataMap.Keys
            a = a + 1
            Me.lst_CatRev.AddItem
            Me.lst_CatRev.List(a) = pvOrigCatRevData(5, U) & " - " & pvOrigCatRevData(4, U)
        Next
''    Else
''        Me.cmd_Permissions.Enabled = False
    End If
    Me.Show
Exit_Routine:

    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-fCam_FrontForm.GenerateAndDisplay", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Cam", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Sub

'Sub UserForm_Terminate()
'
'
'
'End Sub

Public Property Get IsActive() As Boolean: IsActive = pbIsActive: End Property
Private Property Let IsActive(ByVal bNewValue As Boolean): pbIsActive = bNewValue: End Property

Private Property Get sdEmptoNameUserList() As Scripting.Dictionary: Set sdEmptoNameUserList = psdEmptoNameUserList: End Property
Private Property Set sdEmptoNameUserList(ByVal objNewValue As Scripting.Dictionary): Set psdEmptoNameUserList = objNewValue: End Property

Private Property Get sdCatRevDataMap() As Scripting.Dictionary: Set sdCatRevDataMap = psdCatRevDataMap: End Property
Private Property Set sdCatRevDataMap(ByVal objNewValue As Scripting.Dictionary): Set psdCatRevDataMap = objNewValue: End Property

Private Property Get sAppUser() As String: sAppUser = psAppUser: End Property
Private Property Let sAppUser(ByVal sNewValue As String): psAppUser = sNewValue: End Property
