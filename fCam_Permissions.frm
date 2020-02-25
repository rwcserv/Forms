VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fCAM_Permissions 
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9285
   OleObjectBlob   =   "fCam_Permissions.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "fCam_Permissions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                        ' fCam_Permissions
Private pbIsActive As Boolean
Private psUserList As Scripting.Dictionary
Private psAppUser As String
Private psOwnerUser As String
Private psdUsrAssign As Scripting.Dictionary

' Sets the permissions of people that can see the category review
Sub UserForm_Initialize()
    ' Does put it under Activate in doc, but... Get UserList and  assigned users from the runtime. Populate the listboxes
    Dim lTop, lLeft, lRow, lcol As Long
    Me.StartUpPosition = 0
    Me.Top = Application.Top + (Application.Height / 2) - (Me.Height / 2)
    Me.Left = Application.Left + (Application.Width / 2) - (Me.Width / 2)
    Me.lblVersion.Caption = CBA_getVersionStatus(g_GetDB("Cam"), CBA_Cam_Ver, "Camera Tool", "Camera")  ' Get the latest version
''    Call CBA_getUserShortTitle("", "", CBA_SetUser, psAppUser)
    IsActive = True
    'Me.Hide
End Sub
Function Loadup(ByVal usrName As String, ByVal UList As Scripting.Dictionary, ByVal AddedUsers As Scripting.Dictionary, ByVal OUsr As String)
    Dim U As Variant
    Set psUserList = UList
    psAppUser = usrName
    psOwnerUser = OUsr
    Set psdUsrAssign = AddedUsers
    For Each U In psUserList.Keys
        If psOwnerUser = U Then
            Me.lst_AppUsers.AddItem CStr(U), 0
        Else
            If psdUsrAssign.Exists(U) Then
                Me.lst_AppUsers.AddItem CStr(U), Me.lst_AppUsers.ListCount
            Else
                Me.lst_AvailUsers.AddItem CStr(U)
            End If
        End If
    Next
    Me.Show
End Function
Private Sub cmd_Add_Click()
    ' Send new assignment to the runtime data objects and run populateListboxes
    SwapListBoxItem Me.lst_AvailUsers, Me.lst_AppUsers, psAppUser
End Sub
Private Sub cmd_Remove_Click()
    '  Remove assignment from the runtime data objects and run populateListboxes
    SwapListBoxItem Me.lst_AppUsers, Me.lst_AvailUsers, psAppUser
End Sub
Private Sub cmd_SetOwner_Click()
    ' Only the owner can allocate a new owner. At which point they remain assigned, but to a different owner.
    ' Sends to Runtime data objects and runs populateListboxes
    Dim a As Long, cnt As Long, sel As Long
    Dim TooMany As Boolean
    If psAppUser = psOwnerUser Then
        sel = 0
        For a = 0 To Me.lst_AppUsers.ListCount - 1
            If Me.lst_AppUsers.Selected(a) = True Then
                If sel = 0 Then sel = a Else TooMany = True: Exit For
            End If
        Next
        If TooMany = True Then
            MsgBox "Please select only one user to become the Category Review Owner", vbOKOnly
        Else
            If psdUsrAssign.Exists(psOwnerUser) = False Then psdUsrAssign.Add psOwnerUser, psUserList(psOwnerUser)
            psOwnerUser = Me.lst_AppUsers.List(sel)
            mCAM_Runtime.setOwnerUser psOwnerUser
            Me.lst_AppUsers.RemoveItem sel
            Me.lst_AppUsers.AddItem psOwnerUser, 0
        End If
    Else
        MsgBox "Only the Owner can assign a new Owner, Change not applied", vbOKOnly
    End If

End Sub

Function SwapListBoxItem(ByRef list1 As Object, ByRef list2 As Object, Optional ByVal sException As String)
    ' Send new assignment to or remove an assignment from the appropriate listbox
    Dim a As Long
    Dim newcol As Collection
    Dim PCde As String

    For a = 0 To list1.ListCount - 1
        If list1.Selected(a) = True Then
            list2.AddItem list1.List(a)

        End If
    Next
    For a = 0 To list1.ListCount - 1
        If list1.ListCount > a Then
            If sException <> "" Then
                If list1.Selected(a) = True And list1.List(a) <> sException Then
                    list1.RemoveItem (a)
                    a = a - 1
                End If
            Else
                If list1.Selected(a) = True Then
                    list1.RemoveItem (a)
                    a = a - 1
                End If
            End If
        Else
            Exit For
        End If
    Next

End Function
Sub UserForm_Terminate()
    Dim a As Long
    Dim yn As Long
    Dim sdDic As Scripting.Dictionary
    yn = MsgBox("Confirm Changes?", vbYesNo)
    If yn = 6 Then
        Set sdDic = New Scripting.Dictionary
        For a = 0 To Me.lst_AppUsers.ListCount - 1
            If a = 0 Then
                sdDic.Add Me.lst_AppUsers.List(a), 1
            Else
                sdDic.Add Me.lst_AppUsers.List(a), 0
            End If
        Next
        If sdDic.Count > -1 Then Call mCAM_Runtime.ChangePermissions(sdDic)
    End If
End Sub

Public Property Get IsActive() As Boolean: IsActive = pbIsActive: End Property
Private Property Let IsActive(ByVal bNewValue As Boolean): pbIsActive = bNewValue: End Property

