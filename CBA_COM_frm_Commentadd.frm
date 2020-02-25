VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CBA_COM_frm_Commentadd 
   ClientHeight    =   8820
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10365
   OleObjectBlob   =   "CBA_COM_frm_Commentadd.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "CBA_COM_frm_Commentadd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private APCode As Long
Private Sub but_Select_Click()
Dim valud(1 To 3) As String
Dim nocoms As Long
Dim ocoms As Boolean, lRet As Long, bOutput As Boolean, strSQL As String, ADesc

ocoms = False
    CBA_DBtoQuery = 1
    strSQL = "SELECT A_Code,COM_Comments1 , COM_Comments2 , COM_Comments3 FROM COM_Comments where A_Code = " & APCode
    bOutput = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", "ABI_QUERY", CBA_BSA & "LIVE DATABASES\ABI.accdb", CBA_MSAccess, strSQL, 120, , , False)  'Runs DB_Connection module to create connection to dtabase and run query
    If bOutput = True Then
        If Me.box_Comment.Value = CBA_ABIarr(1, 0) & " " & CBA_ABIarr(2, 0) & " " & CBA_ABIarr(3, 0) Then
            GoTo noupdaterqd
        End If
        If CBA_ABIarr(0, 0) <> "" And CBA_ABIarr(1, 0) & CBA_ABIarr(2, 0) & CBA_ABIarr(3, 0) = "" Then
            CBA_DBtoQuery = 2
            strSQL = "DELETE FROM COM_Comments Where A_Code = " & APCode
            bOutput = CBA_DB_Connect.CBA_DB_CC_NonC("UPDATE", "ABI_QUERY", CBA_BSA & "LIVE DATABASES\ABI.accdb", CBA_MSAccess, strSQL, 120, , , False)  'Runs DB_Connection module to create connection to dtabase and run query
        Else
            If IsNull(CBA_ABIarr(3, 0)) And IsNull(CBA_ABIarr(2, 0)) And IsNull(CBA_ABIarr(1, 0)) Then ocoms = False Else ocoms = True
        End If
    End If

valud(1) = Me.box_Comment.Value
If Len(valud(1)) > 60000 Then
    nocoms = 2
    valud(2) = Mid(valud(1), 60001)
    If Len(valud(2)) > 60000 Then
        nocoms = 3
        valud(3) = Mid(valud(2), 60001)
        valud(2) = Mid(valud(2), 1, 60000)
    End If
    valud(1) = Mid(valud(1), 1, 60000)
ElseIf Len(valud(1)) > 0 Then
    nocoms = 1
ElseIf Len(valud(1)) = 0 Then
    If ocoms = True Then
        lRet = MsgBox("Would you like to delete the comment that exists for " & APCode & "-" & ADesc, vbYesNo)
        If lRet = 6 Then
            CBA_DBtoQuery = 2
            strSQL = "DELETE FROM COM_Comments Where A_Code = " & APCode
            bOutput = CBA_DB_Connect.CBA_DB_CC_NonC("UPDATE", "ABI_QUERY", CBA_BSA & "LIVE DATABASES\ABI.accdb", CBA_MSAccess, strSQL, 120, , , False)  'Runs DB_Connection module to create connection to dtabase and run query
        End If
    End If
End If


If ocoms = True Then
    'update if need to change
    CBA_DBtoQuery = 2
    If nocoms = 1 Then strSQL = "UPDATE COM_Comments SET COM_Comments1 = '" & valud(1) & "' Where A_Code = " & APCode
    If nocoms = 2 Then strSQL = "UPDATE COM_Comments SET COM_Comments1 = '" & valud(1) & "', COM_Comments2 = '" & valud(2) & "' Where A_Code = " & APCode
    If nocoms = 3 Then strSQL = "UPDATE COM_Comments SET COM_Comments1 = '" & valud(1) & "', COM_Comments2 = '" & valud(2) & "', COM_Comments3 = '" & valud(3) & "' Where A_Code = " & APCode
    bOutput = CBA_DB_Connect.CBA_DB_CC_NonC("UPDATE", "ABI_QUERY", CBA_BSA & "LIVE DATABASES\ABI.accdb", CBA_MSAccess, strSQL, 120, , , False)  'Runs DB_Connection module to create connection to dtabase and run query
Else
    'insert if need to change
    CBA_DBtoQuery = 2
    If nocoms = 1 Then strSQL = "INSERT INTO COM_Comments (A_Code,COM_Comments1) VALUES(" & APCode & ", '" & valud(1) & "')"
    If nocoms = 2 Then strSQL = "INSERT INTO COM_Comments (A_Code,COM_Comments1,COM_Comments2) VALUES(" & APCode & ", '" & valud(1) & "', '" & valud(2) & "')"
    If nocoms = 3 Then strSQL = "INSERT INTO COM_Comments (A_Code,COM_Comments1,COM_Comments2,COM_Comments3) VALUES(" & APCode & ", '" & valud(1) & "', '" & valud(2) & "', '" & valud(3) & "')"
    bOutput = CBA_DB_Connect.CBA_DB_CC_NonC("UPDATE", "ABI_QUERY", CBA_BSA & "LIVE DATABASES\ABI.accdb", CBA_MSAccess, strSQL, 120, , , False)  'Runs DB_Connection module to create connection to dtabase and run query
End If

noupdaterqd:
CBA_DBtoQuery = 3
Unload Me


End Sub

Sub UserForm_Initialize()
    Me.StartUpPosition = 0
    Me.Top = Application.Top + (Application.Height / 2) - (Me.Height / 2)
    Me.Left = Application.Left + (Application.Width / 2) - (Me.Width / 2)
End Sub
Sub CommentsFormFormulate(ByVal CBA_Comment_APCode As Long)
    Dim strSQL As String, bOutput As Boolean
    
    APCode = CBA_Comment_APCode
    CBA_DBtoQuery = 1
    strSQL = "SELECT A_Code,COM_Comments1 & ' ' & COM_Comments2 & ' ' & COM_Comments3 FROM COM_Comments where A_Code = " & APCode
    bOutput = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", "ABI_QUERY", CBA_BSA & "LIVE DATABASES\ABI.accdb", CBA_MSAccess, strSQL, 120, , , False)  'Runs DB_Connection module to create connection to dtabase and run query
    If bOutput = True Then Me.box_Comment.Value = CBA_ABIarr(1, 0)
    CBA_DBtoQuery = 3


End Sub



