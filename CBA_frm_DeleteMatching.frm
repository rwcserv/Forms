VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CBA_frm_DeleteMatching 
   ClientHeight    =   8820
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13470
   OleObjectBlob   =   "CBA_frm_DeleteMatching.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "CBA_frm_DeleteMatching"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub but_Delete_Click()
Dim i As Long, lRet As Long, LB
Dim strUnass As String
    For i = 0 To Me.box_CCS.ListCount - 1
        If Me.box_CCS.Selected(i) = True Then
            If strUnass = "" Then
                strUnass = Me.box_CCS.List(i) & " - " & Me.box_CCS.List(i, 1) & Chr(10)
            Else
                strUnass = strUnass & Me.box_CCS.List(i) & " - " & Me.box_CCS.List(i, 1) & Chr(10)
            End If
        End If
    Next
    lRet = MsgBox("You have selected to delete all matches from: " & Chr(10) & Chr(10) & strUnass, vbYesNo)
    If lRet = 6 Then
    For i = 0 To Me.box_CCS.ListCount - 1
        If Me.box_CCS.Selected(i) = True Then
            CCM_DM_unassign CLng(Me.box_CCS.List(i, 0))
            CCM_Runtime.CCM_updateMatches
            CBA_COM_Runtime.updateCopyDelete_Results
            Set LB = Me.box_CCS
            With LB
                .Clear
                .List = CBA_COM_Runtime.getCopyDelete_Results
            End With
            Exit For
            
        End If
    Next i
    End If
End Sub

Sub UserForm_Initialize()
    Dim LB
     
    Me.StartUpPosition = 0
    Me.Top = Application.Top + (Application.Height / 2) - (Me.Height / 2)
    Me.Left = Application.Left + (Application.Width / 2) - (Me.Width / 2)

    Set LB = Me.box_CCS
    With LB
        .Clear
        .List = CBA_COM_Runtime.getCopyDelete_Results
    End With
End Sub
Private Sub CCM_DM_unassign(ByVal prod As Long)
'Dim InfoBox As Object
Dim bOutput As Boolean, strSQL As String, MRarr, a As Long, InfoBox, ThisMatch


If prod <> 0 Then
    CBA_DBtoQuery = 3
    strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
    strSQL = strSQL & "DECLARE @ACode as nvarchar(6) = " & prod & Chr(10)
    strSQL = strSQL & "DECLARE @sql_com as nvarchar(Max)" & Chr(10)
    strSQL = strSQL & "DECLARE @listStr Varchar(Max)" & Chr(10)
    strSQL = strSQL & "SELECT @listStr = COALESCE(@listStr+' union select ' ,' select ') +  'A_CG ,''' + Column_Name "
    strSQL = strSQL & "+ '''  from tools.dbo.com_prodmap where ' + COLUMN_NAME + ' is not null and A_Code = ' + @ACode" & Chr(10)
    strSQL = strSQL & "FROM tools.INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'com_prodmap'" & Chr(10)
    strSQL = strSQL & "and substring(Column_Name,1,1) <> 'A'" & Chr(10)
    strSQL = strSQL & "select @sql_com  = @listStr" & Chr(10)
    strSQL = strSQL & "exec(@sql_com)" & Chr(10)
    bOutput = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", "COMRADE", CBA_BasicFunctions.TranslateServerName("599DBL12", Date), "SQLOLEDB.1", strSQL, 60, "Tools")
    MRarr = CBA_COMarr
    strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
    strSQL = strSQL & "DECLARE @ACode as nvarchar(6) = " & prod & Chr(10)
    strSQL = strSQL & "DECLARE @sql_com as nvarchar(Max)" & Chr(10)
    strSQL = strSQL & "DECLARE @listStr Varchar(Max)" & Chr(10)
    strSQL = strSQL & "SELECT @listStr = COALESCE(@listStr+',' ,' ') + Column_Name + ' = NULL'" & Chr(10)
    strSQL = strSQL & "FROM tools.INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'com_prodmap'" & Chr(10)
    strSQL = strSQL & "and substring(Column_Name,1,1) <> 'A'" & Chr(10)
    strSQL = strSQL & "select @sql_com = 'update tools.dbo.com_prodmap SET ' + @listStr + ' where A_Code = ' + @ACode" & Chr(10)
    strSQL = strSQL & "exec(@sql_com)" & Chr(10)
    'Debug.Print strSQL
    bOutput = CBA_DB_Connect.CBA_DB_CC_NonC("UPDATE", "COMRADE", CBA_BasicFunctions.TranslateServerName("599DBL12", Date), "SQLOLEDB.1", strSQL, 60, "Tools")
    CBA_DBtoQuery = 4
    strSQL = ""
    For a = LBound(MRarr, 2) To UBound(MRarr, 2)
        strSQL = strSQL & "insert into tools.dbo.Com_MapChange (AldiUser, DateChanged, AldiProd, CompPCode,  CompType)" & Chr(10)
        strSQL = strSQL & "Values('" & Application.UserName & "', getdate(), '" & prod & "', 'Unassigned', '" & CCM_Mapping.CMM_getComp2Find(MRarr(1, a), MRarr(0, a)) & "')" & Chr(10)
    Next
'  '  Debug.Print strSQL
    bOutput = CBA_DB_Connect.CBA_DB_CC_NonC("UPDATE", "COMRADE", CBA_BasicFunctions.TranslateServerName("599DBL12", Date), "SQLOLEDB.1", strSQL, 60, "Tools")
    If bOutput = True Then
        Set InfoBox = CreateObject("WScript.Shell")
        ThisMatch = InfoBox.Popup("Match Unassigned", 1, "Matched Product")
    End If

End If




End Sub

