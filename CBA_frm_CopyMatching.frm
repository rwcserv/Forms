VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CBA_frm_CopyMatching 
   ClientHeight    =   8820
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14415
   OleObjectBlob   =   "CBA_frm_CopyMatching.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "CBA_frm_CopyMatching"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'#RW Added new mousewheel routines 190701
Private Sub box_CCS_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call mw_SetBoxHook(box_CCS)
End Sub

Private Sub Box_addProdtoList_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
Dim i As Long, SeltoCopy As Long, InfoBox, ColsArr, ChangesArr, LB
Dim strPrompt As String, strDesc As String, strSQL As String, ThisMatch, a As Long
Dim bOutput As Boolean, lRet As Long

If KeyCode = 13 Then
If Box_addProdtoList.Value <> "" Then
    If Box_addProdtoList.Value <> "" Then
        bOutput = CCM_SQLQueries.CBA_COM_MATCHGenPullSQL("FINDPRODDESC", , , , , Box_addProdtoList.Value)
        If bOutput = True Then
            strDesc = Box_addProdtoList.Value & "-" & CBA_CBISarr(0, 0)
        Else
            Box_addProdtoList.Value = ""
            MsgBox "You did not enter a valid product code", vbOKOnly
            Exit Sub
        End If
    End If

    SeltoCopy = 0
    For i = 0 To Me.box_CCS.ListCount - 1
        If Me.box_CCS.Selected(i) = True Then
            SeltoCopy = i
            Exit For
        End If
    Next i
    If SeltoCopy = 0 Then
        Box_addProdtoList.Value = ""
        MsgBox "Please select a product to copy matches from", vbOKOnly
        Exit Sub
    End If
'  '  Debug.Print Me.box_CCS.List(SeltoCopy, 0)
    strPrompt = "You wish to copy all matches from " & Me.box_CCS.List(SeltoCopy, 0) & "-" & Me.box_CCS.List(SeltoCopy, 1) & Chr(10) & Chr(10) & "To:" & Chr(10) & Chr(10) & strDesc
    lRet = MsgBox(strPrompt, vbYesNo)
    If lRet = 6 Then
        CBA_DBtoQuery = 3
        strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
        strSQL = strSQL & "DECLARE @ACodeto as nvarchar(6) = " & Box_addProdtoList.Value
        strSQL = strSQL & "DECLARE @ACodefrom as nvarchar(6) = " & Me.box_CCS.List(SeltoCopy, 0)
        strSQL = strSQL & "DECLARE @sql_com as nvarchar(Max)"
        strSQL = strSQL & "DECLARE @listStr Varchar(Max)"
        strSQL = strSQL & "SELECT @listStr = COALESCE(@listStr+',' ,' ') + Column_Name + '= (select ' + Column_Name + ' from  #Duplicate)'"
        strSQL = strSQL & "FROM tools.INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'com_prodmap'"
        strSQL = strSQL & "and substring(Column_Name,1,1) <> 'A'"
        strSQL = strSQL & "select @sql_com = 'select * into #Duplicate from tools.dbo.com_prodmap where A_Code = ' + @ACodefrom + ' update tools.dbo.com_prodmap SET ' + @listStr + ' where A_Code = ' + @ACodeto + ' drop table #Duplicate'"
        strSQL = strSQL & "exec(@sql_com)"
        bOutput = CBA_DB_Connect.CBA_DB_CC_NonC("UPDATE", "COMRADE", CBA_BasicFunctions.TranslateServerName("599DBL12", Date), "SQLOLEDB.1", strSQL, 60, "Tools")
        If bOutput = True Then
            Set InfoBox = CreateObject("WScript.Shell")
            ThisMatch = InfoBox.Popup("Matches Assigned", 1, "Matched Product")
        End If
        strSQL = "SELECT Column_Name FROM tools.INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'com_prodmap'"
        bOutput = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", "COMRADE", CBA_BasicFunctions.TranslateServerName("599DBL12", Date), "SQLOLEDB.1", strSQL, 60, "Tools")
        
        If bOutput = True Then
            ColsArr = CBA_COMarr
            Erase CBA_COMarr
        Else
            MsgBox "COMRADE ERROR 01" & Chr(10) & Chr(10) & "Contact " & g_Get_Dev_Sts("DevUsers"), vbOKOnly
            Exit Sub
        End If
        CBA_DBtoQuery = 3
        strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
        strSQL = strSQL & "DECLARE @ACodeto as nvarchar(6) = " & Box_addProdtoList.Value & Chr(10)
        strSQL = strSQL & "select * from  tools.dbo.com_prodmap where A_Code =  @ACodeto" & Chr(10)
        'Debug.Print strSQL
        bOutput = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", "COMRADE", CBA_BasicFunctions.TranslateServerName("599DBL12", Date), "SQLOLEDB.1", strSQL, 60, "Tools")
        If bOutput = True Then
            ChangesArr = CBA_COMarr
            Erase CBA_COMarr
            For a = LBound(ChangesArr, 1) To UBound(ChangesArr, 1)
                If ChangesArr(a, 0) <> "NULL" And a > 4 Then
                    CBA_DBtoQuery = 3
                    strSQL = "insert into tools.dbo.Com_MapChange (AldiUser, DateChanged, AldiProd, CompPCode,  CompType)" & Chr(10)
                    strSQL = strSQL & "Values('" & Application.UserName & "', getdate(), '" & Box_addProdtoList & "', '" & ChangesArr(a, 0) & "', '" & CCM_Mapping.CMM_getComp2Find(ColsArr(0, a), ChangesArr(2, 0)) & "')" & Chr(10)
                    'Debug.Print strSQL
                    bOutput = CBA_DB_Connect.CBA_DB_CC_NonC("UPDATE", "COMRADE", CBA_BasicFunctions.TranslateServerName("599DBL12", Date), "SQLOLEDB.1", strSQL, 60, "Tools")
                End If
            Next
        Else
            MsgBox "COMRADE ERROR 02" & Chr(10) & Chr(10) & "Contact " & g_Get_Dev_Sts("DevUsers"), vbOKOnly
            Exit Sub
        End If
    End If
    

    CCM_Runtime.CCM_updateMatches
    CBA_COM_Runtime.updateCopyDelete_Results
    Set LB = Me.box_CCS
    With LB
        .Clear
        .List = CBA_COM_Runtime.getCopyDelete_Results
    End With
    Box_addProdtoList.Value = ""


End If
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





