VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fCOM_Bench 
   ClientHeight    =   4965
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8790
   OleObjectBlob   =   "fCOM_Bench.frx":0000
   StartUpPosition =   1  'CenterOwner
   Tag             =   "fCOM_Bench"
End
Attribute VB_Name = "fCOM_Bench"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private pbIsActive As Boolean
Private psdMatch As Scripting.Dictionary

Private Sub tbx_PCode_Change() ' @RWGen Took out - no ListItem ref? 200129
'Dim dic As Scripting.Dictionary
'Dim v As Variant
'Dim T As cCOM_Bench
'Dim li As ListItem
'    If IsNumeric(tbx_PCode.Value) = False Or tbx_PCode = 0 Then MsgBox "Not a Valid ProductCode": Exit Sub
'    If sdMatch.Exists(CStr(tbx_PCode.Value)) = False Then Me.LV_Bench.ListItems.Clear: Exit Sub
'    Set dic = sdMatch(tbx_PCode.Value)
'
'
'    For Each v In dic
'        Set T = dic(v)
'        Set li = Me.LV_Bench.ListItems.Add(, , CStr(T.CompCode))
'        li.ListSubItems.Add , , CStr(T.PDesc)
'        li.ListSubItems.Add , , CStr(T.Competitor)
'        li.ListSubItems.Add , , CStr(T.MatchDescription)
'    Next
'    ''Me.LV_Bench.Refresh


End Sub

Private Sub UserForm_Initialize()
Dim COM As cCBA_Connect
Dim curProd As String
Dim T As cCOM_Bench
Dim dic As Scripting.Dictionary
Dim tcnt As Long
Dim compProds As Variant
Dim CompDic As Scripting.Dictionary
Dim QN As String
    
    Me.StartUpPosition = 0
    Me.Top = Application.Top + (Application.Height / 2) - (Me.Height / 2)
    Me.Left = Application.Left + (Application.Width / 2) - (Me.Width / 2)
    
    With Me.LV_Bench
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        With .ColumnHeaders
            .Clear
            .Add , , "Product Code", 70
            .Add , , "Description", 185
            .Add , , "Competitor", 80
            .Add , , "MatchType", 80
        End With
    End With
    
    
    
    Set COM = New cCBA_Connect
    QN = "ALL Matches"
    If COM.SetConnection("comrade") = False Then bIsActive = False: Exit Sub
    strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
    strSQL = strSQL & "DECLARE @sql_com as nvarchar(Max)"
    strSQL = strSQL & "DECLARE @listStr Varchar(Max)"
    strSQL = strSQL & "SELECT @listStr =  COALESCE(@listStr+' ' ,'select A_Code, A_CG, ') + Column_Name + ', ''' + Column_Name + ''' from tools.dbo.com_prodmap where ' + COLUMN_NAME + ' is not null '"
    strSQL = strSQL & "+ ' union select A_Code, A_CG, '"
    strSQL = strSQL & "FROM tools.INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'com_prodmap' and substring(Column_Name,1,1) <> 'A'"
    strSQL = strSQL & "select @sql_com =  substring(@listStr,1,len(@listStr)-26) + ' Order by 1,4'"
    strSQL = strSQL & "exec(@sql_com)"
    If COM.Query(strSQL, QN) = False Then bIsActive = False: Exit Sub
    compProds = CBA_COM_Runtime.CBA_COM_getCompProds
    Set CompDic = New Scripting.Dictionary
    For a = LBound(compProds, 2) To UBound(compProds, 2)
        If CompDic.Exists(CStr(compProds(1, a))) Then
            CompDic(CStr(compProds(1, a))) = CStr(compProds(2, a))
        Else
            CompDic.Add CStr(compProds(1, a)), CStr(compProds(2, a))
        End If
    Next
    Erase compProds
    Set sdMatch = New Scripting.Dictionary
    Do Until COM.op(QN).EOF
        If curProd = "1" Then Stop
        If COM.op(QN)(0) <> curProd Then
            If curProd <> "" Then sdMatch.Add CStr(curProd), dic
            curProd = CLng(COM.op(QN)(0))
            Set dic = New Scripting.Dictionary
        End If
        Set T = New cCOM_Bench
        T.CG = CLng(COM.op(QN)(1))
        T.CompCode = CStr(COM.op(QN)(2))
        T.MatchType = CStr(CCM_Mapping.CMM_getComp2Find(COM.op(QN)(3), T.CG))
        T.Competitor = CStr(CCM_Mapping.MatchType(T.MatchType).CompetitorLng)
        T.MatchDescription = CStr(CCM_Mapping.MatchType(T.MatchType).Description)
        T.PCode = CLng(COM.op(QN)(0))
        If CompDic.Exists(CStr(COM.op(QN)(2))) Then T.PDesc = CompDic(CStr(COM.op(QN)(2)))
        If T.MatchType <> "" Then dic.Add CStr(T.MatchType), T
        COM.op(QN).MoveNext
    Loop
    If curProd <> "" Then sdMatch.Add CStr(curProd), dic
    bIsActive = COM.IsLive

End Sub
Public Property Get bIsActive() As Boolean: bIsActive = pbIsActive: End Property
Private Property Let bIsActive(ByVal bNewValue As Boolean): pbIsActive = bNewValue: End Property
Private Property Get sdMatch() As Scripting.Dictionary: Set sdMatch = psdMatch: End Property
Private Property Set sdMatch(ByVal NewValue As Scripting.Dictionary): Set psdMatch = NewValue: End Property
