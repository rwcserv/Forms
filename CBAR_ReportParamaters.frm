VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CBAR_ReportParamaters 
   ClientHeight    =   9330
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4800
   OleObjectBlob   =   "CBAR_ReportParamaters.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "CBAR_ReportParamaters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private CBAR_Activation As Boolean
Private CBAR_DatesScraped() As Date
Private CBAR_BuyerEmailerWKS As Scripting.Dictionary

Private Sub but_GenerateReport_Click()
Dim tR As CBAR_Report
Dim GraphsToMake As Collection
Dim a As Long, lRet As Long, RCell, strProds As String, strPProds As String, strSQL As String
Dim CBA_CBISCN As ADODB.Connection, CBA_CBISRS As ADODB.Recordset, Dto, DFrom
Dim weeks As Long, dates As Long, d As Long, yn As Long
Dim colProds As Collection, colPProds As Collection
Dim pro As Variant
tR = CBAR_Runtime.getActiveReport
Select Case tR.ReportNo
'    If tR.ReportNo = 1 Or tR.ReportNo = 6 Then
    Case 1, 4, 6, 7, 8, 9, 10
        If tR.State <> "" And tR.Competitor <> "" And (tR.BD <> "" Or tR.GBD <> "" Or tR.CG <> 0 Or Not tR.AldiProds Is Nothing) Then
            Unload Me
            Select Case tR.ReportNo
                Case 1: CBAR_ActiveMatch.AMR
                Case 4: CBAR_REPORTMMR.CBAR_MMRRuntime
                Case 6: CBAR_NoLongerPromo.CBAR_NoLongerOnPromo
                Case 7: CBAR_StartPromo.CBAR_OnPromoReport
                Case 8: CBAR_PermPriceChange.CBAR_PermPriceChange
                Case 9: CBAR_MatchedNotOnWeb.CBAR_MatchedNotOnWeb
                Case 10: CBAR_AldiNoMatch.CBAR_AldiNoMatch
            End Select
        Else
            If tR.ReportNo = 1 Or tR.ReportNo = 4 Then
                
                yn = MsgBox("Run for ALL Matches?" & Chr(10) & Chr(10) & "This will take a few minutes..", vbYesNo)
                If yn = 6 Then
                    Unload Me
                    If tR.ReportNo = 1 Then CBAR_ActiveMatch.AMR
                    If tR.ReportNo = 4 Then CBAR_REPORTMMR.CBAR_MMRRuntime
                End If
            Else
                MsgBox "Alongside State and Competitors, you must choose one of the following variables to run the Active Match Report: GBD, BD, CG, CG/SCG, Product Codes"
            End If
        End If

    Case 2
        If tR.State <> "" And tR.Competitor <> "" And tR.DateFrom <> "12:00:00 AM" And tR.DateTo <> "12:00:00 AM" And (tR.BD <> "" Or tR.GBD <> "" Or tR.CG <> 0 Or Not tR.AldiProds Is Nothing) Then
            If tR.AldiProds Is Nothing Then
                lRet = MsgBox("It is advised that this report is run by productcode, as running by other parameters alone can create 100s of graphs." & Chr(10) & Chr(10) & "Would you like to refine your query to add productcodes?" & Chr(10) & Chr(10) & "(n.b. you can paste from an excel list into the Aldi Product Number Text Box)", vbYesNo)
                If lRet = 6 Then Exit Sub
            End If
            Unload Me
            CBAR_PPHistory.CBAR_PPHistoryRuntime
        Else
            MsgBox "Alongside State,Competitors,DateFrom and DateTo, you must choose one of the following variables to run the Price and Promotion History Report: GBD, BD, CG, CG/SCG, Product Codes"
        End If
    
    
    Case 3
        If tR.State <> "" And tR.Competitor <> "" And (tR.BD <> "" Or tR.GBD <> "" Or tR.CG <> 0 Or Not tR.AldiProds Is Nothing) Then
            Unload Me
            CBAR_REPORTStateVar.StateVariationReport
        Else
            MsgBox "Alongside State and Competitors, you must choose one of the following variables to run the State Variation Report: GBD, BD, CG, CG/SCG, Product Codes"
        End If
   
    Case 5
        If tR.State <> "" And tR.Competitor <> "" And tR.DateFrom <> "12:00:00 AM" And tR.DateTo <> "12:00:00 AM" And (tR.BD <> "" Or tR.GBD <> "" Or tR.CG <> 0) Then
            Unload Me
            CBAR_ReportPromoAnalysis.CBAR_PAREport
        Else
            MsgBox "Alongside State,Competitors,DateFrom and DateTo, you must choose one of the following variables to run the Promotional Activity Report: GBD, BD, CG, CG/SCG, Product Codes"
        End If
        
    Case 11                                                 ' 11, 12
        If tR.State <> "" And tR.DateFrom <> 0 And tR.DateTo <> 0 Then
            Set GraphsToMake = New Collection
            For a = 0 To Me.lbx_ReportOptions.ListCount - 1
                If Me.lbx_ReportOptions.Selected(a) = True Then
                    If a = 0 Then GraphsToMake.Add "All"
                    If a = 1 Then GraphsToMake.Add "ColesCOS"
                    If a = 2 Then GraphsToMake.Add "WWCOS"
                    If a = 3 Then GraphsToMake.Add "DMCOS"
                    If a = 4 Then GraphsToMake.Add "ColesSB"
                    If a = 5 Then GraphsToMake.Add "ColesVal"
                    If a = 6 Then GraphsToMake.Add "ColesPL"
                    If a = 7 Then GraphsToMake.Add "ColesPhantoms"
                    If a = 8 Then GraphsToMake.Add "WWHB"
                    If a = 9 Then GraphsToMake.Add "WWSelect"
                    If a = 10 Then GraphsToMake.Add "WWPhantoms"
                    If a = 11 Then GraphsToMake.Add "DM1"
                    If a = 12 Then GraphsToMake.Add "DM2"
                    If a = 13 Then GraphsToMake.Add "DMPhantoms"
                    If a = 14 Then GraphsToMake.Add "Branded"
                    If a = 15 Then GraphsToMake.Add "Phantoms"
                    If a = 16 Then GraphsToMake.Add "DMQ"
                    If a = 17 Then GraphsToMake.Add "FC1"
                    If a = 18 Then GraphsToMake.Add "FC2"
                    If a = 19 Then GraphsToMake.Add "FCQ"
                End If
            Next
            If GraphsToMake.Count = 0 Then
                MsgBox "No Charts were requested. Please select at least one chart", vbOKOnly
                GoTo Exit_Routine               '''Unload Me
            End If
        
            Unload Me
''            If tR.ReportNo = 11 Then
''                CBAR_REPORTWBABD.WBA tR.State, tR.datefrom, tR.dateto, "weeks", GraphsToMake, tR.CG, tR.Matchhistory
''            ElseIf tR.ReportNo = 12 Then
                CBAR_REPORTWBA.WBA tR.State, tR.DateFrom, tR.DateTo, "months", GraphsToMake, tR.CG, tR.Matchhistory
''            End If
        Else
            MsgBox "This report requires 'DateFrom', 'DateTo', 'State', and 'Match History' paramaters to be popultated", vbOKOnly
            'Unload Me
        End If
        
    Case 12                     '14
            
        If tR.State <> "" And tR.DateFrom <> 0 And tR.DateTo <> 0 And tR.Competitor <> "" Then
            
            Set CBA_CBISCN = New ADODB.Connection
            With CBA_CBISCN
                .ConnectionTimeout = 300
                .CommandTimeout = 300
                .Open "Provider= SQLNCLI10; DATA SOURCE= 599DBL01; ;INTEGRATED SECURITY=sspi;"
            End With
            
            Set colProds = New Collection: Set colPProds = New Collection
            For Each RCell In CBAR_Data.Columns(29).Cells
                If RCell.Value = "" And RCell.Offset(1, 0).Value = "" And RCell.Offset(2, 0).Value = "" Then Exit For
                colProds.Add RCell.Value
            Next
            For Each RCell In CBAR_Data.Columns(30).Cells
                If RCell.Value = "" And RCell.Offset(1, 0).Value = "" And RCell.Offset(2, 0).Value = "" Then Exit For
                colPProds.Add RCell.Value
            Next
            
            If CBAR_Runtime.CBAR_CheckTop150List(colProds, colPProds) = False Then Unload Me: Exit Sub
                    
                    
            For Each pro In colProds
                If strProds = "" Then strProds = pro Else strProds = strProds & ", " & pro
            Next
             For Each pro In colPProds
                If strPProds = "" Then strPProds = pro Else strPProds = strPProds & ", " & pro
            Next
            
            strSQL = "select distinct  isnull(con_productcode, productcode) from cbis599p.dbo.product where productcode in (" & strProds & ")" & Chr(10)
            strSQL = strSQL & "union select isnull(con_productcode, productcode) from cbis599p.dbo.product where productcode in (" & strPProds & ") or con_productcode in (" & strPProds & ")"
            'Debug.Print strSQL

            Set CBA_CBISRS = New ADODB.Recordset
            CBA_CBISRS.Open strSQL, CBA_CBISCN
            If CBA_CBISRS.EOF Then
                ReDim CBA_CBISarr(0, 0)
                CBA_CBISarr(0, 0) = 0
            Else
                CBA_CBISarr = CBA_CBISRS.GetRows()
            End If
            If CBA_CBISarr(0, 0) <> 0 Then
                For a = LBound(CBA_CBISarr, 2) To UBound(CBA_CBISarr, 2)
                    If a = LBound(CBA_CBISarr, 2) Then strProds = CBA_CBISarr(0, a) Else strProds = strProds & " , " & CBA_CBISarr(0, a)
                Next
                Dto = DateAdd("d", 3 - WeekDay(tR.DateTo, vbMonday), tR.DateTo)
                DFrom = DateAdd("d", 2 + WeekDay(tR.DateFrom, vbMonday), tR.DateFrom)
                Unload Me
                CBAR_REPORTTop150.Top150Run DFrom, Dto, strProds, tR.Matchhistory
            Else
                MsgBox "The Top 150 Product List was not compiled properly. Please check with " & g_Get_Dev_Sts("DevUsers"), vbOKOnly
                Exit Sub
            End If
            
            
            
        Else
            MsgBox "This report requires 'DateFrom', 'DateTo', 'Competitor', 'State', and 'Match History' paramaters to be popultated", vbOKOnly
            'Unload Me
        End If

    Case 13             '15
        Me.Hide
        CBAR_Runtime.setActiveReportParamater "State", "National"
        CBAR_Runtime.setActiveReportParamater "competitor", "All Competitors"
        tR = CBAR_Runtime.getActiveReport
        If tR.BD = "Produce" Or tR.CG = 58 Then Else Dto = CBA_COM_Runtime.CBA_getWedDate
        weeks = 4
        DFrom = DateAdd("WW", -weeks + 1, Dto)
        dates = 0
        ReDim CBAR_DatesScraped(1 To 1)
        For d = 0 To DateDiff("D", DFrom, Dto)
            If WeekDay(DFrom + d, vbWednesday) = 1 Then
                dates = dates + 1
                ReDim Preserve CBAR_DatesScraped(1 To dates)
                CBAR_DatesScraped(dates) = DFrom + d
            End If
        Next
        If CBA_COM_SetupMatchArray.CBA_SetupMatchArray(False, DFrom, Dto, , , , True) = True Then
            Set CBAR_BuyerEmailerWKS = New Scripting.Dictionary
            CBAR_NoLongerPromo.CBAR_NoLongerOnPromo True
            CBAR_StartPromo.CBAR_OnPromoReport True
            CBAR_PermPriceChange.CBAR_PermPriceChange True
            CBAR_MatchedNotOnWeb.CBAR_MatchedNotOnWeb True
            CBAR_AldiNoMatch.CBAR_AldiNoMatch True
            
            
            
            
        Else
            MsgBox "error in running match objects"
        End If
        

        CBA_COM_Emailer.EmailBDs CBAR_BuyerEmailerWKS
        
        
        
        
    End Select
Exit_Routine:
    On Error Resume Next
    Set CBA_CBISCN = Nothing
    Set CBA_CBISRS = Nothing
    
        
        
        
    
    'Unload Me
End Sub
Function getEmailerScrapedDatesArray() As Date()
    getEmailerScrapedDatesArray = CBAR_DatesScraped
End Function
Function setBuyerEmailerWorksheet(ByVal QueryName As String, ByRef wks As Worksheet)
    If CBAR_BuyerEmailerWKS Is Nothing Then Set CBAR_BuyerEmailerWKS = New Scripting.Dictionary
    If CBAR_BuyerEmailerWKS.Exists(QueryName) = True Then CBAR_BuyerEmailerWKS.Remove QueryName
    CBAR_BuyerEmailerWKS.Add QueryName, wks
End Function
Private Sub cbx_MatchHistory_Change()
Dim win As Boolean
    If CBAR_Activation = True Then Exit Sub
    If cbx_MatchHistory <> "" Then win = CBAR_Runtime.setActiveReportParamater("Matchhistory", cbx_MatchHistory.Value)
    If win = False Then MsgBox "An Invalid entry has been made", vbOKOnly
End Sub
Private Sub cbx_BD_Change()
Dim win As Boolean
    If CBAR_Activation = True Then Exit Sub
    win = CBAR_Runtime.setActiveReportParamater("BD", cbx_BD.Value)
    'If win = False Then MsgBox "An Invalid entry has been made", vbOKOnly
End Sub
Private Sub cbx_BD_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call mw_SetBoxHook(cbx_BD)
End Sub

Private Sub cbx_CGs_Change()
Dim win As Boolean, col, scg
    If CBAR_Activation = True Then Exit Sub
    If cbx_CGs.Value = "" Then
        Me.cbx_SCGs.Visible = False: Me.lbl_SCG.Visible = False
        Exit Sub
    End If
    win = CBAR_Runtime.setActiveReportParamater("CG", , Mid(cbx_CGs.Value, 1, 2))
    If win = False Then
        MsgBox "An Invalid entry has been made", vbOKOnly
    Else
        Me.cbx_SCGs.Visible = True: Me.lbl_SCG.Visible = True
        Me.cbx_SCGs.Clear
        Set col = CBA_COM_Runtime.getSCGs
        For Each scg In col
            If Mid(scg, 1, InStr(1, scg, "-") - 1) = Mid(Me.cbx_CGs.Value, 1, InStr(1, Me.cbx_CGs.Value, "-") - 1) Then Me.cbx_SCGs.AddItem scg
        Next
    End If
End Sub
Private Sub cbx_CGs_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call mw_SetBoxHook(cbx_CGs)
End Sub

Private Sub cbx_Comp_Change()
Dim win As Boolean
    If CBAR_Activation = True Then Exit Sub
    win = CBAR_Runtime.setActiveReportParamater("Competitor", cbx_Comp.Value)
    If win = False Then MsgBox "An Invalid entry has been made", vbOKOnly
End Sub
Private Sub cbx_Comp_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call mw_SetBoxHook(cbx_Comp)
End Sub

Private Sub cbx_GBD_Change()
Dim win As Boolean
    If CBAR_Activation = True Then Exit Sub
    win = CBAR_Runtime.setActiveReportParamater("GBD", cbx_GBD.Value)
    'If win = False Then MsgBox "An Invalid entry has been made", vbOKOnly
End Sub
Private Sub cbx_GBD_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call mw_SetBoxHook(cbx_GBD)
End Sub

Private Sub cbx_SCGs_Change()
Dim win As Boolean
    If CBAR_Activation = True Then Exit Sub
    If IsNumeric(Mid(Me.cbx_SCGs.Value, InStr(1, cbx_SCGs.Value, "-") + 1, 2)) Then win = CBAR_Runtime.setActiveReportParamater("SCG", , Mid(cbx_SCGs.Value, InStr(1, cbx_SCGs.Value, "-") + 1, 2))
    'If win = False Then MsgBox "An Invalid entry has been made", vbOKOnly
End Sub
Private Sub cbx_SCGs_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call mw_SetBoxHook(cbx_SCGs)
End Sub

Private Sub cbx_State_Change()
Dim win As Boolean
    If CBAR_Activation = True Then Exit Sub
    win = CBAR_Runtime.setActiveReportParamater("State", cbx_State.Value)
    If win = False Then MsgBox "An Invalid entry has been made", vbOKOnly
End Sub
Private Sub tbx_AldiProds_Afterupdate()
Dim win As Boolean
    'If CBAR_Activation = True Or (KeyCode <> 13 And KeyCode <> 9) Then Exit Sub
    
    
    win = CBAR_Runtime.setActiveReportParamater("AldiProds", tbx_AldiProds.Value)
    'If win = False Then MsgBox "An Invalid entry has been made", vbOKOnly
End Sub
Private Sub tbx_Datefrom_Afterupdate()
Dim win As Boolean
Dim dtetopass As Date
'If CBAR_Activation = True Or (KeyCode <> 13 And KeyCode <> 9) Then Exit Sub
If isDate(tbx_Datefrom.Value) = True Then
    dtetopass = tbx_Datefrom.Value
ElseIf tbx_Datefrom.Value = "" Then
    dtetopass = 0
Else
    tbx_Datefrom.Value = ""
    MsgBox "An Invalid entry has been made", vbOKOnly
    Exit Sub
End If
win = CBAR_Runtime.setActiveReportParamater("datefrom", , , dtetopass)
If win = False And dtetopass <> "12:00:00 AM" Then MsgBox "An Invalid entry has been made", vbOKOnly
End Sub
Private Sub tbx_Dateto_Afterupdate()
Dim win As Boolean
Dim dtetopass As Date
'If CBAR_Activation = True Or (KeyCode <> 13 And KeyCode <> 9) Then Exit Sub
If isDate(tbx_Dateto.Value) = True Then
    dtetopass = tbx_Dateto.Value
ElseIf tbx_Dateto.Value = "" Then
    dtetopass = 0
Else
    tbx_Dateto.Value = ""
    MsgBox "An Invalid entry has been made", vbOKOnly
    Exit Sub
End If
win = CBAR_Runtime.setActiveReportParamater("Dateto", , , dtetopass)
If win = False And dtetopass <> "12:00:00 AM" Then MsgBox "An Invalid entry has been made", vbOKOnly
End Sub
Private Sub setupbox(ByRef obj As Object)
Dim TempArr() As Variant
Dim col As New Collection, th
Dim a As Long, LB

    ReDim TempArr(1 To 1)
    
    Select Case obj.Name
    
    Case Me.cbx_CGs.Name
        Set col = CBA_COM_Runtime.getCGs
    Case Me.cbx_BD.Name
        Set col = CBA_COM_Runtime.getBuyers
    Case Me.cbx_GBD.Name
        Set col = CBAR_Runtime.CBAR_getGBSs
    Case Me.cbx_State.Name
        ReDim TempArr(1 To 6)
        TempArr(1) = "NSW": TempArr(2) = "VIC": TempArr(3) = "QLD": TempArr(4) = "SA": TempArr(5) = "WA": TempArr(6) = "National"
        Set LB = Me.cbx_State
        With LB
            .Clear
            .List = TempArr
            .Value = "National"
            CBAR_Activation = False: CBAR_Runtime.setActiveReportParamater "State", "National": CBAR_Activation = True
        End With
        Exit Sub
    Case Me.cbx_Comp.Name
        ReDim TempArr(1 To 5)
        TempArr(1) = "Woolworths": TempArr(2) = "Coles": TempArr(3) = "Dan Murphys": TempArr(4) = "First Choice": TempArr(5) = "All Competitors"
        Set LB = Me.cbx_Comp
        With LB
            .Clear
            .List = TempArr
            .Value = "All Competitors"
            CBAR_Activation = False: CBAR_Runtime.setActiveReportParamater "competitor", "All Competitors": CBAR_Activation = True
        End With
        Exit Sub
    Case Me.cbx_MatchHistory.Name
        ReDim TempArr(1 To 2)
        TempArr(1) = "Yes": TempArr(2) = "No"
        Set LB = Me.cbx_MatchHistory
        With LB
            .Clear
            .List = TempArr
        End With
        Exit Sub
    End Select
    
    a = 1
    For Each th In col
        a = a + 1
        ReDim Preserve TempArr(1 To a)
        TempArr(a) = th
    Next
    Set LB = obj
    With LB
        .Clear
        .List = TempArr
    End With



End Sub
Sub UserForm_Initialize()
    Dim tR As CBAR_Report
    CBAR_Activation = True
    Me.StartUpPosition = 0
    Me.Top = Application.Top + (Application.Height / 2) - (Me.Height / 2)
    Me.Left = Application.Left + (Application.Width / 2) - (Me.Width / 2)

    tR = CBAR_Runtime.getActiveReport
    CBAR_Runtime.createActiveReport tR.ReportName
    tR = CBAR_Runtime.getActiveReport
    EnableDisableFeatures tR.ReportNo

    setupbox Me.cbx_CGs
    setupbox Me.cbx_BD
    setupbox Me.cbx_GBD
    setupbox Me.cbx_State
    setupbox Me.cbx_Comp
    setupbox Me.cbx_MatchHistory
    CBAR_Activation = False

End Sub
Private Function EnableDisableFeatures(ByVal Repno As Long)
    Dim a As Long
    Select Case Repno
    
        Case 1, 3, 4, 6, 7, 8, 9, 10
            Me.tbx_Datefrom.Visible = False: Me.lbl_DateFrom.Visible = False
            Me.tbx_Dateto.Visible = False: Me.lbl_DateTo.Visible = False
            Me.tbx_AldiProds.Visible = True: Me.lbl_AldiProds.Visible = True
            Me.cbx_BD.Visible = True: Me.lbl_BD.Visible = True
            Me.cbx_CGs.Visible = True: Me.lbl_CG.Visible = True
            Me.cbx_Comp.Visible = True: Me.lbl_Comp.Visible = True
            Me.cbx_GBD.Visible = True: Me.lbl_GBD.Visible = True
            Me.cbx_MatchHistory.Visible = False: Me.lbl_MatchHistory.Visible = False
            If Me.cbx_CGs.Value <> "" Then
                Me.cbx_SCGs.Visible = True: Me.lbl_SCG.Visible = True
            Else
                Me.cbx_SCGs.Visible = False: Me.lbl_SCG.Visible = False
            End If
            Me.cbx_State.Visible = True: Me.lbl_State.Visible = True
            Me.lbx_ReportOptions.Visible = False
        Case 2, 5
            Me.tbx_Datefrom.Visible = True: Me.lbl_DateFrom.Visible = True
            Me.tbx_Dateto.Visible = True: Me.lbl_DateTo.Visible = True
            Me.tbx_AldiProds.Visible = True: Me.lbl_AldiProds.Visible = True
            Me.cbx_BD.Visible = True: Me.lbl_BD.Visible = True
            Me.cbx_CGs.Visible = True: Me.lbl_CG.Visible = True
            Me.cbx_Comp.Visible = True: Me.lbl_Comp.Visible = True
            Me.cbx_GBD.Visible = True: Me.lbl_GBD.Visible = True
            Me.cbx_MatchHistory.Visible = False: Me.lbl_MatchHistory.Visible = False
            If Me.cbx_CGs.Value <> "" Then
                Me.cbx_SCGs.Visible = True: Me.lbl_SCG.Visible = True
            Else
                Me.cbx_SCGs.Visible = False: Me.lbl_SCG.Visible = False
            End If
            Me.cbx_State.Visible = True: Me.lbl_State.Visible = True
            Me.lbx_ReportOptions.Visible = False
        Case 7, 12
            Me.cbx_MatchHistory.Value = "Yes"
            With Me.lbx_ReportOptions
            .Clear
            For a = 1 To 16
                .AddItem CBAR_Data.Cells(a, 25).Value
            Next
            .Visible = True
            End With
        Case 11
            Me.tbx_AldiProds.Visible = False: Me.lbl_AldiProds.Visible = False
            Me.cbx_MatchHistory.Value = "Yes"
            With Me.lbx_ReportOptions
            .Clear
            For a = 1 To 16
                .AddItem CBAR_Data.Cells(a, 25).Value
            Next
            .Visible = True
            End With
        Case 13
            Me.tbx_Datefrom.Visible = False: Me.lbl_DateFrom.Visible = False
            Me.tbx_Dateto.Visible = False: Me.lbl_DateTo.Visible = False
            Me.tbx_AldiProds.Visible = False: Me.lbl_AldiProds.Visible = False
            Me.cbx_BD.Visible = False: Me.lbl_BD.Visible = False
            Me.cbx_CGs.Visible = False: Me.lbl_CG.Visible = False
            Me.cbx_Comp.Visible = False: Me.lbl_Comp.Visible = False
            Me.cbx_GBD.Visible = False: Me.lbl_GBD.Visible = False
            Me.cbx_MatchHistory.Visible = False: Me.lbl_MatchHistory.Visible = False
            Me.cbx_SCGs.Visible = False: Me.lbl_SCG.Visible = False
            Me.cbx_State.Visible = False: Me.lbl_State.Visible = False
            Me.lbx_ReportOptions.Visible = False
        
        Case Else
            Me.tbx_Datefrom.Visible = True
            Me.tbx_Dateto.Visible = True
            Me.tbx_AldiProds.Visible = True
            Me.cbx_BD.Visible = True
            Me.cbx_CGs.Visible = True
            Me.cbx_Comp.Visible = True
            Me.cbx_GBD.Visible = True
            Me.cbx_MatchHistory.Visible = True
            If Me.cbx_CGs.Value <> "" Then Me.cbx_SCGs = True Else Me.cbx_SCGs = False
            Me.cbx_State.Visible = True
            Me.lbx_ReportOptions.Visible = False
    
    End Select

End Function
Private Sub UserForm_Terminate()
    Unload Me
End Sub
