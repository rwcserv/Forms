VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CBA_BTF_frm_SCGNo 
   Caption         =   "Forecasts By Sub Commodity Group"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   26040
   OleObjectBlob   =   "CBA_BTF_frm_SCGNo.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "CBA_BTF_frm_SCGNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit       ' SCG @CBA_BTF      Changed 20/05/2019

' RW@ 190102 Date Issues Hard-Coded for a while

Private pbSetupIP As Boolean
Private pCGDta(1 To 4) As CBA_BTF_SCGNonProd
Private pcurSCGNo As Long
Private psDataUsed As String
Private pWksht As Worksheet
Private plLastRow As Long, bPassOK As Boolean
Private pcurSCGDesc As String
''Private pWkBk As Workbook
Private pdtDate As Date, pdteDate As Date
Private pcurCGNo As Long
Private pbytCurPClass As Byte
Private pbXLevel As Boolean                                 ' Indication that the item is not the latest item
Private pbHasBeenUpdated As Boolean
Private pb_NoDataReturned As Boolean
Private plYearIdx As Long

'#RW Added new mousewheel routines 190701
Private Sub lbx_SCGNo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call mw_SetBoxHook(lbx_SCGNo)
End Sub

' This form will retrieve prior values from the CBis database by Sub-Commodity Group, which can be used to forecast subsequent costs/sales

Private Sub but_6Less_Click()
    If p_TestEntry = False Then Exit Sub
    Call SetMonthStart(-1) ' Set the labels and the month positioning -1
End Sub

Private Sub but_6More_Click()
    If p_TestEntry = False Then Exit Sub
    Call SetMonthStart(1) ' Set the labels and the month positioning +1
End Sub

Private Sub pAll_Changes(sFieldToBeChanged As String, Optional ByVal dblVal As Single = -1.1111, Optional ByVal bSetupIP As Boolean = False)
    ' Handle and format all changes onto the screen - it may even format it when it is initiated ...
    Dim lLblNo As Long, a As Long, bDspMsg As Boolean, bDspMsgb As Boolean, bBlankPull As Boolean
    
    Static bInUse As Boolean
    ' If fields are not being initialised...
    If bSetupIP = False Then
        If p_TestEntry = False Then Exit Sub
    End If
    If bInUse = True Then ' If this procedure is iteratively called (i.e. a change to the field in this proc calls it again), hop out
        Exit Sub
    End If
    bInUse = True: bBlankPull = False
    ' Get the element no - now called the label no
    lLblNo = g_GetNo(sFieldToBeChanged)
    Select Case sFieldToBeChanged
        
        Case "A2A_Uplift_AfterUpdate"
            If g_IsNumeric(Me.A2A_Uplift.Value, False) = True Then
                dblVal = g_UnFmt(Me.A2A_Uplift.Value, "sng")
                If dblVal <> 0 Then
                    For a = 1 To 12
                        GoSub GSUplift
                    Next
                    CBA_BTF_Runtime.UpdateTotals pWksht, pCGDta, pbytCurPClass, plYearIdx, True
                End If
            Else
                If Trim(Me.A2A_Uplift) > "" Then
                    bDspMsg = True
                Else
                    bBlankPull = True
                End If
            End If
            Me.A2A_Uplift = ""
        
        Case "A2A_RCVMargin_AfterUpdate"
            If g_IsNumeric(Me.A2A_RCVMargin.Value, False) = True Then
                dblVal = g_UnFmt(Me.A2A_RCVMargin.Value, "sng")
                If dblVal <> 0 Then
                    For a = 1 To 12
                        GoSub GSMargin
                    Next
                    CBA_BTF_Runtime.UpdateTotals pWksht, pCGDta, pbytCurPClass, plYearIdx, True
                End If
            Else
                If Trim(Me.A2A_RCVMargin.Value) > "" Then
                    bDspMsg = True
                Else
                    bBlankPull = True
                End If
            End If
            Me.A2A_RCVMargin.Value = ""
        
        Case "A2A_FPOSRET_AfterUpdate"
            If g_IsNumeric(Me.A2A_FPOSRET.Value, False) = True Then
                dblVal = g_UnFmt(Me.A2A_FPOSRET.Value, "sng")
                If dblVal <> 0 Then
                    For a = 1 To 12
                        GoSub GSPOSRET
                    Next
                    CBA_BTF_Runtime.UpdateTotals pWksht, pCGDta, pbytCurPClass, plYearIdx, True
                End If
            Else
                If Trim(Me.A2A_FPOSRET.Value) > "" Then
                    bDspMsg = True
                Else
                    bBlankPull = True
                End If
            End If
            Me.A2A_FPOSRET.Value = ""
            
        Case "M" & lLblNo & "_Uplift_AfterUpdate"
            If g_IsNumeric(Me("M" & lLblNo & "_Uplift").Value, False) = True Then
                dblVal = g_UnFmt(Me("M" & lLblNo & "_Uplift").Value, "sng")
                a = lLblNo
                GoSub GSUplift
                If Not bSetupIP Then CBA_BTF_Runtime.UpdateTotals pWksht, pCGDta, pbytCurPClass, plYearIdx, True
            Else
                bDspMsg = True
            End If
        
        Case "M" & lLblNo & "_RCVMargin_AfterUpdate"
            If g_IsNumeric(Me("M" & lLblNo & "_RCVMargin").Value, False) = True Then
                dblVal = g_UnFmt(Me("M" & lLblNo & "_RCVMargin").Value, "sng")
                a = lLblNo
                GoSub GSMargin
                If Not bSetupIP Then CBA_BTF_Runtime.UpdateTotals pWksht, pCGDta, pbytCurPClass, plYearIdx, True
            Else
                bDspMsg = True
            End If
        
        Case "M" & lLblNo & "_FPOSRET_AfterUpdate"
            If g_IsNumeric(Me("M" & lLblNo & "_FPOSRET").Value, False) = True Then
                dblVal = g_UnFmt(Me("M" & lLblNo & "_FPOSRET").Value, "sng")
                a = lLblNo
                GoSub GSPOSRET
                If Not bSetupIP Then CBA_BTF_Runtime.UpdateTotals pWksht, pCGDta, pbytCurPClass, plYearIdx, True
            Else
                bDspMsg = True
            End If
                
    End Select
    ' If a warning msg has been designated
    If bDspMsgb = True Then
        MsgBox "Some values were not affected as there is no current or prior data - to fill these enter into their individual fields", vbOKOnly
        pbHasBeenUpdated = True ' records have been updated
    ' If an error msg has been designated
    ElseIf bDspMsg = True Then
        MsgBox "An invalid entry has been made", vbOKOnly
    ' If fields are not being initialised...
    ElseIf bSetupIP = False Then
        If bBlankPull = False Then pbHasBeenUpdated = True ' records have been updated
    End If
    ' Set the iterative bool off
    bInUse = False
    Exit Sub
    
GSUplift:
    If bSetupIP = False Then pCGDta(pbytCurPClass).UpdateValue "Uplift", dblVal / 100, a, plYearIdx
    Me.Controls("M" & a & "_FPOSRET") = Format(pCGDta(pbytCurPClass).ForeRetail((plYearIdx * 12) + a), "#,0")
    Me.Controls("M" & a & "_Uplift").Value = Format(pCGDta(pbytCurPClass).Uplift((plYearIdx * 12) + a) * 100, "0.00") & " %"

    Return
GSPOSRET:
    If bSetupIP = False Then pCGDta(pbytCurPClass).UpdateValue "ForeRetail", dblVal, a, plYearIdx
    Me.Controls("M" & a & "_FPOSRET") = Format(pCGDta(pbytCurPClass).ForeRetail((plYearIdx * 12) + a), "#,0")
    Me.Controls("M" & a & "_Uplift").Value = Format(pCGDta(pbytCurPClass).Uplift((plYearIdx * 12) + a) * 100, "0.00") & " %"
    Return
GSMargin:
    If bSetupIP = False Then pCGDta(pbytCurPClass).UpdateValue "ForeRCVMargin", dblVal / 100, a, plYearIdx
    Me("M" & a & "_RCVMargin").Value = Format(pCGDta(pbytCurPClass).ForeRCVMargin((plYearIdx * 12) + a) * 100, "0.00") & " %"
    Return
End Sub


Private Sub btn_Apply_Click()
    ' This routine will format and write the forecasts to the 'Forecast DB' database table SCGData
    ' This data will be used in the several forecasting reports

    Dim CN As ADODB.Connection, RS As ADODB.Recordset
    Dim lLblNo As Long, bReForecast As Boolean, strSQL As String, lCG As Long, lSCG As Long, lPCls As Long, sDate As String
    Dim bFCChg As Boolean, sReturn As String
    Static dtDateLastWritten As Date, lCGNoLastWritten As Long, lSCGNoLastWritten As Long, lstcPCls As Long
    ' Get the date of the cutoff to see if will apply
    sDate = CDate("01-" & Me("M1_LBL").Caption)
    sReturn = CBA_BT_getCutOffDate(CDate(sDate))
    If sReturn = "NoSave" Then
        MsgBox "No save of these forecasts are able to be performed as the dates selected are either too early or too late", vbOKOnly
        Exit Sub
    ElseIf sReturn = "Format" Then
        bReForecast = False
    ElseIf sReturn = "ReFormat" Then
        bReForecast = True
    Else
        MsgBox "No save of these forecasts are able to be performed at this time as no cutoff dates for the year can be found", vbOKOnly
        Exit Sub
    End If

    bFCChg = False
    If p_TestEntry = False Then Exit Sub
    If BTF_LastFCastTest(pcurCGNo, CLng(pbytCurPClass), "SCG", pdteDate, "Apply", "") = True Then Exit Sub
    ' Save last date so that it is uniquely last
    lPCls = g_UnFmt(CBA_BTF_Runtime.CBA_BTF_pclassDecypher(cbx_PCls.Value), "lng")
    If dtDateLastWritten = Format(Now(), "dd/mm/yyyy hh:nn") And lCGNoLastWritten = pcurCGNo And lSCGNoLastWritten = pcurSCGNo And lstcPCls = lPCls Then
        MsgBox CBA_FCAST_1Min, vbOKOnly
        Exit Sub
    End If
    
    dtDateLastWritten = Format(Now(), "dd/mm/yyyy hh:nn"): lCGNoLastWritten = pcurCGNo: lSCGNoLastWritten = pcurSCGNo
    pbHasBeenUpdated = False
    
    Set CN = New ADODB.Connection
    Set RS = New ADODB.Recordset
    CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("ForeCast") & ";" ';INTEGRATED SECURITY=sspi;"
    lCG = Me.cbx_CG.Column(0, Me.cbx_CG.ListIndex)
    lSCG = Me.lbx_SCGNo.Column(0, Me.lbx_SCGNo.ListIndex)
''    lPCls = g_UnFmt(CBA_BTF_Runtime.CBA_BTF_pclassDecypher(cbx_PCls.Value), "lng")
    lstcPCls = lPCls
    For lLblNo = 1 To 12
        sDate = CDate("01-" & Me("M" & lLblNo & "_LBL").Caption)
        If bReForecast = False Then
            strSQL = "INSERT INTO SCGData (ProductClass, CG, SCG, MonthNo,YearNo, ForecastDate, FUplift, FRetail, FMarginP, UserName,DateTimeSubmitted )" & Chr(10)
        Else
            strSQL = "INSERT INTO SCGData (ProductClass, CG, SCG, MonthNo, YearNo, ForecastDate, FReUplift, FReRetail,FReMarginP, UserName,DateTimeSubmitted )" & Chr(10)
        End If
        strSQL = strSQL & "VALUES (" & lPCls & "," & lCG & "," & lSCG & "," & Month(sDate) & "," & Year(sDate) & "," & g_GetSQLDate(sDate, "mm/dd/yyyy") _
           & "," & Round(pCGDta(pbytCurPClass).Uplift((plYearIdx * 12) + lLblNo), 5) & "," & Round(pCGDta(pbytCurPClass).ForeRetail((plYearIdx * 12) + lLblNo), 5) & "," & Round(pCGDta(pbytCurPClass).ForeRCVMargin((plYearIdx * 12) + lLblNo), 5) _
           & ",'" & CBA_User & "'," & g_GetSQLDate(dtDateLastWritten, "mm/dd/yyyy hh:nn") & " );"
        RS.Open strSQL, CN
    Next
Exit_Routine:
    MsgBox CBA_FCAST_Apply, vbOKOnly
    On Error Resume Next
    
    CN.Close
    Set CN = Nothing
    Set RS = Nothing
End Sub

Sub UserForm_Initialize()
    Dim sSQL As String
    Dim lRow As Long
    pbSetupIP = True
    Me.StartUpPosition = 0
    Me.Top = Application.Top + ((Application.Height / 4) * 2.75) '- (Me.Height / 4)
    Me.Left = Application.Left + (Application.Width / 2) - (Me.Width / 2)
''    Set pWkBk = ActiveWorkbook
    Set pWksht = ActiveWorkbook.ActiveSheet
    Me.lblVersion.Caption = CBA_getVersionStatus(g_GetDB("ForeCast"), CBA_FCAST_Ver, "Forcasting Tool", "FCast")
    
    ' Fill the Commodity Group list box
    sSQL = "SELECT CGNo, Description FROM cbis599p.dbo.CommodityGroup WHERE CGNo <> 2 AND CGNo <> 55 ;"
    CBA_DBtoQuery = 599
    bPassOK = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", "CBIS_QUERY", CBA_BasicFunctions.TranslateServerName(CBA_DBtoQuery, Date), "SQLNCLI10", sSQL, 120, , , False)  'Runs CBA_DB_Connection module to create connection to dtabase and run query
    If bPassOK = True Then
        Me.cbx_CG.Clear
        For lRow = 0 To UBound(CBA_CBISarr, 2)
            Me.cbx_CG.AddItem NZ(CBA_CBISarr(0, lRow), 0)
            Me.cbx_CG.List(lRow, 1) = NZ(CBA_CBISarr(0, lRow), 0) & "-" & NZ(CBA_CBISarr(1, lRow), "")
        Next
    End If
    
    Me.cbx_PCls.AddItem "Core Range"
    Me.cbx_PCls.AddItem "Food Specials"
    Me.cbx_PCls.AddItem "Non-Food Specials"
    Me.cbx_PCls.AddItem "Seasonal"
    Me.cbx_PCls.Value = "Core Range"
    pbytCurPClass = 1
    cbx_Chart.AddItem "POS Quantity"
    cbx_Chart.AddItem "POS Retail"
''    cbx_Chart.AddItem "Cost per Unit"
    psDataUsed = "POS Quantity"
    cbx_Chart.Value = psDataUsed
    pdteDate = CDate("01/01/" & Year(Date))
    plYearIdx = 0                         ' @RWFC 200107 Added but assumed
    Call SetMonthStart(0)                ' Set the labels and the month positioning
    pbSetupIP = False
End Sub

Private Sub cbx_CG_Change()
    Dim sSQL As String, lRow As Long
    ' Fill the Sub Commodity Group list box
    If pbSetupIP = True Then Exit Sub
    If NZ(Me.cbx_CG.ListIndex, -1) < 0 Then Exit Sub
    sSQL = "SELECT SCGNo, Description AS SCGDesc FROM cbis599p.dbo.SubCommodityGroup WHERE CGNo=" & Me.cbx_CG.Column(0, Me.cbx_CG.ListIndex) & ";"
    CBA_DBtoQuery = 599
    bPassOK = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", "CBIS_QUERY", CBA_BasicFunctions.TranslateServerName(CBA_DBtoQuery, Date), "SQLNCLI10", sSQL, 120, , , False)  'Runs CBA_DB_Connection module to create connection to dtabase and run query
    If bPassOK = True Then
        Me.lbx_SCGNo.Clear
        For lRow = 0 To UBound(CBA_CBISarr, 2)
            Me.lbx_SCGNo.AddItem NZ(CBA_CBISarr(0, lRow), 0)
            Me.lbx_SCGNo.List(lRow, 1) = NZ(CBA_CBISarr(0, lRow), 0) & "-" & NZ(CBA_CBISarr(1, lRow), "")
        Next
    End If
    pcurCGNo = CLng(Val(cbx_CG.Value))
    pb_NoDataReturned = True
    Call p_PopulateExt
    
End Sub

Private Sub cbx_Chart_Change()
    If pcurCGNo = 0 Then Exit Sub
    psDataUsed = cbx_Chart.Value
    If pbSetupIP = False Then CBA_BTF_Runtime.CBA_BTF_ChartChange pWksht, pcurSCGDesc, psDataUsed, plLastRow
End Sub

Private Sub cbx_PCls_Change()
    pbytCurPClass = g_UnFmt(CBA_BTF_Runtime.CBA_BTF_pclassDecypher(cbx_PCls.Value), "lng")
''    If pcurCGNo > 0 Then Call lbx_SCGNo_Change
    If pcurSCGNo = 0 Or pcurCGNo = 0 Then Exit Sub
    ''Debug.Print pbytCurPClass & ",";
    pb_NoDataReturned = Not pCGDta(pbytCurPClass).isDataContained
    Call p_PopulateExt
End Sub


''Private Sub lbx_SCGNo_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
''    If p_TestApplied() = True Then
''        Cancel = vbCancel
''    End If
''End Sub

Private Sub lbx_SCGNo_Change()

    Dim a As Long
    
    If pbSetupIP Or Me.lbx_SCGNo.ListIndex = -1 Then Exit Sub
    Me.Hide
    pcurSCGNo = Me.lbx_SCGNo.Column(0, Me.lbx_SCGNo.ListIndex)
    pcurSCGDesc = CBA_BasicFunctions.g_Right(Me.lbx_SCGNo.Column(1, Me.lbx_SCGNo.ListIndex), 3)
    If CBA_BasicFunctions.isRunningSheetDisplayed Then
        CBA_BasicFunctions.RunningSheetAddComment 7, 4, "Pulling Data for: " & pcurSCGDesc
    Else
        CBA_BasicFunctions.CBA_Running "SubCommodityGroup Level Forecasting"
        CBA_BasicFunctions.RunningSheetAddComment 7, 4, "Pulling Data for: " & pcurSCGDesc
    End If
    Application.ScreenUpdating = False
    
    Set CBA_COM_CBISCN = New ADODB.Connection
    With CBA_COM_CBISCN
        .ConnectionTimeout = 100
        .CommandTimeout = 300
        .Open "Provider= SQLNCLI10; DATA SOURCE= 599DBL01; ;INTEGRATED SECURITY=sspi;"
    End With
    
    pdtDate = BTF_Start_Date("01/12/" & Year(Date)) ' @RWFC 200107 Allied it to the current date regardless of selection
    
    ' Load all the product classes
    For a = 1 To 4
        Set pCGDta(a) = New CBA_BTF_SCGNonProd
        If CBA_BasicFunctions.isRunningSheetDisplayed Then CBA_BasicFunctions.RunningSheetAddComment 7, 4, "Retrieving Data: ProductClass: " & CBA_BTF_pclassDecypher(CStr(a))
        pCGDta(a).CreateSCG pcurCGNo, pcurSCGNo, a, pdtDate
    Next
''    ' Set the current class
''''    pbytCurPClass = g_UnFmt(CBA_BTF_Runtime.CBA_BTF_pclassDecypher(cbx_PCls.Value), "lng")
    pb_NoDataReturned = Not pCGDta(pbytCurPClass).isDataContained
    CBA_COM_CBISCN.Close
    Set CBA_COM_CBISCN = Nothing
    Call p_PopulateExt("")
End Sub

Private Sub PopulateForm(ByVal StartMonth As Long, ByVal EndMonth As Long)
    ' Populate the form
    Dim yradd As Long, madd As Long, a As Long
''    Dim RS As ADODB.Recordset, CN As ADODB.Connection
    If EndMonth < StartMonth Then
        yradd = 1: madd = EndMonth - 12
    Else
''        yradd = 0: madd = Month(Date)
        yradd = 0: madd = 0
    End If
    
    ' If no data...
    If pb_NoDataReturned Then
        For a = StartMonth To EndMonth
            Me.Controls("M" & a - madd & "_PYRCVMargin").Caption = Format(0, "0.00%")
            Me.Controls("M" & a - madd & "_PYPOSRET").Caption = Format(0, "#,0")
            Me.Controls("M" & a - madd & "_RCVMargin").Value = Format(0, "0.00%")
            Me.Controls("M" & a - madd & "_Uplift").Value = Format(0, "0.00%")
            Me.Controls("M" & a - madd & "_FPOSRET").Value = Format(0, "#,0")
            Me.Controls("P" & a - madd & "_Uplift").Value = Format(0, "0.00 %")
            Me.Controls("P" & a - madd & "_FPOSRET").Value = Format(0, "#,0")
            Me.Controls("P" & a - madd & "_RCVMargin").Value = Format(0#, "0.00 %")
        Next
        Exit Sub
    End If
    
    ' Load data from the existing loaded arrays
    For a = StartMonth To EndMonth
        If pCGDta(pbytCurPClass).Description <> "" Then
            With pCGDta(pbytCurPClass)
                Me.Controls("M" & a - madd & "_PYRCVMargin").Caption = Format(.RCVMargin((plYearIdx * 12) + a) * 100, "0.00") & " %"
                Me.Controls("M" & a - madd & "_PYPOSRET").Caption = Format(.POSRET((plYearIdx * 12) + a), "#,0")
                Me.Controls("P" & a - madd & "_Uplift").Value = Format(.PUplift((plYearIdx * 12) + a) * 100, "0.00") & " %"
                Me.Controls("P" & a - madd & "_RCVMargin").Value = Format(.PForeRCVMargin((plYearIdx * 12) + a) * 100, "0.00") & " %"
                Me.Controls("P" & a - madd & "_FPOSRET").Value = Format(.PForeRetail((plYearIdx * 12) + a), "#,0")
                Call pAll_Changes("M" & a - madd & "_RCVMargin_AfterUpdate", , True)
                Call pAll_Changes("M" & a - madd & "_Uplift_AfterUpdate", , True)
                Call pAll_Changes("M" & a - madd & "_FPOSRET_AfterUpdate", , True)
            End With
        End If
    Next a
Exit_Routine:
    On Error Resume Next
    Exit Sub
    
End Sub

Private Function PopulateSheet() As Long
    Dim m As Long, mref As Long, lElemntNo As Long, lRowNo As Long, a As Long, dblVal As Single
    m = Month(Date) - 1
    lRowNo = 7
    If pb_NoDataReturned Then
        For lElemntNo = 1 To 12
            lRowNo = lRowNo + 1
            pWksht.Cells(lRowNo, 4).Value = lRowNo
            pWksht.Cells(lRowNo, 5).Value = lElemntNo
            pWksht.Cells(lRowNo, 6).Value = Format(0, "#,0")
            pWksht.Cells(lRowNo, 7).Value = Format(0, "0.0 %")
            pWksht.Cells(lRowNo, 8).Value = Format(0, "$#,0")
            pWksht.Cells(lRowNo, 9).Value = Format(0, "0.00 %")
            pWksht.Cells(lRowNo, 10).Value = Format(0, "0.00 %")
        Next
    Else
''        For lElemntNo = 12 To 1 Step -1
        For lElemntNo = 1 To 12
            GoSub GSDoRow
        Next
''        lElemntNo = 12
''        GoSub GSDoRow
        CBA_BasicFunctions.CBA_HeatMap Range(pWksht.Cells(8, 7), pWksht.Cells(lRowNo, 7))
        CBA_BasicFunctions.CBA_HeatMap Range(pWksht.Cells(8, 9), pWksht.Cells(lRowNo, 9))
        CBA_BasicFunctions.CBA_HeatMap Range(pWksht.Cells(8, 10), pWksht.Cells(lRowNo, 10))
    End If

    PopulateSheet = lRowNo
    Exit Function
    
GSDoRow:
    If pCGDta(pbytCurPClass).Description <> "" Then
        lRowNo = lRowNo + 1
        pWksht.Cells(lRowNo, 4).Value = pCGDta(pbytCurPClass).YearNo((plYearIdx * 12) + lElemntNo)
        pWksht.Cells(lRowNo, 5).Value = lElemntNo
        pWksht.Cells(lRowNo, 6).Value = Format(pCGDta(pbytCurPClass).POSQTY((plYearIdx * 12) + lElemntNo), "#,0")
        Call CBA_BT_FmtCellVals(pWksht.Cells(lRowNo, 7), pCGDta(pbytCurPClass).POSYOYQTY((plYearIdx * 12) + lElemntNo))
        If lElemntNo = 1 Then mref = 12 Else mref = lElemntNo - 1
        pWksht.Cells(lRowNo, 8).Value = Format(pCGDta(pbytCurPClass).POSRET((plYearIdx * 12) + lElemntNo), "$#,0")
        If NZ(pCGDta(pbytCurPClass).POSYOYRET((plYearIdx * 12) + lElemntNo), 0) <> 0 Then
            Call CBA_BT_FmtCellVals(pWksht.Cells(lRowNo, 9), (pCGDta(pbytCurPClass).POSRET((plYearIdx * 12) + lElemntNo) - pCGDta(pbytCurPClass).POSYOYRET((plYearIdx * 12) + lElemntNo)) / pCGDta(pbytCurPClass).POSYOYRET((plYearIdx * 12) + lElemntNo))
        End If
        Call CBA_BT_FmtCellVals(pWksht.Cells(lRowNo, 10), pCGDta(pbytCurPClass).RCVMargin((plYearIdx * 12) + lElemntNo))
    End If
    Return
End Function

Private Sub SetMonthStart(lPlusMinus As Long)
''    Static bHasBeenRun As Boolean
    Dim LBL As Control, lElemntNo As Long, lColno As Long, bPlus As Boolean, bMinus As Boolean, lMM As Long, lYY As Long, dtDate As Date '', sDate As String
    
    bPlus = CBA_BT_FwdBwd(pdteDate, 1)
    bMinus = CBA_BT_FwdBwd(pdteDate, -1)
    ' Set the current position of the Month Labels
    If lPlusMinus = 1 And bPlus Then
        pdteDate = DateAdd("M", 12, pdteDate)
        plYearIdx = plYearIdx + 1
    ElseIf lPlusMinus = -1 And bMinus Then
        pdteDate = DateAdd("M", -12, pdteDate)
        plYearIdx = plYearIdx - 1
    End If
    If lPlusMinus <> 0 Then   ' Test the new position to see what to do with the buttons
        bPlus = CBA_BT_FwdBwd(pdteDate, 1)
        If plYearIdx + 1 > 2 Then bPlus = False                                         ' @RWFC 200107 Stop too many elements
        bMinus = CBA_BT_FwdBwd(pdteDate, -1)
    End If
    ' Set the select date at the beginning of December in the prior year
    pdtDate = BTF_Start_Date("01/12/" & Year(pdteDate) - 1)
    ' (pdteDate is set at the 01/01/ of the selected year)
    ' Set the button enabled status according to their status
    Me.but_6Less.Enabled = bMinus
    Me.but_6more.Enabled = bPlus
    ' If the labels have been setup and only the position is required, return it
    'If bHasBeenRun = True And lPlusMinus = 0 Then Exit Sub
    ' Set the labels and the worksheet so that they reflect the current position
    dtDate = DateAdd("M", -1, pdteDate)
    lColno = 9: lMM = Month(dtDate): lYY = Year(dtDate)
    For Each LBL In Me.Controls
        If InStr(1, LBL.Name, "_LBL") > 0 Then
            lElemntNo = CLng(Mid(LBL.Name, 2, InStr(1, LBL.Name, "_LBL") - 2))
            LBL.Caption = Format(DateAdd("M", lElemntNo, "01/" & lMM & "/" & lYY), "MMM-YY")
            lColno = lColno + 1
            pWksht.Cells(2, lColno).Value = DateAdd("M", lElemntNo, "01/" & lMM & "/" & lYY)
        End If
    Next
    ' Set the flag to say it has been run
''    bHasBeenRun = True
    If lPlusMinus <> 0 Then
        ' Populate the form, worksheet etc
        Call p_PopulateExt("")
    End If
End Sub

Private Function p_TestEntry() As Boolean
    On Error Resume Next
    p_TestEntry = True
    If Me.cbx_CG.ListIndex < 0 Then
        p_TestEntry = False
        MsgBox "No Commodity Group has been selected", vbOKOnly
    ElseIf Me.cbx_PCls.ListIndex < 0 Then
        p_TestEntry = False
        MsgBox "No Product Class has been selected", vbOKOnly
    ElseIf Me.lbx_SCGNo.ListIndex < 0 Then
        p_TestEntry = False
        MsgBox "No Sub Group has been selected", vbOKOnly
    End If
End Function

Private Sub p_TestLastForecasts(ByVal DspMsg As String)
    ' Test to see if there are other group forecasts (i.e. SCG or Product)
    Static lLastCGNo As Long, LastPClass As Byte, LastYear As Long, SavedMsg As String
    Dim bNeedNewTest As Boolean
    If DspMsg = "DspSaved" Then
        If SavedMsg <> "" Then MsgBox SavedMsg, vbOKOnly
        SavedMsg = ""
        Exit Sub
    End If
    If lLastCGNo <> pcurCGNo Or LastPClass <> pbytCurPClass Or LastYear <> Year(pdteDate) Then bNeedNewTest = True
    If bNeedNewTest Then
        pbXLevel = BTF_LastFCastTest(pcurCGNo, CLng(pbytCurPClass), "SCG", pdteDate, "Test", DspMsg)
    Else
        DspMsg = ""
    End If
    lLastCGNo = pcurCGNo: LastPClass = pbytCurPClass: LastYear = Year(pdteDate)
    If DspMsg <> "" Then SavedMsg = DspMsg
    DspMsg = ""
End Sub

Sub UserForm_Terminate()
    Application.DisplayAlerts = False
    On Error Resume Next
    pWksht.Delete
''    pWkBk.Close
    Err.Clear
    On Error GoTo 0
    Application.DisplayAlerts = True
End Sub

Private Sub p_PopulateExt(Optional sType As String = "Init")
    ' Populate the form, worksheet etc
    Application.ScreenUpdating = False
    Call p_TestLastForecasts("DspMsg") ' Test to see if there are other group forecasts (i.e. SCG or Product)
    Me.MousePointer = fmMousePointerHourGlass
    PopulateForm 1, 12
    plLastRow = PopulateSheet
''    Call cbx_Chart_Change
    If CBA_BasicFunctions.isRunningSheetDisplayed Then CBA_BasicFunctions.CBA_Close_Running
    CBA_BTF_Runtime.UpdateTotals pWksht, pCGDta, pbytCurPClass, plYearIdx
    If sType = "" Then
        Me.Show vbModeless
    End If
    Call cbx_Chart_Change
    pWksht.Activate
    pWksht.Select
    Application.ScreenUpdating = True
    Call p_TestLastForecasts("DspSaved") ' Dsplay the msg that was set earlier
    Me.MousePointer = fmMousePointerDefault
End Sub

Private Sub A2A_Uplift_AfterUpdate()
    Call pAll_Changes("A2A_Uplift_AfterUpdate")
End Sub

Private Sub A2A_RCVMargin_AfterUpdate()
    Call pAll_Changes("A2A_RCVMargin_AfterUpdate")
End Sub

Private Sub A2A_FPOSRET_AfterUpdate()
    Call pAll_Changes("A2A_FPOSRET_AfterUpdate")
End Sub

Private Sub M1_Uplift_AfterUpdate()
    Call pAll_Changes("M1_Uplift_AfterUpdate")
End Sub

Private Sub M1_FPOSRET_AfterUpdate()
    Call pAll_Changes("M1_FPOSRET_AfterUpdate")
End Sub

Private Sub M1_RCVMargin_AfterUpdate()
    Call pAll_Changes("M1_RCVMargin_AfterUpdate")
End Sub



Private Sub M2_Uplift_AfterUpdate()
    Call pAll_Changes("M2_Uplift_AfterUpdate")
End Sub

Private Sub M2_FPOSRET_AfterUpdate()
    Call pAll_Changes("M2_FPOSRET_AfterUpdate")
End Sub

Private Sub M2_RCVMargin_AfterUpdate()
    Call pAll_Changes("M2_RCVMargin_AfterUpdate")
End Sub



Private Sub M3_Uplift_AfterUpdate()
    Call pAll_Changes("M3_Uplift_AfterUpdate")
End Sub

Private Sub M3_FPOSRET_AfterUpdate()
    Call pAll_Changes("M3_FPOSRET_AfterUpdate")
End Sub

Private Sub M3_RCVMargin_AfterUpdate()
    Call pAll_Changes("M3_RCVMargin_AfterUpdate")
End Sub



Private Sub M4_Uplift_AfterUpdate()
    Call pAll_Changes("M4_Uplift_AfterUpdate")
End Sub

Private Sub M4_FPOSRET_AfterUpdate()
    Call pAll_Changes("M4_FPOSRET_AfterUpdate")
End Sub

Private Sub M4_RCVMargin_AfterUpdate()
    Call pAll_Changes("M4_RCVMargin_AfterUpdate")
End Sub



Private Sub M5_Uplift_AfterUpdate()
    Call pAll_Changes("M5_Uplift_AfterUpdate")
End Sub

Private Sub M5_FPOSRET_AfterUpdate()
    Call pAll_Changes("M5_FPOSRET_AfterUpdate")
End Sub

Private Sub M5_RCVMargin_AfterUpdate()
    Call pAll_Changes("M5_RCVMargin_AfterUpdate")
End Sub



Private Sub M6_Uplift_AfterUpdate()
    Call pAll_Changes("M6_Uplift_AfterUpdate")
End Sub

Private Sub M6_FPOSRET_AfterUpdate()
    Call pAll_Changes("M6_FPOSRET_AfterUpdate")
End Sub

Private Sub M6_RCVMargin_AfterUpdate()
    Call pAll_Changes("M6_RCVMargin_AfterUpdate")
End Sub



Private Sub M7_Uplift_AfterUpdate()
    Call pAll_Changes("M7_Uplift_AfterUpdate")
End Sub

Private Sub M7_FPOSRET_AfterUpdate()
    Call pAll_Changes("M7_FPOSRET_AfterUpdate")
End Sub

Private Sub M7_RCVMargin_AfterUpdate()
    Call pAll_Changes("M7_RCVMargin_AfterUpdate")
End Sub



Private Sub M8_Uplift_AfterUpdate()
    Call pAll_Changes("M8_Uplift_AfterUpdate")
End Sub

Private Sub M8_FPOSRET_AfterUpdate()
    Call pAll_Changes("M8_FPOSRET_AfterUpdate")
End Sub

Private Sub M8_RCVMargin_AfterUpdate()
    Call pAll_Changes("M8_RCVMargin_AfterUpdate")
End Sub



Private Sub M9_Uplift_AfterUpdate()
    Call pAll_Changes("M9_Uplift_AfterUpdate")
End Sub

Private Sub M9_FPOSRET_AfterUpdate()
    Call pAll_Changes("M9_FPOSRET_AfterUpdate")
End Sub

Private Sub M9_RCVMargin_AfterUpdate()
    Call pAll_Changes("M9_RCVMargin_AfterUpdate")
End Sub



Private Sub M10_Uplift_AfterUpdate()
    Call pAll_Changes("M10_Uplift_AfterUpdate")
End Sub

Private Sub M10_FPOSRET_AfterUpdate()
    Call pAll_Changes("M10_FPOSRET_AfterUpdate")
End Sub

Private Sub M10_RCVMargin_AfterUpdate()
    Call pAll_Changes("M10_RCVMargin_AfterUpdate")
End Sub



Private Sub M11_Uplift_AfterUpdate()
    Call pAll_Changes("M11_Uplift_AfterUpdate")
End Sub

Private Sub M11_FPOSRET_AfterUpdate()
    Call pAll_Changes("M11_FPOSRET_AfterUpdate")
End Sub

Private Sub M11_RCVMargin_AfterUpdate()
    Call pAll_Changes("M11_RCVMargin_AfterUpdate")
End Sub



Private Sub M12_Uplift_AfterUpdate()
    Call pAll_Changes("M12_Uplift_AfterUpdate")
End Sub

Private Sub M12_FPOSRET_AfterUpdate()
    Call pAll_Changes("M12_FPOSRET_AfterUpdate")
End Sub

Private Sub M12_RCVMargin_AfterUpdate()
    Call pAll_Changes("M12_RCVMargin_AfterUpdate")
End Sub






