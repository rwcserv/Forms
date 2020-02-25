VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CBA_BTF_ProductLevel 
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   25680
   OleObjectBlob   =   "CBA_BTF_ProductLevel.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "CBA_BTF_ProductLevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit       '  Product @CBA_BTF  Changed 12/04/2019

' RW@ 190102 Date Issues Hard-Coded for a while not tested on here as this is in hold mode
' RW@ 190107 Added Mousewheel - Needs Testing as doesn't seem to work

Private pCGs As Variant
Private pbSetupIP As Boolean
Private pSCGDta() As CBA_BTF_ProdData
Private pcurProd As Long, pcurProdPoint As Long
'Private ForeCastExists(1 To 12) As Boolean
Private psDataUsed As String
Private pWksht As Worksheet, pWkBk As Workbook
Private plLastRow As Long
Private pcurProdDesc As String
Private pCBA_pSCGDta As CBA_BTF_SCG
Private pcurCGNo As Long, pcurSCGNo As Long, pbytCurPClass As Byte, curlPClass As Long
Private plPosition As Long, pdtDate As Date, pdteDate As Date
Private pbXLevel As Boolean                                 ' Indication that the item is not the latest item selected
Private pbHasBeenUpdated As Boolean

'#RW Added new mousewheel routines 190701
Private Sub lbx_Prod_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call mw_SetBoxHook(lbx_Prod)
End Sub

' This form will retrieve prior values from the CBis database by product, which can be used to forecast subsequent costs/sales

Private Sub but_6Less_Click()
    Call SetMonthStart(-1) ' Set the labels and the month positioning -1
End Sub

Private Sub but_6More_Click()
    Call SetMonthStart(1) ' Set the labels and the month positioning +1.
End Sub

Private Sub pAll_Changes(sFieldToBeChanged As String, Optional ByVal dblVal As Single = -1.1111, Optional ByVal bSetupIP As Boolean = False)
    ' Handle and format all changes onto the screen - it may even format it when it is initiated ...
    Dim lLbl As Long, bDspMsg As Boolean, bDspMsgb As Boolean, bBlankPull As Boolean
    Static bInUse As Boolean
    ' If fields are not being initialised...
    If bSetupIP = False Then
        If p_TestEntry = False Then Exit Sub
    End If
    If bInUse = True Then ' If this procedure is iteratively called (i.e. Change to the field in this proc calls it again), hop out
        Exit Sub
    End If
    bInUse = True: bBlankPull = False
    ' Get the element no - now called the label no
    lLbl = g_GetNo(sFieldToBeChanged)
    Select Case sFieldToBeChanged
        
        Case "A2A_Uplift_AfterUpdate"
            If g_IsNumeric(Me.A2A_Uplift.Value, False) = True Then
                dblVal = g_UnFmt(Me.A2A_Uplift.Value, "sng")
                For lLbl = 1 To 12
                    If Not pSCGDta(lLbl, pcurProdPoint) Is Nothing Then
                        GoSub GS_Uplift
                    Else
                        bDspMsgb = True
                    End If
                Next
                p_UpdateTotals True
            Else
                If Trim(Me.A2A_Uplift) > "" Then
                    bDspMsg = True
                Else
                    bBlankPull = True
                End If
            End If
            Me.A2A_Uplift = ""
        
        Case "A2A_AvgPrice_AfterUpdate"
            If g_IsNumeric(Me.A2A_AvgPrice.Value, False) = True Then
                dblVal = g_UnFmt(Me.A2A_AvgPrice.Value, "sng")
                For lLbl = 1 To 12
                    If Not pSCGDta(lLbl, pcurProdPoint) Is Nothing Then
                        GoSub GS_AvgPrice
                    Else
                        bDspMsgb = True
                    End If
                Next
            Else
                If Trim(Me.A2A_AvgPrice.Value) > "" Then
                    bDspMsg = True
                Else
                    bBlankPull = True
                End If
            End If
            Me.A2A_AvgPrice.Value = ""
        
        Case "A2A_AvgUnitCost_AfterUpdate"
            If g_IsNumeric(Me.A2A_AvgUnitCost.Value, False) = True Then
                dblVal = g_UnFmt(Me.A2A_AvgUnitCost.Value, "sng")
                For lLbl = 1 To 12
                    If Not pSCGDta(lLbl, pcurProdPoint) Is Nothing Then
                        GoSub GS_AvgUnitCost
                    Else
                        bDspMsgb = True
                    End If
                Next
            Else
                If Trim(Me.A2A_AvgUnitCost.Value) > "" Then
                    bDspMsg = True
                Else
                    bBlankPull = True
                End If
            End If
            Me.A2A_AvgUnitCost.Value = ""
        
        Case "A2A_FPOSQTY_AfterUpdate"
            If g_IsNumeric(Me.A2A_FPOSQTY.Value, False) = True Then
                dblVal = g_UnFmt(Me.A2A_FPOSQTY.Value, "sng")
                For lLbl = 1 To 12
                    If Not pSCGDta(lLbl, pcurProdPoint) Is Nothing Then
                        GoSub GS_FPOSQTY
                    Else
                        bDspMsgb = True
                    End If
                Next
                p_UpdateTotals True
            Else
                If Trim(Me.A2A_FPOSQTY) > "" Then
                    bDspMsg = True
                Else
                    bBlankPull = True
                End If
            End If
            Me.A2A_FPOSQTY = ""
        
        Case "A2A_FPOSRET_AfterUpdate"
            If g_IsNumeric(Me.A2A_FPOSRET.Value, False) = True Then
                dblVal = g_UnFmt(Me.A2A_FPOSRET.Value, "sng")
                For lLbl = 1 To 12
                    If Not pSCGDta(lLbl, pcurProdPoint) Is Nothing Then
                        GoSub GS_FPOSRET
                    Else
                        bDspMsgb = True
                    End If
                Next
                p_UpdateTotals True
            Else
                If Trim(Me.A2A_FPOSRET) > "" Then
                    bDspMsg = True
                Else
                    bBlankPull = True
                End If
            End If
            Me.A2A_FPOSRET = ""
        
        Case "M" & lLbl & "_Uplift_AfterUpdate"
            If g_IsNumeric(Me("M" & lLbl & "_Uplift").Value, False) = True Then
                dblVal = g_UnFmt(Me("M" & lLbl & "_Uplift").Value, "sng")
                GoSub GS_Uplift
                If Not bSetupIP Then p_UpdateTotals True
            Else
                If Trim(Me("M" & lLbl & "_Uplift")) > "" Then
                    bDspMsg = True
                Else
                    bBlankPull = True
                End If
            End If
            Me("M" & lLbl & "_Uplift") = ""
        
        Case "M" & lLbl & "_FPOSRET_Change"
            If g_IsNumeric(Me("M" & lLbl & "_FPOSRET").Value, False) = True Or dblVal = -1.1111 Then
'                If dblVal = -1.1111 Then
                dblVal = g_UnFmt(Me.Controls("M" & lLbl & "_FPOSRET").Value, "sng")
''                Me.Controls("M" & lLbl & "_FPOSRET").Value = Format(dblVal, "#,0")
''                If bSetupIP = False Then
                    GoSub GS_FPOSRET
                    If Not bSetupIP Then p_UpdateTotals True
'                End If
            Else
                Me("M" & lLbl & "_FPOSRET").Value = pSCGDta(lLbl, pcurProdPoint).ForeRET
                bDspMsg = True
            End If
        
        Case "M" & lLbl & "_FPOSQTY_Change"
            If g_IsNumeric(Me("M" & lLbl & "_FPOSQTY").Value, False) = True Or dblVal = -1.1111 Then
'                If dblVal = -1.1111 Then
                dblVal = g_UnFmt(Me.Controls("M" & lLbl & "_FPOSQTY").Value, "sng")
''                Me.Controls("M" & lLbl & "_FPOSQTY").Value = Format(dblVal, "#,0")
''                If bSetupIP = False Then
                    GoSub GS_FPOSQTY
                    If Not bSetupIP Then p_UpdateTotals True
''                End If
            Else
                Me("M" & lLbl & "_FPOSQTY").Value = pSCGDta(lLbl, pcurProdPoint).ForeQTY
                bDspMsg = True
            End If
        
        Case "M" & lLbl & "_AvgUnitCost_AfterUpdate"
            If g_IsNumeric(Me("M" & lLbl & "_AvgUnitCost").Value, False) = True Or dblVal = -1.1111 Then
'                If dblVal = -1.1111 Then
                dblVal = g_UnFmt(Me.Controls("M" & lLbl & "_AvgUnitCost").Value, "sng")
''                Me("M" & lLbl & "_AvgUnitCost").Value = Format(dblVal, "0.00")
''                If bSetupIP = False Then
                     GoSub GS_AvgUnitCost
''                End If
            Else
                bDspMsg = True
                Me("M" & lLbl & "_AvgUnitCost").Value = pSCGDta(lLbl, pcurProdPoint).CPU
            End If
        
        Case "M" & lLbl & "_AvgPrice_AfterUpdate"
            If g_IsNumeric(Me("M" & lLbl & "_AvgPrice").Value, False) = True Or dblVal = -1.1111 Then
''                If dblVal = -1.1111 Then
                dblVal = g_UnFmt(Me.Controls("M" & lLbl & "_AvgPrice").Value, "sng")
''                Me("M" & lLbl & "_AvgPrice").Value = Format(dblVal, "0.00")
''                If bSetupIP = False Then
                    GoSub GS_AvgPrice
''                End If
            Else
                bDspMsg = True
                Me("M" & lLbl & "_AvgPrice").Value = pSCGDta(lLbl, pcurProdPoint).ForeAvgPrice
            End If
                
    End Select
    ' Set the iterative bool off
    bInUse = False
    ' If lLbl msg has been designated
    If bDspMsgb = True Then
        MsgBox "Some values were not affected as there is no current or prior data - to fill these enter into their individual fields", vbOKOnly
        pbHasBeenUpdated = True ' records have been updated
    ' If lLbl msg has been designated
    ElseIf bDspMsg = True Then
        MsgBox "An invalid entry has been made", vbOKOnly
    ' If fields are not being initialised...
    ElseIf bSetupIP = False Then
        If bBlankPull = False Then pbHasBeenUpdated = True      ' records have been updated
    End If
    Exit Sub
    
    
GS_Uplift:
    If bSetupIP = False Then
        Call TestProductInArray(pcurProd, curlPClass, pcurProdPoint, lLbl) ' Ensure its in the arrays
        pSCGDta(lLbl, pcurProdPoint).UpdateValue "Uplift", dblVal
    End If
    Me.Controls("M" & lLbl & "_FPOSRET") = Format(pSCGDta(lLbl, pcurProdPoint).ForeRET, "#,0")
    Me.Controls("M" & lLbl & "_FPOSQTY") = Format(pSCGDta(lLbl, pcurProdPoint).ForeQTY, "#,0")
Return

    
GS_FPOSRET:
    If bSetupIP = False Then
        Call TestProductInArray(pcurProd, curlPClass, pcurProdPoint, lLbl) ' Ensure its in the arrays
        pSCGDta(lLbl, pcurProdPoint).UpdateValue "ForeRET", dblVal
    End If
    Me.Controls("M" & lLbl & "_FPOSRET") = Format(pSCGDta(lLbl, pcurProdPoint).ForeRET, "#,0")
Return

GS_FPOSQTY:
    If bSetupIP = False Then
        Call TestProductInArray(pcurProd, curlPClass, pcurProdPoint, lLbl) ' Ensure its in the arrays
        pSCGDta(lLbl, pcurProdPoint).UpdateValue "ForeQTY", dblVal
    End If
    Me.Controls("M" & lLbl & "_FPOSQTY") = Format(pSCGDta(lLbl, pcurProdPoint).ForeQTY, "#,0")
Return

GS_AvgUnitCost:
    If bSetupIP = False Then
        Call TestProductInArray(pcurProd, curlPClass, pcurProdPoint, lLbl) ' Ensure its in the arrays
        pSCGDta(lLbl, pcurProdPoint).UpdateValue "ForeCPU", dblVal
    End If
    Me.Controls("M" & lLbl & "_AvgUnitCost") = Format(pSCGDta(lLbl, pcurProdPoint).ForeCPU, "0.00")
Return

GS_AvgPrice:
    If bSetupIP = False Then
        Call TestProductInArray(pcurProd, curlPClass, pcurProdPoint, lLbl) ' Ensure its in the arrays
        pSCGDta(lLbl, pcurProdPoint).UpdateValue "ForeAvgPrice", dblVal
    End If
    Me.Controls("M" & lLbl & "_AvgPrice") = Format(pSCGDta(lLbl, pcurProdPoint).ForeAvgPrice, "0.00")
Return

End Sub

Private Sub btn_Apply_Click()
    ' This routine will format and write the forecasts to the 'Forecast DB' database table ProductData
    ' This data will be used in the several forecasting reports
    Dim strSQL As String, sDate As String, sReturn As String, b As Long, c As Long, bReForecast As Boolean
    Dim CN As ADODB.Connection, RS As ADODB.Recordset
    Static dtDateLastWritten As Date
    
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

    If p_TestEntry = False Then Exit Sub
    If BTF_LastFCastTest(pcurCGNo, curlPClass, "PD", pdteDate, "Apply", "") Then Exit Sub
    ' Save last date so that it is uniquely last
    If dtDateLastWritten = Format(Now(), "dd/mm/yyyy hh:nn") Then
        MsgBox CBA_FCAST_1Min, vbOKOnly
        Exit Sub
    End If
    dtDateLastWritten = Format(Now(), "dd/mm/yyyy hh:nn")
    pbHasBeenUpdated = False
    
    Set CN = New ADODB.Connection
    Set RS = New ADODB.Recordset
    
    CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("ForeCast") & ";" ';INTEGRATED SECURITY=sspi;"FPOSRET
    ' Write the whole array to the database
    For b = LBound(pSCGDta, 2) To UBound(pSCGDta, 2)
        For c = LBound(pSCGDta, 1) To UBound(pSCGDta, 1)
            If Not pSCGDta(c, b) Is Nothing Then
                If curlPClass = pSCGDta(c, b).ProdClass Then
                    strSQL = "INSERT INTO ProductData (ProductCode,CG,SCG,ProductClass,MonthNo,YearNo,ForecastDate"
                    If bReForecast = False Then
                        strSQL = strSQL & ",ForeAvgPrice,ForeAvgUnitCost,ForeAvgPOSQty,ForeAvgPOSRet,UserName,DateTimeSubmitted)" & vbCrLf
                    Else
                        strSQL = strSQL & ",ReForeAvgPrice,ReForeAvgUnitCost,ReForeAvgPOSQty,ReForeAvgPOSRET,UserName,DateTimeSubmitted)" & vbCrLf
                    End If
                    ' Update from the values on form and update from module array
                    strSQL = strSQL & "VALUES ( " & pSCGDta(c, b).productcode & "," & pcurCGNo & "," & pcurSCGNo & "," & curlPClass & "," _
                            & Month(sDate) & "," & Year(sDate) & "," & g_GetSQLDate(sDate, "mm/dd/yyyy") & "," _
                            & Round(pSCGDta(c, b).ForeAvgPrice, 5) & "," & Round(pSCGDta(c, b).ForeCPU, 5) & "," _
                            & pSCGDta(c, b).ForeQTY & "," & pSCGDta(c, b).ForeRET _
                            & ",'" & CBA_User & "'," & g_GetSQLDate(dtDateLastWritten, "mm/dd/yyyy hh:nn") & " )"
            '      '' Debug.Print strSQL
                        RS.Open strSQL, CN
                End If
            End If
        Next
    Next
Exit_Routine:
    MsgBox CBA_FCAST_Apply, vbOKOnly
    On Error Resume Next
    CN.Close
    Set CN = Nothing
    Set RS = Nothing
End Sub

Sub UserForm_Initialize()
    Dim a As Long, b As Long, bfound As Boolean
    Me.lblVersion.Caption = CBA_getVersionStatus(g_GetDB("ForeCast"), CBA_FCAST_Ver, "Forcasting Tool", "FCast")

'IF we need to chenge how many months out a reforecast is considered we change this value:
''ReforecastingMonths = 3
'END of Comment

    pbSetupIP = True
    Me.StartUpPosition = 0
    Me.Top = Application.Top + ((Application.Height / 4) * 2.75) '- (Me.Height / 4)
    Me.Left = Application.Left + (Application.Width / 2) - (Me.Width / 2)
    Set pWkBk = ActiveWorkbook
    Set pWksht = ActiveSheet
    ' Call the sql to fill the CG and SCG combo boxes
    CBA_COM_SQLQueries.CBA_COM_GenPullSQL "CGSCGList"
    pCGs = CBA_CBISarr
    cbx_CG.Clear
    For a = LBound(pCGs, 2) To UBound(pCGs, 2)
        bfound = False
        For b = 0 To cbx_CG.ListCount - 1
            If cbx_CG.List(b) = pCGs(0, a) & " - " & pCGs(1, a) Then
                bfound = True
                Exit For
            End If
        Next b
        If bfound = False And pCGs(0, a) <> 2 And pCGs(0, a) <> 55 Then
            cbx_CG.AddItem pCGs(0, a) & " - " & pCGs(1, a)
        End If
    Next a

    ' Set up the product class
    Me.cbx_PCls.AddItem "Core Range"
    Me.cbx_PCls.AddItem "Food Specials"
    Me.cbx_PCls.AddItem "Non-Food Specials"
    Me.cbx_PCls.AddItem "Seasonal"
    Me.cbx_PCls.Value = "Core Range"
    pbytCurPClass = 1
    curlPClass = CLng(pbytCurPClass)
    ' sET UP CHART TYPE
    cbx_Chart.AddItem "POS Quantity"
    cbx_Chart.AddItem "POS Retail"
''    cbx_Chart.AddItem "Cost per Unit"
    cbx_Chart.Value = "POS Quantity"
    pdteDate = CDate("01/01/" & Year(Date))
    Call SetMonthStart(0)                ' Set the labels and the month positioning
    pbSetupIP = False
End Sub

Private Sub cbx_CG_Change()
    Dim a As Long
    'If pbSetupIP = False Then InterfaceData
        
    cbx_SCG.Clear
    For a = LBound(pCGs, 2) To UBound(pCGs, 2)
        If CStr(pCGs(0, a) & " - " & pCGs(1, a)) = CStr(cbx_CG.Value) Then
            cbx_SCG.AddItem pCGs(2, a) & " - " & pCGs(3, a)
            pcurCGNo = pCGs(0, a)
        End If
    Next a
    If cbx_SCG.Value = cbx_SCG.List(0) Then
        Call cbx_SCG_Change
    Else
        cbx_SCG.Value = cbx_SCG.List(0)
    End If

End Sub

Private Sub cbx_Chart_Change()
    If pcurCGNo = 0 Then Exit Sub
    psDataUsed = cbx_Chart.Value
    If pbSetupIP = False Then CBA_BTF_Runtime.CBA_BTF_ChartChange pWksht, pcurProdDesc, psDataUsed, plLastRow
End Sub

Private Sub cbx_PCls_Change()
    If pbSetupIP = False Then pbytCurPClass = g_UnFmt(CBA_BTF_Runtime.CBA_BTF_pclassDecypher(cbx_PCls.Value), "byt")
    curlPClass = CLng(pbytCurPClass)
    If pcurCGNo > 0 Then Call cbx_SCG_Change
End Sub

Private Sub cbx_SCG_Change()
    Dim b As Long, c As Long, lAdded As Long

    If cbx_SCG.Value = "" Then
        pcurSCGNo = 0
        Exit Sub
    Else
        pcurSCGNo = NZ(Left(cbx_SCG.Value, InStr(1, cbx_SCG.Value, " - ") - 1), 0)
    End If
    If pbSetupIP = False Then
        Me.Hide
        If CBA_BasicFunctions.isRunningSheetDisplayed Then
            CBA_BasicFunctions.RunningSheetAddComment 7, 4, "Pulling Data..."
        Else
            CBA_BasicFunctions.CBA_Running "Product Level Forecasting"
            CBA_BasicFunctions.RunningSheetAddComment 7, 4, "Pulling Data..."
        End If
        
        'InterfaceData
        Set CBA_COM_CBISCN = New ADODB.Connection
        With CBA_COM_CBISCN
            .ConnectionTimeout = 100
            .CommandTimeout = 300
            .Open "Provider= SQLNCLI10; DATA SOURCE= 599DBL01; ;INTEGRATED SECURITY=sspi;"
        End With
        
        Set pCBA_pSCGDta = New CBA_BTF_SCG
        pCBA_pSCGDta.CreateSCG Val(cbx_CG.Value), Val(cbx_SCG.Value)
        ' Add the proper items to the list
        Me.lbx_Prod.Clear
        lAdded = 0
        pSCGDta = pCBA_pSCGDta.getSCGData
        For b = LBound(pSCGDta, 2) To UBound(pSCGDta, 2)
            For c = LBound(pSCGDta, 1) To UBound(pSCGDta, 1)
                If Not pSCGDta(c, b) Is Nothing Then
''                On Error Resume Next
''                    If pSCGDta(c, b).ProductCode = 0 Then
                        ''Err.Clear
''                    ElseIf pbytCurPClass = pSCGDta(c, b).ProdClass Then
                    If curlPClass = pSCGDta(c, b).ProdClass Then
                        lAdded = lAdded + 1
                        Me.lbx_Prod.AddItem pSCGDta(c, b).productcode & "-" & pSCGDta(c, b).Description
                        Exit For
                    End If
                End If
''                On Error GoTo 0
            Next
        Next
        ' If there are no records selected
        If lAdded = 0 Then
            CBA_bFCast_NoDataReturned = True
            Call PopulateForm(1, 12)
            Call PopulateSheet
        Else
            CBA_bFCast_NoDataReturned = False
        End If
        
        CBA_COM_CBISCN.Close
        Set CBA_COM_CBISCN = Nothing
        DoEvents
        If CBA_BasicFunctions.isRunningSheetDisplayed Then CBA_BasicFunctions.CBA_Close_Running
        p_UpdateTotals
        Me.Show
    End If
End Sub

Private Sub lbx_Prod_Change()
    Dim a As Long
    
    Application.ScreenUpdating = False
        Range(pWksht.Cells(8, 1), pWksht.Cells(39, 1)).EntireRow.Clear
        For a = 0 To lbx_Prod.ListCount
            If lbx_Prod.Selected(a) = True Then
                pcurProd = CLng(Trim(Mid(lbx_Prod.List(a), 1, InStr(1, lbx_Prod.List(a), "-") - 1)))
                pcurProdDesc = lbx_Prod.List(a)
                Exit For
            End If
        Next

        ' Get the new ProdPoint
        pcurProdPoint = TestProductInArray(pcurProd, 0)
        If pcurProdPoint > -1 Then
            Call p_TestLastForecasts("DspMsg") ' Test to see if there are other group forecasts (i.e. SCG or Product)
            ' Forcast date logic - make it from 1 to 12
            PopulateForm 1, 12
            plLastRow = PopulateSheet()
            psDataUsed = cbx_Chart.Value
            CBA_BTF_Runtime.CBA_BTF_ChartChange pWksht, pcurProdDesc, psDataUsed, plLastRow
        End If
HopOut:

    Application.ScreenUpdating = True
    Call p_TestLastForecasts("DspSaved") ' Dsplay the msg that was set earlier

End Sub

Private Sub p_UpdateTotals(Optional bFreeze As Boolean = False)

    ' Put the totals for the CG/SCG/PClass at the top af the spreadsheet
    Dim TotRet(1 To 12) As Single, TotQTY(1 To 12) As Single, b As Long, c As Long, TTotRet As Single, TTotQTY As Single
    Dim TCQTY As Single, TCRET As Single
    If bFreeze Then Application.ScreenUpdating = False
    ''    lElemntNo = Month(Date)
    Call g_EraseAry(TotQTY): Call g_EraseAry(TotRet): TTotQTY = 0: TTotRet = 0: TCQTY = 0: TCRET = 0
    
    For b = LBound(pSCGDta, 2) To UBound(pSCGDta, 2)        ' All objs
        For c = LBound(pSCGDta, 1) To UBound(pSCGDta, 1)    ' All Months
            If Not pSCGDta(c, b) Is Nothing Then
                TotQTY(c) = TotQTY(c) + pSCGDta(c, b).ForeQTY
                TotRet(c) = TotRet(c) + (pSCGDta(c, b).ForeQTY * pSCGDta(c, b).ForeAvgPrice)
                TTotQTY = TTotQTY + pSCGDta(c, b).ForeQTY
                TTotRet = TTotRet + (pSCGDta(c, b).ForeQTY * pSCGDta(c, b).ForeAvgPrice)
            End If
            Next
    Next
    For b = LBound(pSCGDta, 2) To UBound(pSCGDta, 2)                    ' All objs
        For c = LBound(pSCGDta, 1) To UBound(pSCGDta, 1)                ' All Months
            If Not pSCGDta(c, b) Is Nothing Then
                If pSCGDta(c, b).ProdClass = CLng(pbytCurPClass) Then     ' Current month only
                    TCQTY = TCQTY + pSCGDta(c, b).ForeQTY
                    TCRET = TCRET + (pSCGDta(c, b).ForeQTY * pSCGDta(c, b).ForeAvgPrice)
                    If NZ(TotRet(c), 0) <> 0 Then pWksht.Cells(3, 9 + c) = Format(TotQTY(c) / TotRet(c), "0.00%")
                    pWksht.Cells(4, 9 + c) = Format(TotRet(c), "#,0")
                End If
            End If
        Next
    Next
    If NZ(TCRET, 0) <> 0 Then pWksht.Cells(3, 9) = Format(TCQTY / TCRET, "0.00%")
    pWksht.Cells(4, 9) = Format(TCRET, "#,0")
    Range(pWksht.Cells(3, 9), pWksht.Cells(4, 9)).Font.ColorIndex = 1
    If NZ(TTotRet, 0) <> 0 Then pWksht.Cells(3, 22) = Format(TTotQTY / TTotRet, "0.00%")
    pWksht.Cells(4, 22) = Format(TTotRet, "#,0")
    
    If bFreeze Then Application.ScreenUpdating = True

''    ' Put the totals for the CG/SCG/PClass at the top af the spreadsheet
''    Dim TotRet As Single, TotQTY As Single, TTotRet As Single, TTotQTY As Single, b As Long, c As Long
''
''    For b = LBound(pSCGDta, 2) To UBound(pSCGDta, 2)        ' All objs
''        TotQTY = 0: TotRet = 0
''        For c = LBound(pSCGDta, 1) To UBound(pSCGDta, 1)    ' All Months
''            If Not pSCGDta(c, b) Is Nothing Then
''                TotQTY = TotQTY + pSCGDta(c, b).ForeQTY
''                TotRet = TotRet + (pSCGDta(c, b).ForeQTY * pSCGDta(c, b).ForeAvgPrice)
''                TTotQTY = TTotQTY + TotQTY
''                TTotRet = TTotRet + TotRet
''            End If
''        Next c
''        pWksht.Cells(3, 9 + c) = Format(TotQTY, "#,0")
''        pWksht.Cells(4, 9 + c) = Format(TotRet, "#,0")
''    Next b
''    pWksht.Cells(3, 22) = Format(TTotQTY, "#,0")
''    pWksht.Cells(4, 22) = Format(TTotRet, "#,0")
    
End Sub

Private Sub PopulateForm(ByVal StartMonth As Long, ByVal EndMonth As Long)
    Dim yradd As Long, madd As Long, a As Long
    Dim RS As ADODB.Recordset, CN As ADODB.Connection
    ' Calc the dates to be used
    If EndMonth < 12 Then
        yradd = 1: madd = EndMonth - 12
    Else
        yradd = 0: madd = 0
    End If
    ' If no data...
    If CBA_bFCast_NoDataReturned Then
        For a = StartMonth To EndMonth
            GoSub GSZeros
        Next a
        Exit Sub
    End If
  ' Debug.Print "=" & pcurProdPoint & "...";
    ' For each month add the calculated data
    For a = StartMonth To EndMonth
        If Not pSCGDta(a, pcurProdPoint) Is Nothing Then
            With pSCGDta(a, pcurProdPoint)
                ''If pbytCurPClass = pSCGDta(a, pcurProdPoint).ProdClass Then
                    If .POSQTY = 0 Then
                        Me.Controls("M" & a - madd & "_PYAvgPrice").Caption = "0.00"
                    Else
                        Me.Controls("M" & a - madd & "_PYAvgPrice").Caption = Format(.POSRET / .POSQTY, "0.00")
                    End If
                    Me.Controls("M" & a - madd & "_PYAvgUnitCost").Caption = Format(.CPU, "0.00")
                    Me.Controls("M" & a - madd & "_PYPOSQTY").Caption = Format(.POSQTY, "#,0")
                    Me.Controls("M" & a - madd & "_PYPOSRET").Caption = Format(.POSRET, "#,0")

                    Call pAll_Changes("M" & a - madd & "_AvgPrice_AfterUpdate", , True)
                    Call pAll_Changes("M" & a - madd & "_AvgUnitCost_AfterUpdate", , True)
                    Call pAll_Changes("M" & a - madd & "_FPOSQTY_Change", , True)
                    Call pAll_Changes("M" & a - madd & "_FPOSRET_Change", , True)
                ''End If
            End With
        Else
            GoSub GSZeros
        End If
    Next a
    
Exit_Routine:
    On Error Resume Next
    Set RS = Nothing
    Set CN = Nothing
    Exit Sub
    
GSZeros:
    Me.Controls("M" & a - madd & "_PYAvgPrice").Caption = "n/a"   ''Format(0, "0.00")
    Me.Controls("M" & a - madd & "_PYAvgUnitCost").Caption = "n/a"   '' Format(0, "0.00")
    Me.Controls("M" & a - madd & "_PYPOSQTY").Caption = "n/a"   '' Format(0, "#,0")
    Me.Controls("M" & a - madd & "_PYPOSRET").Caption = "n/a"   '' Format(0, "#,0")
    Me.Controls("M" & a - madd & "_AvgPrice").Value = Format(0, "0.00")
    Me.Controls("M" & a - madd & "_AvgUnitCost").Value = Format(0, "0.00")
    Me.Controls("M" & a - madd & "_FPOSQTY").Value = Format(0, "#,0")
    Me.Controls("M" & a - madd & "_FPOSRET").Value = Format(0, "#,0")
    Return
    
End Sub

Private Function PopulateSheet() As Long
    Dim a As Long, m As Long, lRowNo As Long, lElemntNo As Long, mref As Long '', testingtheInstance
    On Error GoTo Err_Routine
    
    m = Month(pdtDate) - 1
    lRowNo = 7
    If CBA_bFCast_NoDataReturned Then
        With pWksht
            For lElemntNo = 1 To 12
                lRowNo = lRowNo + 1
               pWksht.Cells(lRowNo, 3).Value = lRowNo
               pWksht.Cells(lRowNo, 4).Value = ""
               pWksht.Cells(lRowNo, 5).Value = lElemntNo
               pWksht.Cells(lRowNo, 6).Value = Format(0, "#,0")
               pWksht.Cells(lRowNo, 7).Value = Format(0, "#,0")
               pWksht.Cells(lRowNo, 8).Value = Format(0, "0.0 %")
               pWksht.Cells(lRowNo, 9).Value = Format(0, "0.0 %")
               pWksht.Cells(lRowNo, 10).Value = Format(0, "0.0 %")
               pWksht.Cells(lRowNo, 11).Value = Format(0, "$#,0")
               pWksht.Cells(lRowNo, 12).Value = Format(0, "0.00 %")
               pWksht.Cells(lRowNo, 13).Value = Format(0, "0.00 %")
               pWksht.Cells(lRowNo, 14).Value = Format(0, "$0.00")
               pWksht.Cells(lRowNo, 15).Value = Format(0, "0.00 %")
            Next
        End With
    Else
        ' To get it in reverse date order
        For lElemntNo = m To 1 Step -1
           GoSub GSDoRow
        Next
        ''            If m = 12 Then Exit Function
        For lElemntNo = 12 To m + 1 Step -1
           GoSub GSDoRow
        Next
''        lRowNo = lRowNo + 1
''        pWksht.Cells(lRowNo, 6).Value = Format(pSCGDta(m + 1, pcurProdPoint).POSQTY, "#,0")
''        pWksht.Cells(lRowNo, 7).Value = Format(pSCGDta(m, pcurProdPoint).POSPYQTY, "#,0")
        
        
        ' Fix up the last item entered as 0 as there is no prior data captured from the prior year again
        pWksht.Cells(lRowNo, 10).Value = ""   ' Format(getFCTotals(m + 1, m + 1, "POSPYQTY", "POSPYQTY"), "0.0%")
        ' Fix up the last item entered from the prior year
        pWksht.Cells(lRowNo, 9).Value = Format(getFCTotals(m + 1, m, "POSQTY", "POSPYQTY"), "0.0%")
        For a = 8 To 12
            If a <> 11 Then CBA_BasicFunctions.CBA_HeatMap Range(pWksht.Cells(8, a), pWksht.Cells(lRowNo, a))
        Next
    End If

Exit_Routine:
    PopulateSheet = lRowNo
    Exit Function

GSDoRow:
    If Not pSCGDta(lElemntNo, pcurProdPoint) Is Nothing Then
        lRowNo = lRowNo + 1
        pWksht.Cells(lRowNo, 3).Value = pSCGDta(lElemntNo, pcurProdPoint).productcode
        pWksht.Cells(lRowNo, 4).Value = pSCGDta(lElemntNo, pcurProdPoint).Year
        pWksht.Cells(lRowNo, 5).Value = pSCGDta(lElemntNo, pcurProdPoint).Month
        pWksht.Cells(lRowNo, 6).Value = Format(pSCGDta(lElemntNo, pcurProdPoint).POSQTY, "#,0")
        pWksht.Cells(lRowNo, 7).Value = Format(pSCGDta(lElemntNo, pcurProdPoint).POSPYQTY, "#,0")
        GoSub GSPercents
        pWksht.Cells(lRowNo, 11).Value = Format(pSCGDta(lElemntNo, pcurProdPoint).POSRET, "$#,0")
        pWksht.Cells(lRowNo, 13).Value = Format(pSCGDta(lElemntNo, pcurProdPoint).RCVMargin, "0.00%")
        pWksht.Cells(lRowNo, 14).Value = Format(pSCGDta(lElemntNo, pcurProdPoint).CPU, "$0.00")
        pWksht.Cells(lRowNo, 15).Value = Format(pSCGDta(lElemntNo, pcurProdPoint).CPUYOY, "0.00%")
    End If
    Return
GSPercents:
    If lElemntNo = 1 Then mref = 12 Else mref = lElemntNo - 1

    ' TY=This Year, LY=Last Year,  TM=This Month, LM=Last Month
    ' Calculate the year's growth in qty { = ((TY MM Vals) - (LY Same MM Vals)) / (LY Same MM Vals)   }
    pWksht.Cells(lRowNo, 8).Value = Format(getFCTotals(lElemntNo, lElemntNo, "POSQTY", "POSPYQTY"), "0.0%")
   
    ' Calculate the year's growth in retail { = ((TY MM $s) - (LY Same MM $s)) / (LY Same MM $s)   }
    pWksht.Cells(lRowNo, 12).Value = Format(getFCTotals(lElemntNo, lElemntNo, "POSRET", "POSPYRET"), "0.0%")
        
    ' Calculate the month's growth +ve or -ve over the last year { = ((TM Vals) - (LM Vals)) / (LM Vals)   }
    pWksht.Cells(lRowNo, 9).Value = Format(getFCTotals(lElemntNo, mref, "POSQTY", "POSQTY"), "0.0%")
    
    ' Calculate the month's growth +ve or -ve over the prior year { = ((TM(LY) Vals) - (LM(LY) Vals)) / (LM(LY) Vals)   }
     pWksht.Cells(lRowNo, 10).Value = Format(getFCTotals(lElemntNo, mref, "POSPYQTY", "POSPYQTY"), "0.0%")
    
    Return

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-PopulateSheet", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "ForeCast", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Private Function getFCTotals(lDividee As Long, lDividedBy As Long, sDividee As String, sDividedBy As String) As Single
    ' Will bring back a double that is the result of two input numbers
    Dim dblDividedBy As Single, dblDividee As Single, lEl As Long, dblEl As Single, sel As String
    getFCTotals = 0
    If Not pSCGDta(lDividedBy, pcurProdPoint) Is Nothing And Not pSCGDta(lDividee, pcurProdPoint) Is Nothing Then
        lEl = lDividedBy: sel = sDividedBy
        GoSub GSEls
        dblDividedBy = dblEl
        lEl = lDividee: sel = sDividee
        GoSub GSEls
        dblDividee = dblEl
        If dblDividedBy > 0 Then getFCTotals = (dblDividee - dblDividedBy) / dblDividedBy
    End If
    Exit Function
GSEls:
    Select Case sel
    Case "POSQTY"
        dblEl = pSCGDta(lEl, pcurProdPoint).POSQTY
    Case "POSPYQTY"
        dblEl = pSCGDta(lEl, pcurProdPoint).POSPYQTY
    Case "POSRET"
        dblEl = pSCGDta(lEl, pcurProdPoint).POSRET
    Case "POSPYRET"
        dblEl = pSCGDta(lEl, pcurProdPoint).POSPYRET
''    Case "POSQTY"
''        dblEl = pSCGDta(lEl, pcurProdPoint).POSQTY
''    Case "POSQTY"
''        dblEl = pSCGDta(lEl, pcurProdPoint).POSQTY
    End Select
    Return
    
End Function

Private Function TestProductInArray(ByVal lProdCode As Long, lPCls As Long, Optional ByRef lHaveArrayEl As Long = -1, _
                                            Optional ByVal sMthReq As String = "", _
                                            Optional ByVal bAddIfDoesntExist As Boolean = False) As Long
    ' PLEASE NOTE: THE SUB-ELEMENTS CONFORM TO THE MONTHS - IE EL 1 = MONTH 1...
    
    ' Get the product b number and return it in a long. Fill in any missing b's or c's if req.
    Dim b As Long, c As Long, bProdFound As Boolean, lbFrom As Long, lbTo As Long, lNewArrayEl As Long
    Dim lMthReq As Long, lYrReq As Long, sDesc As String
    Dim sSQL As String, bPassOK As Boolean
    If sMthReq > "" Then sMthReq = "01/" & sMthReq & "/20" & Right(Me("M" & sMthReq & "_LBL"), 2)
    
    TestProductInArray = -1
    If sMthReq > "" Then lMthReq = Month(CDate(sMthReq)): lYrReq = Year(CDate(sMthReq))
    lbFrom = LBound(pSCGDta, 2): lbTo = UBound(pSCGDta, 2)   ' Set the default b from/to
''    lNewArrayEl = 0
    ' If the element has already been found...
    If lHaveArrayEl > 0 Then lbFrom = lHaveArrayEl: lbTo = lHaveArrayEl
    ' For each b element
    For b = lbFrom To lbTo
        For c = LBound(pSCGDta, 1) To UBound(pSCGDta, 1)
            If pSCGDta(c, b) Is Nothing Then
            ElseIf pSCGDta(c, b).productcode = lProdCode Then
                If lPCls = 0 Then lPCls = pSCGDta(c, b).ProdClass
                lNewArrayEl = b
                bProdFound = True
                sDesc = pSCGDta(c, b).Description
                If lMthReq > 0 Then
                    If pSCGDta(c, b).Month = lMthReq Then GoTo Exit_Routine
                Else
                    If bAddIfDoesntExist = False Then GoTo Exit_Routine
                End If
            End If
        Next c
    Next b
    ' Add the b Element, if it isn't in there - Note: Els will exist but be nothing at this stage
    If bProdFound = False Then
        ReDim Preserve pSCGDta(LBound(pSCGDta, 1) To UBound(pSCGDta, 1), LBound(pSCGDta, 2) To UBound(pSCGDta, 2) + 1)
        lNewArrayEl = UBound(pSCGDta, 2)
    End If
    ' Add the Month Element
    If lMthReq > 0 Then
        For c = LBound(pSCGDta, 1) To UBound(pSCGDta, 1)
            If pSCGDta(c, b) Is Nothing Then
                If c = lMthReq Then
                    ' Get the product details....
                    If sDesc = "" Then
                        sSQL = "SELECT Description " & _
                               "FROM cbis599p.dbo.Product " & _
                               "WHERE ProductCode=" & lProdCode & ";"
                        CBA_DBtoQuery = 599
                        bPassOK = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", "CBIS_QUERY", CBA_BasicFunctions.TranslateServerName(CBA_DBtoQuery, Date), "SQLNCLI10", sSQL, 120, , , False)
                        If bPassOK = True Then
                            sDesc = CBA_CBISarr(0, 0)
                        End If
                    End If
                    Set pSCGDta(lMthReq, b) = New CBA_BTF_ProdData
                    pSCGDta(lMthReq, b).Generate lProdCode, lYrReq, lMthReq, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, sDesc, lPCls
                End If
            End If
        Next c
    End If
Exit_Routine:
    lHaveArrayEl = lNewArrayEl
    TestProductInArray = lNewArrayEl
    
End Function

Private Function SetMonthStart(lPlusMinus As Long) As Long
    Static bHasBeenRun As Boolean
    Dim LBL As Control, lLblNo As Long, lColno As Long, bPlus As Boolean, bMinus As Boolean '', lMM As Long, lYY As Long, dtDate As Date '', sDate As String
    
    bPlus = CBA_BT_FwdBwd(pdteDate, 1)
    bMinus = CBA_BT_FwdBwd(pdteDate, -1)
    ' Set the current position of the Month Labels
    If plPosition = 0 And lPlusMinus < 1 Then
        plPosition = 1
        pdtDate = BTF_Start_Date("01/12/" & Year(pdteDate) - 1)
    ElseIf lPlusMinus = 1 Then
        plPosition = plPosition + 1
        pdteDate = DateAdd("M", 12, pdteDate)
    ElseIf lPlusMinus = -1 Then
        plPosition = plPosition - 1
        pdteDate = DateAdd("M", -12, pdteDate)
    End If
    ' Set the function return
    SetMonthStart = plPosition
    ' Set the button enabled status according to the position
    If CBA_FC_MAX_YRS = 0 Then
        Me.but_6Less.Enabled = False
        Me.but_6more.Enabled = False
    ElseIf plPosition = 0 Then
        Me.but_6Less.Enabled = False
        Me.but_6more.Enabled = True
    ElseIf plPosition = 1 And plPosition < CBA_FC_MAX_YRS Then
        Me.but_6Less.Enabled = True
        Me.but_6more.Enabled = True
    ElseIf plPosition >= CBA_FC_MAX_YRS Then
        Me.but_6Less.Enabled = True
        Me.but_6more.Enabled = False
    End If
    ' If the labels have been setup and only the position is required, return it
    If bHasBeenRun = True And lPlusMinus = 0 Then Exit Function
    ' Set the labels and the worksheet so that they reflect the current position
    lColno = 9
    For Each LBL In Me.Controls
        If InStr(1, LBL.Name, "_LBL") > 0 Then
            lLblNo = CLng(Mid(LBL.Name, 2, InStr(1, LBL.Name, "_LBL") - 2))
            LBL.Caption = Format(DateAdd("M", lLblNo, "01/" & Month(pdtDate) & "/" & Year(pdtDate)), "MMM-YY")
            lColno = lColno + 1
            pWksht.Cells(2, lColno).Value = DateAdd("M", lLblNo, "01/" & Month(pdtDate) & "/" & Year(pdtDate))
        End If
    Next
    ' Set the flag to say it has been run
    bHasBeenRun = True
End Function

Private Function p_TestEntry() As Boolean
    On Error Resume Next
    p_TestEntry = True
    If Me.cbx_CG.ListIndex < 0 Then
        p_TestEntry = False
        MsgBox "No Commodity Group has been selected", vbOKOnly
    ElseIf Me.cbx_PCls.ListIndex < 0 Then
        p_TestEntry = False
        MsgBox "No Product Class has been selected", vbOKOnly
    ElseIf Me.cbx_SCG.ListIndex < 0 Then
        p_TestEntry = False
        MsgBox "No Sub Commodity Group has been selected", vbOKOnly
    ElseIf Me.lbx_Prod.ListIndex < 0 Then
        p_TestEntry = False
        MsgBox "No Product has been selected", vbOKOnly
    End If
End Function

Private Sub p_TestLastForecasts(ByVal DspMsg As String)
    ' Test to see if there are other group forecasts (i.e. SCG or CG) that need to be warned of
    Static lLastCGNo As Long, lLastPClass As Long, SavedMsg As String
    Dim bNeedNewTest As Boolean
    If DspMsg = "DspSaved" Then
        If SavedMsg <> "" Then MsgBox SavedMsg, vbOKOnly
        SavedMsg = ""
        Exit Sub
    End If
    If lLastCGNo <> pcurCGNo Or lLastPClass <> curlPClass Then bNeedNewTest = True
    If bNeedNewTest = True Then
        pbXLevel = BTF_LastFCastTest(pcurCGNo, curlPClass, "PD", pdteDate, "Test", DspMsg)
    Else
        DspMsg = ""
    End If
    lLastCGNo = pcurCGNo: lLastPClass = curlPClass
    If DspMsg <> "" Then SavedMsg = DspMsg
    DspMsg = ""
End Sub

Sub UserForm_Terminate()
    Application.DisplayAlerts = False
    On Error Resume Next
    pWkBk.Close
    Err.Clear
    On Error GoTo 0
    Application.DisplayAlerts = True
End Sub


Private Sub A2A_AvgPrice_AfterUpdate()
    Call pAll_Changes("A2A_AvgPrice_AfterUpdate")
End Sub

Private Sub A2A_AvgUnitCost_AfterUpdate()
    Call pAll_Changes("A2A_AvgUnitCost_AfterUpdate")
End Sub

Private Sub A2A_FPOSQTY_AfterUpdate()
    Call pAll_Changes("A2A_FPOSQTY_AfterUpdate")
End Sub

Private Sub A2A_FPOSRET_AfterUpdate()
    Call pAll_Changes("A2A_FPOSRET_AfterUpdate")
End Sub

Private Sub A2A_Uplift_AfterUpdate()
    Call pAll_Changes("A2A_Uplift_AfterUpdate")
End Sub



Private Sub M1_Uplift_AfterUpdate()
    Call pAll_Changes("M1_Uplift_AfterUpdate")
End Sub

Private Sub M1_FPOSRET_Change()
    Call pAll_Changes("M1_FPOSRET_Change")
End Sub

Private Sub M1_FPOSQTY_Change()
    Call pAll_Changes("M1_FPOSQTY_Change")
End Sub

Private Sub M1_AvgUnitCost_AfterUpdate()
    Call pAll_Changes("M1_AvgUnitCost_AfterUpdate")
End Sub

Private Sub M1_AvgPrice_AfterUpdate()
    Call pAll_Changes("M1_AvgPrice_AfterUpdate")
End Sub


Private Sub M2_Uplift_AfterUpdate()
    Call pAll_Changes("M2_Uplift_AfterUpdate")
End Sub

Private Sub M2_FPOSRET_Change()
    Call pAll_Changes("M2_FPOSRET_Change")
End Sub

Private Sub M2_FPOSQTY_Change()
    Call pAll_Changes("M2_FPOSQTY_Change")
End Sub

Private Sub M2_AvgUnitCost_AfterUpdate()
    Call pAll_Changes("M2_AvgUnitCost_AfterUpdate")
End Sub

Private Sub M2_AvgPrice_AfterUpdate()
    Call pAll_Changes("M2_AvgPrice_AfterUpdate")
End Sub


Private Sub M3_Uplift_AfterUpdate()
    Call pAll_Changes("M3_Uplift_AfterUpdate")
End Sub

Private Sub M3_FPOSRET_Change()
    Call pAll_Changes("M3_FPOSRET_Change")
End Sub

Private Sub M3_FPOSQTY_Change()
    Call pAll_Changes("M3_FPOSQTY_Change")
End Sub

Private Sub M3_AvgUnitCost_AfterUpdate()
    Call pAll_Changes("M3_AvgUnitCost_AfterUpdate")
End Sub

Private Sub M3_AvgPrice_AfterUpdate()
    Call pAll_Changes("M3_AvgPrice_AfterUpdate")
End Sub


Private Sub M4_Uplift_AfterUpdate()
    Call pAll_Changes("M4_Uplift_AfterUpdate")
End Sub

Private Sub M4_FPOSRET_Change()
    Call pAll_Changes("M4_FPOSRET_Change")
End Sub

Private Sub M4_FPOSQTY_Change()
    Call pAll_Changes("M4_FPOSQTY_Change")
End Sub

Private Sub M4_AvgUnitCost_AfterUpdate()
    Call pAll_Changes("M4_AvgUnitCost_AfterUpdate")
End Sub

Private Sub M4_AvgPrice_AfterUpdate()
    Call pAll_Changes("M4_AvgPrice_AfterUpdate")
End Sub


Private Sub M5_Uplift_AfterUpdate()
    Call pAll_Changes("M5_Uplift_AfterUpdate")
End Sub

Private Sub M5_FPOSRET_Change()
    Call pAll_Changes("M5_FPOSRET_Change")
End Sub

Private Sub M5_FPOSQTY_Change()
    Call pAll_Changes("M5_FPOSQTY_Change")
End Sub

Private Sub M5_AvgUnitCost_AfterUpdate()
    Call pAll_Changes("M5_AvgUnitCost_AfterUpdate")
End Sub

Private Sub M5_AvgPrice_AfterUpdate()
    Call pAll_Changes("M5_AvgPrice_AfterUpdate")
End Sub


Private Sub M6_Uplift_AfterUpdate()
    Call pAll_Changes("M6_Uplift_AfterUpdate")
End Sub

Private Sub M6_FPOSRET_Change()
    Call pAll_Changes("M6_FPOSRET_Change")
End Sub

Private Sub M6_FPOSQTY_Change()
    Call pAll_Changes("M6_FPOSQTY_Change")
End Sub

Private Sub M6_AvgUnitCost_AfterUpdate()
    Call pAll_Changes("M6_AvgUnitCost_AfterUpdate")
End Sub

Private Sub M6_AvgPrice_AfterUpdate()
    Call pAll_Changes("M6_AvgPrice_AfterUpdate")
End Sub


Private Sub M7_Uplift_AfterUpdate()
    Call pAll_Changes("M7_Uplift_AfterUpdate")
End Sub

Private Sub M7_FPOSRET_Change()
    Call pAll_Changes("M7_FPOSRET_Change")
End Sub

Private Sub M7_FPOSQTY_Change()
    Call pAll_Changes("M7_FPOSQTY_Change")
End Sub

Private Sub M7_AvgUnitCost_AfterUpdate()
    Call pAll_Changes("M7_AvgUnitCost_AfterUpdate")
End Sub

Private Sub M7_AvgPrice_AfterUpdate()
    Call pAll_Changes("M7_AvgPrice_AfterUpdate")
End Sub


Private Sub M8_Uplift_AfterUpdate()
    Call pAll_Changes("M8_Uplift_AfterUpdate")
End Sub

Private Sub M8_FPOSRET_Change()
    Call pAll_Changes("M8_FPOSRET_Change")
End Sub

Private Sub M8_FPOSQTY_Change()
    Call pAll_Changes("M8_FPOSQTY_Change")
End Sub

Private Sub M8_AvgUnitCost_AfterUpdate()
    Call pAll_Changes("M8_AvgUnitCost_AfterUpdate")
End Sub

Private Sub M8_AvgPrice_AfterUpdate()
    Call pAll_Changes("M8_AvgPrice_AfterUpdate")
End Sub


Private Sub M9_Uplift_AfterUpdate()
    Call pAll_Changes("M9_Uplift_AfterUpdate")
End Sub

Private Sub M9_FPOSRET_Change()
    Call pAll_Changes("M9_FPOSRET_Change")
End Sub

Private Sub M9_FPOSQTY_Change()
    Call pAll_Changes("M9_FPOSQTY_Change")
End Sub

Private Sub M9_AvgUnitCost_AfterUpdate()
    Call pAll_Changes("M9_AvgUnitCost_AfterUpdate")
End Sub

Private Sub M9_AvgPrice_AfterUpdate()
    Call pAll_Changes("M9_AvgPrice_AfterUpdate")
End Sub


Private Sub M10_Uplift_AfterUpdate()
    Call pAll_Changes("M10_Uplift_AfterUpdate")
End Sub

Private Sub M10_FPOSRET_Change()
    Call pAll_Changes("M10_FPOSRET_Change")
End Sub

Private Sub M10_FPOSQTY_Change()
    Call pAll_Changes("M10_FPOSQTY_Change")
End Sub

Private Sub M10_AvgUnitCost_AfterUpdate()
    Call pAll_Changes("M10_AvgUnitCost_AfterUpdate")
End Sub

Private Sub M10_AvgPrice_AfterUpdate()
    Call pAll_Changes("M10_AvgPrice_AfterUpdate")
End Sub


Private Sub M11_Uplift_AfterUpdate()
    Call pAll_Changes("M11_Uplift_AfterUpdate")
End Sub

Private Sub M11_FPOSRET_Change()
    Call pAll_Changes("M11_FPOSRET_Change")
End Sub

Private Sub M11_FPOSQTY_Change()
    Call pAll_Changes("M11_FPOSQTY_Change")
End Sub

Private Sub M11_AvgUnitCost_AfterUpdate()
    Call pAll_Changes("M11_AvgUnitCost_AfterUpdate")
End Sub

Private Sub M11_AvgPrice_AfterUpdate()
    Call pAll_Changes("M11_AvgPrice_AfterUpdate")
End Sub


Private Sub M12_Uplift_AfterUpdate()
    Call pAll_Changes("M12_Uplift_AfterUpdate")
End Sub

Private Sub M12_FPOSRET_Change()
    Call pAll_Changes("M12_FPOSRET_Change")
End Sub

Private Sub M12_FPOSQTY_Change()
    Call pAll_Changes("M12_FPOSQTY_Change")
End Sub

Private Sub M12_AvgUnitCost_AfterUpdate()
    Call pAll_Changes("M12_AvgUnitCost_AfterUpdate")
End Sub

Private Sub M12_AvgPrice_AfterUpdate()
    Call pAll_Changes("M12_AvgPrice_AfterUpdate")
End Sub

