VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CBA_AST_frm_AssocFld 
   Caption         =   "Associated Field Entry"
   ClientHeight    =   12240
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   25125
   OleObjectBlob   =   "CBA_AST_frm_AssocFld.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "CBA_AST_frm_AssocFld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit   ' @CBA_ASyst 190625

Private psField As String, psDate As String, pbRegions As Boolean, pbRegion1 As Boolean, pbRegionR As Boolean, pbDates As Boolean, pbProducts As Boolean, pbProductCode As Boolean, pbComp As Boolean, pbCann As Boolean
Private pbUseCtlDate As Boolean, pbHasBeenSaved As Boolean, pCGs, pbSetupIP As Boolean, pbProductsIP As Boolean, paReg() As String, plProductID As Long
Private ptProdRows(1 To 8) As CBA_AST_clsProdRows, plAuth As Long, pbAllUpd As Boolean, psForm As String

Const plFrmID = 5, CREGS As String = "Min,Der,Stp,Pre,Dan,Bre,Rgy,Jkt", PRODLEN As Long = 30, CANNLEN As Long = 45, COMPLEN As Long = 45, FORMTAG = "AssocFlds"

'#RW Added new mousewheel routines 190701
Private Sub lstProducts_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call mw_SetBoxHook(lstProducts)
End Sub
Private Sub cboCG_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call mw_SetBoxHook(cboCG)
End Sub
Private Sub cboSCG_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call mw_SetBoxHook(cboSCG)
End Sub

Private Sub chk01Regions_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call p_FormatReturn(Mid(Me.ActiveControl.ActiveControl.Name, 4, 2))
    Call p_UpdFld("chk01Regions")
End Sub
Private Sub chk02Regions_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call p_FormatReturn(Mid(Me.ActiveControl.ActiveControl.Name, 4, 2))
    Call p_UpdFld("chk02Regions")
End Sub
Private Sub chk03Regions_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call p_FormatReturn(Mid(Me.ActiveControl.ActiveControl.Name, 4, 2))
    Call p_UpdFld("chk03Regions")
End Sub
Private Sub chk04Regions_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call p_FormatReturn(Mid(Me.ActiveControl.ActiveControl.Name, 4, 2))
    Call p_UpdFld("chk04Regions")
End Sub
Private Sub chk05Regions_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call p_FormatReturn(Mid(Me.ActiveControl.ActiveControl.Name, 4, 2))
    Call p_UpdFld("chk05Regions")
End Sub
Private Sub chk06Regions_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call p_FormatReturn(Mid(Me.ActiveControl.ActiveControl.Name, 4, 2))
    Call p_UpdFld("chk06Regions")
End Sub
Private Sub chk07Regions_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call p_FormatReturn(Mid(Me.ActiveControl.ActiveControl.Name, 4, 2))
    Call p_UpdFld("chk07Regions")
End Sub
Private Sub chk08Regions_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call p_FormatReturn(Mid(Me.ActiveControl.ActiveControl.Name, 4, 2))
    Call p_UpdFld("chk08Regions")
End Sub

Private Sub cmdNA_Click()
    Dim lIdx As Long, sNA As String
    If pbRegions Or pbRegion1 Or pbRegionR Then
    ElseIf pbComp Or pbCann Then
        Me.txtReturn = "N/A"
        sNA = "N/A"
    Else
        For lIdx = 1 To 3
            Me("txtDateF" & lIdx) = ""
            Me("txtDateT" & lIdx) = ""
        Next
    End If
    Call p_FormatReturn(, sNA)
End Sub

Private Sub cmdProdAccept_Click()
    pbProductsIP = False
    Call p_lstProdListChange("Search")
    Me.cmdProdAccept.Visible = False
End Sub

Private Sub cmdSave_Click()
    If Me.txtReturn.BackColor = CBA_Red Then
        MsgBox "There is an error in the date/s entered" & vbCrLf & "Please ensure the dates are consecutive and do not overlap"
        Exit Sub
    Else
        pbHasBeenSaved = True
        If pbRegion1 Then Call p_GetProdRegionClass("Write")
        If pbRegionR Then Call p_UpdProdRegionRecs("Write")
        CBA_strAldiMsg = Me.txtReturn
        CBA_bDataChg = True
        Unload Me
    End If
End Sub

Private Sub lstProducts_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call p_FormatReturn
End Sub

Private Sub txt01Upd_Change()
    Dim SNo As String
    If pbSetupIP = True Then Exit Sub
    SNo = "01"
    If varFldVars.sField3 > "" And UCase(txt01Upd.Value) = "Y" Then
        If p_CalcRegionValues(Me("txt" & SNo & varFldVars.sField3), SNo, varFldVars.sField3, True) = True Then
            ' Routine will process
            Me.cmdSave.Visible = True
        End If
        varFldVars.sField3 = ""
    End If
End Sub

Private Sub txt02Upd_Change()
    Dim SNo As String
    If pbSetupIP = True Then Exit Sub
    SNo = "02"
    If varFldVars.sField3 > "" And UCase(txt02Upd.Value) = "Y" Then
        If p_CalcRegionValues(Me("txt" & SNo & varFldVars.sField3), SNo, varFldVars.sField3, True) = True Then
            ' Routine will process
            Me.cmdSave.Visible = True
        End If
        varFldVars.sField3 = ""
    End If
End Sub

Private Sub txt03Upd_Change()
    Dim SNo As String
    If pbSetupIP = True Then Exit Sub
    SNo = "03"
    If varFldVars.sField3 > "" And UCase(txt03Upd.Value) = "Y" Then
        If p_CalcRegionValues(Me("txt" & SNo & varFldVars.sField3), SNo, varFldVars.sField3, True) = True Then
            ' Routine will process
            Me.cmdSave.Visible = True
        End If
        varFldVars.sField3 = ""
    End If
End Sub

Private Sub txt04Upd_Change()
    Dim SNo As String
    If pbSetupIP = True Then Exit Sub
    SNo = "04"
    If varFldVars.sField3 > "" And UCase(txt04Upd.Value) = "Y" Then
        If p_CalcRegionValues(Me("txt" & SNo & varFldVars.sField3), SNo, varFldVars.sField3, True) = True Then
            ' Routine will process
            Me.cmdSave.Visible = True
        End If
        varFldVars.sField3 = ""
    End If
End Sub

Private Sub txt05Upd_Change()
    Dim SNo As String
    If pbSetupIP = True Then Exit Sub
    SNo = "05"
    If varFldVars.sField3 > "" And UCase(txt05Upd.Value) = "Y" Then
        If p_CalcRegionValues(Me("txt" & SNo & varFldVars.sField3), SNo, varFldVars.sField3, True) = True Then
            ' Routine will process
            Me.cmdSave.Visible = True
        End If
        varFldVars.sField3 = ""
    End If
End Sub

Private Sub txt06Upd_Change()
    Dim SNo As String
    If pbSetupIP = True Then Exit Sub
    SNo = "06"
    If varFldVars.sField3 > "" And UCase(txt06Upd.Value) = "Y" Then
        If p_CalcRegionValues(Me("txt" & SNo & varFldVars.sField3), SNo, varFldVars.sField3, True) = True Then
            ' Routine will process
            Me.cmdSave.Visible = True
        End If
        varFldVars.sField3 = ""
    End If
End Sub

Private Sub txt07Upd_Change()
    Dim SNo As String
    If pbSetupIP = True Then Exit Sub
    SNo = "07"
    If varFldVars.sField3 > "" And UCase(txt07Upd.Value) = "Y" Then
        If p_CalcRegionValues(Me("txt" & SNo & varFldVars.sField3), SNo, varFldVars.sField3, True) = True Then
            ' Routine will process
            Me.cmdSave.Visible = True
        End If
        varFldVars.sField3 = ""
    End If
End Sub

Private Sub txt08Upd_Change()
    Dim SNo As String
    If pbSetupIP = True Then Exit Sub
    SNo = "08"
    If varFldVars.sField3 > "" And UCase(txt08Upd.Value) = "Y" Then
        If p_CalcRegionValues(Me("txt" & SNo & varFldVars.sField3), SNo, varFldVars.sField3, True) = True Then
            ' Routine will process
            Me.cmdSave.Visible = True
        End If
        varFldVars.sField3 = ""
    End If
End Sub

Private Sub txtDateF1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If g_IsDate(Me.txtDateF1, True) Then
        pbUseCtlDate = True
    Else
        pbUseCtlDate = False
        varCal.sDate = psDate
    End If
    Call CBA_BT_CalendarShow(Me.Top, Me.Left, Me.txtDateF1, CBA_DM2Y, , pbUseCtlDate)
''    Call p_UpdIP(Me.txtDateF1, True)
    Call p_FormatReturn
End Sub
Private Sub txtDateF2_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If g_IsDate(Me.txtDateF2, True) Then
        pbUseCtlDate = True
    Else
        pbUseCtlDate = False
        varCal.sDate = psDate
    End If
    Call CBA_BT_CalendarShow(Me.Top, Me.Left, Me.txtDateF2, CBA_DM2Y, , pbUseCtlDate)
''    Call p_UpdIP(Me.txtDateF2, True)
    Call p_FormatReturn
End Sub
Private Sub txtDateF3_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If g_IsDate(Me.txtDateF3, True) Then
        pbUseCtlDate = True
    Else
        pbUseCtlDate = False
        varCal.sDate = psDate
    End If
    Call CBA_BT_CalendarShow(Me.Top, Me.Left, Me.txtDateF3, CBA_DM2Y, , pbUseCtlDate)
''    Call p_UpdIP(Me.txtDateF3, True)
    Call p_FormatReturn
End Sub
Private Sub txtDateT1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If g_IsDate(Me.txtDateT1, True) Then
        pbUseCtlDate = True
    Else
        pbUseCtlDate = False
        varCal.sDate = psDate
    End If
    Call CBA_BT_CalendarShow(Me.Top, Me.Left, Me.txtDateT1, CBA_DM2Y, , pbUseCtlDate)
''    Call p_UpdIP(Me.txtDateT1, True)
    Call p_FormatReturn
End Sub
Private Sub txtDateT2_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If g_IsDate(Me.txtDateT2, True) Then
        pbUseCtlDate = True
    Else
        pbUseCtlDate = False
        varCal.sDate = psDate
    End If
    Call CBA_BT_CalendarShow(Me.Top, Me.Left, Me.txtDateT2, CBA_DM2Y, , pbUseCtlDate)
''    Call p_UpdIP(Me.txtDateT2, True)
    Call p_FormatReturn
End Sub
Private Sub txtDateT3_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If g_IsDate(Me.txtDateT3, True) Then
        pbUseCtlDate = True
    Else
        pbUseCtlDate = False
        varCal.sDate = psDate
    End If
    Call CBA_BT_CalendarShow(Me.Top, Me.Left, Me.txtDateT3, CBA_DM2Y, , pbUseCtlDate)
''    Call p_UpdIP(Me.txtDateT3, True)
    Call p_FormatReturn
End Sub

Private Sub txtProdSearch_Change()
     pbProductsIP = True
     Call p_lstProdListChange("Change")
End Sub

Private Sub cboCG_Change()
    Dim a As Long
    If pbSetupIP = True Then Exit Sub
        
    pbSetupIP = True
    cboSCG.Clear
    For a = LBound(pCGs, 2) To UBound(pCGs, 2)
        If CStr(pCGs(0, a) & " - " & pCGs(1, a)) = CStr(cboCG.Value) Then
            cboSCG.AddItem pCGs(2, a) & " - " & pCGs(3, a)
        End If
    Next a
    pbSetupIP = False: pbProductsIP = True
    Call p_lstProdListChange("Change")
''    If cboSCG.Value = cboSCG.List(0) Then
''        Call cboSCG_Change
''    Else
''        cboSCG.Value = cboSCG.List(0)
''    End If

End Sub
Private Sub cboSCG_Change()
''    Dim a As Long, b As Long, lAdded As Long
    If pbSetupIP = True Then Exit Sub
    pbProductsIP = True
    Call p_lstProdListChange("Change")
''        ' Add the proper items to the list
''        Me.cboSCG.Clear
''        If Me.cboCG.ListIndex = -1 Then Exit Sub
''        ' For all the CGs, get the SCG and Desc
''        For a = LBound(pCGs, 2) To UBound(pCGs, 2)
''            If Me.cboCG = pCGs(0, a) & " - " & pCGs(1, a) Then
''                Me.cboSCG.AddItem pCGs(3, a) & IIf(NZ(pCGs(4, a), "") > "", " - " & pCGs(4, a), "")
''            End If
''        Next a
''    End If
End Sub

Private Sub txtUnitCost_AfterUpdate()
    Call p_ApplyValues("UnitCost")
End Sub

Private Sub txtSupplierCostSupport_AfterUpdate()
    Call p_ApplyValues("SupplierCostSupport")
End Sub

''Private Sub txtPriorSales_AfterUpdate()
''    Call p_ApplyValues("PriorSales")
''End Sub

''Private Sub txtExpectedSales_AfterUpdate()
''    Call p_ApplyValues("ExpectedSales")
''End Sub

Private Sub txtCalculatedSales_AfterUpdate()
    Call p_ApplyValues("CalculatedSales")
End Sub

Private Sub txtUPSPW_AfterUpdate()
    Call p_ApplyValues("UPSPW")
End Sub

Private Sub txtFillQty_AfterUpdate()
    Call p_ApplyValues("FillQty")
End Sub

Private Sub txtRetailPrice_AfterUpdate()
    Call p_ApplyValues("RetailPrice")
End Sub

''Private Sub txtCurrRetailPrice_Change()
''    Call p_ApplyValues("CurrRetailPrice")
''End Sub

Private Sub txtSalesMultiplier_AfterUpdate()
    Call p_ApplyValues("SalesMultiplier")
End Sub

Private Sub UserForm_Initialize()
    ' Will allow entry of various field format
    Dim aMsg() As String, lIdx As Long, lTop As Long, sMsg As String, a As Long, b As Long, bfound As Boolean, SNo As String
    pbSetupIP = True: pbAllUpd = False: CBA_bDataChg = False '': CBA_lFrmID = 5
    ' Format the input Msg so that we can turn it into an array - should come in the format 'Heading_To_Display 1stParm 2ndParm etc
    ' Pre-format the input parms
    Me.StartUpPosition = 0
    plAuth = CBA_lAuth
    Me.cmdProdAccept.Visible = False
    plProductID = CBA_lProduct_ID
    sMsg = Replace(CBA_strAldiMsg, " / ", "/")
    sMsg = Replace(sMsg, vbCrLf, "~")
    sMsg = Replace(sMsg, " ", "~")
    aMsg = Split(sMsg, "~")
    If Right(aMsg(0), 4) = "Date" Or InStr(1, aMsg(0), "Dates") > 0 Then sMsg = Replace(sMsg, "-", "~")
    sMsg = Replace(sMsg, vbCrLf, "~")
    pbRegions = False: pbRegion1 = False: pbRegionR = False: pbDates = False: pbProducts = False: pbProductCode = False: pbComp = False: pbCann = False
    aMsg = Split(sMsg & "~~~~~~~~~", "~")
    psField = Replace(aMsg(0), "_", " ")
    ' Pick if the form that calls is the Region Entry forms and decide which (it matters as to how it is saved and input into the screen)
    If psField = "Region1" Then
        psForm = "frmProducts"
        psField = "Region"
    ElseIf psField = "RegionRows" Then
        psForm = "frmProductRows"
        psField = "Region"
    End If
    ' Decide on the heading
    If psField = "Region" Then
        Me.lblHdg.Caption = psField & " Data"
    Else
        Me.lblHdg.Caption = psField & " Entry"
    End If
    ' If a 'Date' field then show the dates and make the Regions list invisible
    If Right(psField, 4) = "Date" Or InStr(1, aMsg(0), "Dates") > 0 Then
        ' Hide the formatted controls                                                           ' Note spaces can be CR
        Me.Width = 342
        Me.Height = 400
        Me.frmReturn.Top = 228
        Me.frmReturn.Left = 10
        Me.cmdTBA.Enabled = False
        Me.lblApply.Visible = True
        ' Save the default date
        psDate = varCal.sDate
        ' Set the dates that were priorly selected
        If aMsg(1) <> "N/A" Then
            Me.txtDateF1 = g_FixDate(aMsg(1), CBA_DM2Y)
            Me.txtDateT1 = g_FixDate(aMsg(2), CBA_DM2Y)
            Me.txtDateF2 = g_FixDate(aMsg(3), CBA_DM2Y)
            Me.txtDateT2 = g_FixDate(aMsg(4), CBA_DM2Y)
            Me.txtDateF3 = g_FixDate(aMsg(5), CBA_DM2Y)
            Me.txtDateT3 = g_FixDate(aMsg(6), CBA_DM2Y)
        End If
        Me.frmeDate1.Visible = True
        Me.frmeDate2.Visible = True
        Me.frmeDate3.Visible = True
        pbDates = True
    ElseIf InStr(1, psField, "Regions") > 0 Then ' If the 'Region' field then make the Regions list Visible, and make it centre screen
        GoSub GSRegion
        ' Hide the formatted controls                                                           ' Note spaces can be CR
        Me.Width = 342
''        Me.frmReturn.Top = 278
''        Me.frmReturn.Left = 1
        ' Set the visibility and position
        Me.frmeRegions.Visible = True
        Me.lblApply.Visible = False
        Me.frmeRegions.Caption = "Select Regions"
        lTop = 48
        For lIdx = 1 To 8
            SNo = Format(lIdx, "00")
            Me("chk" & SNo & "Regions").Top = lTop
            lTop = lTop + 28
        Next
        Me.frmeRegions.Left = 120
        Me.frmeRegions.Width = 100
        Me.frmeRegions.Height = 274
        Me.frmReturn.Top = 366
        Me.frmReturn.Left = 1
        Me.Height = 520
        
        sMsg = aMsg(1)
        
        pbRegions = True
        Me.cmdTBA.Enabled = False
        ' Set up the regions
        paReg = Split(CREGS, ",")
        ' Set the Regions that were priorly selected
        For lIdx = 1 To 8
            SNo = Format(lIdx, "00")
            If UCase(sMsg) = "N/A" Or UCase(sMsg) = "ALL" Or InStr(1, sMsg, paReg(lIdx - 1)) > 0 Then
                Me("chk" & SNo & "Regions").Value = True
            Else
                Me("chk" & SNo & "Regions").Value = False
            End If
        Next
        Call p_FormatReturn

    ElseIf InStr(1, psField, "Region") > 0 And psForm = "frmProducts" Then ' If the 'Region' field then make the Regions list Visible, and make it centre screen
        pbRegion1 = True
        GoSub GSRegion
        ' Get the records into the class module
        Call p_GetProdRegionClass("Read")
        Call g_SetupIP(FORMTAG, 1, False)
        
    ElseIf InStr(1, psField, "Region") > 0 And psForm = "frmProductRows" Then ' If the 'Region' field then make the Regions list Visible, and make it centre screen
        pbRegionR = True
        GoSub GSRegion
        ' Get the records into the class module
        Call p_UpdProdRegionRecs("Read")
        Call g_SetupIP(FORMTAG, 1, False)
    ElseIf InStr(1, psField, "Compl") > 0 Then      ' If the 'Complementary' field then make the lstProducts list Visible, and make it centre screen
        pbComp = True
        ''Me.cmdNA.Visible = False
        GoSub GSExistingProducts
    ElseIf InStr(1, psField, "Canni") > 0 Then      ' If the 'Cannibalised' field then make the lstProducts list Visible, and make it centre screen
        pbCann = True
        ''Me.cmdNA.Visible = False
        GoSub GSExistingProducts
    ElseIf InStr(1, psField, "Product") > 0 Then    ' If the 'Product' field then make the lstProducts list Visible, and make it centre screen - single select
        pbProductCode = True
        Me.lstProducts.MultiSelect = fmMultiSelectSingle
        GoSub GSExistingProducts
    End If
    ' Position the form as per the last
    Me.Top = g_PosForm(0, 0, 0, "Top")
    Me.Left = g_PosForm(0, Me.Width, 0, "Left")
    
    ' Format the data input into the field at the bottom
    Call p_FormatReturn
    pbSetupIP = False
    Exit Sub
    
GSExistingProducts:
    pbProducts = True
    ' Hide the formatted controls                                                           ' Note spaces can be CR
    If pbProductCode = False Then
        Me.cmdTBA.Enabled = False
    Else
        Me.lblBestFit.Visible = True
        Me.cmdTBA.Visible = False
        Me.cmdNA.Visible = False
    End If
    ' Add the existing products to the list
    If aMsg(1) <> "N/A" Then
        For lIdx = 1 To UBound(aMsg, 1)
            If NZ(aMsg(lIdx), "") > "" Then
                Me.lstProducts.AddItem Replace(aMsg(lIdx), "_", " ")
            End If
        Next
        ' Select the existing products
        For lIdx = 0 To Me.lstProducts.ListCount - 1
            Me.lstProducts.Selected(lIdx) = True
        Next
    End If
    ' Call the sql to fill the CG and SCG combo boxes
    CBA_COM_SQLQueries.CBA_COM_GenPullSQL "CGSCGList"
    pCGs = CBA_CBISarr
    cboCG.Clear
    For a = LBound(pCGs, 2) To UBound(pCGs, 2)
        bfound = False
        ' For all the CGs (note: the SCG are inc in the array so there maybe many entries per CG) ...
        For b = 0 To cboCG.ListCount - 1
            If cboCG.List(b) = pCGs(0, a) & " - " & pCGs(1, a) Then
                bfound = True
                Exit For
            End If
        Next b
        ' Only Add the CG if it doesn't exist
        If bfound = False Then                      'And pCGs(0, a) <> 2 And pCGs(0, a) <> 55 Then
            cboCG.AddItem pCGs(0, a) & IIf(NZ(pCGs(1, a), "") > "", " - " & pCGs(1, a), "")
        End If
    Next a
    ' Set the visibility and position
    Me.Width = 355
    Me.Height = 454
    Me.txtReturn.Left = 60
    Me.txtReturn.Width = 170
    Me.frmReturn.Top = 294
    Me.frmReturn.Left = 10

    Me.txtReturn.TextAlign = fmTextAlignLeft
    Me.frmeProducts.Visible = True
    Me.frmeProducts.Left = 35
    Me.frmeProducts.Height = 208
Return

GSRegion:
        ' To stop prior updating of upd fields, set up is on
        Call g_SetupIP(FORMTAG, 1, True, True)
        ' Show the formatted controls                                                           ' Note spaces can be CR
        Me.Width = 577
        Me.Height = 492
        Me.frmeRegions.Visible = True
        Me.frmeRegions.Top = 50
        Me.frmeRegions.Left = 2
        Me.frmeRegions.Width = 565
        Me.frmReturn.Top = 300
        Me.frmReturn.Left = 100
        Me.Image1.Left = 440
        Me.lblHdg.Left = 110
        sMsg = aMsg(1)
        ' Set up the regions
        paReg = Split(CREGS, ",")
        CBA_lFrmID = plFrmID
        For lIdx = 1 To 8
            Set ptProdRows(lIdx) = New CBA_AST_clsProdRows
            Call ptProdRows(lIdx).FormInit(Me.Controls, lIdx, Me.Top, Me.Left, FORMTAG)
        Next
        Me.cmdTBA.Enabled = False
        Me.frmeNA.Visible = False
Return
End Sub

Private Sub p_UpdFld(sField As String)
    If pbRegion1 Or pbRegionR Then
        pbSetupIP = True
        Me("txt" & Mid(sField, 4, 2) & "Upd") = "Y"
        Me.cmdSave.Visible = True
        pbSetupIP = False
    End If
End Sub

Private Sub p_ApplyValues(sField As String)
    Dim lIdx As Long, SNo As String, sFmt As String
    'Dim lPrior As Long, lCalced As Long, lUPSPW As Long, sngMult As Single, lWeeks As Long, sRegion As String
    ''Dim Mult_Calc_USW As String
    Static bRecused As Boolean
    If bRecused Then Exit Sub
    If plAuth = 0 Then Exit Sub
    If Trim(NZ(Me("txt" & sField).Value, "")) = "" Then Exit Sub
    bRecused = True
  
    ' Apply the values from the AllField to the Rows
    For lIdx = 1 To 8
        SNo = Format(lIdx, "00")
        If g_IsNumeric(Me("txt" & sField)) = False Then GoTo Exit_AV
''        Me("txt" & sNo & sField) = Format(Me("txt" & sField), sFmt)
        If p_CalcRegionValues(Me("txt" & sField), SNo, sField, False) = True Then
            ' Routine will process
        Else
            ' Get the format...
            If sFmt = "" Then sFmt = AST_FillTagArrays(sField, plFrmID, plAuth, "Format")
            Me("txt" & SNo & sField) = Format(Me("txt" & sField), sFmt)
        End If
    Next
    Me.cmdSave.Visible = True: pbAllUpd = True
Exit_AV:
    Me("txt" & sField) = "": bRecused = False
End Sub

Private Function p_FormatReturn(Optional sNum As String = "00", Optional sNA As String = "") As Boolean
    Dim lIdx As Long, sField As String, sSuff As String, sPref As String, sMsg As String, sLast As String, sName As String, SNo As String, bVisTrue As Boolean, lPLen As Long
    Dim bLock As Boolean, bChkLock As Boolean, sFmt As String, sChkFmt As String, sLock As String, sChkLock As String
    ' Format the data input into the 'return field' to see how it looks before return
    p_FormatReturn = True
    sSuff = ""
    
    ' If Regions
    If pbRegions Then
        DoEvents
        For lIdx = 1 To 8
            SNo = Format(lIdx, "00"): bVisTrue = False
            If (Me("chk" & SNo & "Regions").Value = True And sNum <> SNo) Or (Me("chk" & SNo & "Regions").Value = False And sNum = SNo) Then
                bVisTrue = True
                sField = sField & sSuff & paReg(lIdx - 1)
                sSuff = " / "
            End If
        Next
        If sSuff = "" Then sField = "N/A"
    ElseIf pbRegion1 Or pbRegionR Then
        DoEvents
        ' Test whether the check boxes are valid to change
        bChkLock = (AST_FillTagArrays(Me("chk01Regions").Tag, plFrmID, plAuth, "Lock") = "lock")
        ' Format the 'All Entry' fields as being unlocked or invisible
        If bChkLock = True Then
            Call p_FormatFlds(Me("txtUnitCost"), IIf(bChkLock, "True", "False"))
            Call p_FormatFlds(Me("txtSupplierCostSupport"), IIf(bChkLock, "True", "False"))
''            Call p_FormatFlds(Me("txtPriorSales"), "False")
''            Call p_FormatFlds(Me("txtExpectedSales"), "False")
            Me("txtPriorSales").Visible = False
            Me("txtExpectedSales").Visible = False
            Call p_FormatFlds(Me("txtCalculatedSales"), IIf(bChkLock, "True", "False"))
            Call p_FormatFlds(Me("txtUPSPW"), IIf(bChkLock, "True", "False"))
            Call p_FormatFlds(Me("txtFillQty"), IIf(bChkLock, "True", "False"))
            'Call p_FormatFlds(Me("txtCurrRetailPrice"), IIf(bChkLock, "True", "False"))
            Me("txtCurrRetailPrice").Visible = False
            Call p_FormatFlds(Me("txtRetailPrice"), IIf(bChkLock, "True", "False"))
            Call p_FormatFlds(Me("txtSalesMultiplier"), IIf(bChkLock, "True", "False"))
        Else
            Me("txtUnitCost").Visible = Not bChkLock
            Me("txtSupplierCostSupport").Visible = Not bChkLock
'            Me("txtPriorSales").Visible = Not bChkLock
'            Me("txtExpectedSales").Visible = Not bChkLock
            Me("txtCalculatedSales").Visible = Not bChkLock
            Me("txtUPSPW").Visible = Not bChkLock
            Me("txtFillQty").Visible = Not bChkLock
'            Me("txtCurrRetailPrice").Visible = Not bChkLock
            Me("txtRetailPrice").Visible = Not bChkLock
            Me("txtSalesMultiplier").Visible = Not bChkLock
        End If

        ' Format the other fields as being unlocked or invisible
        For lIdx = 1 To 8
            SNo = Format(lIdx, "00"): bVisTrue = False
            If (Me("chk" & SNo & "Regions").Value = True And sNum <> SNo) Or (Me("chk" & SNo & "Regions").Value = False And sNum = SNo) Then
                bVisTrue = True
                sField = sField & sSuff & paReg(lIdx - 1)
                sSuff = " / "
                Me("chk" & SNo & "Regions").Visible = bVisTrue
                Me("chk" & SNo & "Regions").Locked = bChkLock
            End If
            If bVisTrue = True Then
                Call p_FormatFlds(Me("txt" & SNo & "UnitCost"), IIf(bChkLock, "True", "False"))
                Call p_FormatFlds(Me("txt" & SNo & "SupplierCostSupport"), IIf(bChkLock, "True", "False"))
                Call p_FormatFlds(Me("txt" & SNo & "PriorSales"), IIf(bChkLock, "True", "False"))
                Call p_FormatFlds(Me("txt" & SNo & "ExpectedSales"), IIf(bChkLock, "True", "False"))
                Call p_FormatFlds(Me("txt" & SNo & "CalculatedSales"), IIf(bChkLock, "True", "False"))
                Call p_FormatFlds(Me("txt" & SNo & "UPSPW"), IIf(bChkLock, "True", "False"))
                Call p_FormatFlds(Me("txt" & SNo & "FillQty"), IIf(bChkLock, "True", "False"))
                Call p_FormatFlds(Me("txt" & SNo & "CurrRetailPrice"), IIf(bChkLock, "True", "False"))
                Call p_FormatFlds(Me("txt" & SNo & "RetailPrice"), IIf(bChkLock, "True", "False"))
                Call p_FormatFlds(Me("txt" & SNo & "SalesMultiplier"), IIf(bChkLock, "True", "False"))
            End If
            Me("txt" & SNo & "UnitCost").Visible = bVisTrue
            Me("txt" & SNo & "SupplierCostSupport").Visible = bVisTrue
            Me("txt" & SNo & "PriorSales").Visible = bVisTrue
            Me("txt" & SNo & "ExpectedSales").Visible = bVisTrue
            Me("txt" & SNo & "CalculatedSales").Visible = bVisTrue
            Me("txt" & SNo & "UPSPW").Visible = bVisTrue
            Me("txt" & SNo & "FillQty").Visible = bVisTrue
            Me("txt" & SNo & "CurrRetailPrice").Visible = bVisTrue
            Me("txt" & SNo & "RetailPrice").Visible = bVisTrue
            Me("txt" & SNo & "SalesMultiplier").Visible = bVisTrue
        Next
        If sSuff = "" Then sField = "N/A"
    ElseIf pbDates Then ' If Dates
        sSuff = "": sMsg = "": sPref = "": sLast = "01/01/1900"
        For lIdx = 1 To 3
            If g_IsDate(Me("txtDateF" & lIdx), True) And g_IsDate(Me("txtDateT" & lIdx), True) Then
                If CDate(g_FixDate(Me("txtDateF" & lIdx))) > CDate(g_FixDate(Me("txtDateT" & lIdx))) Then
                    sMsg = sMsg & sPref & "'From Date " & lIdx & "' can't be greater than 'To Date " & lIdx & "'"
                    sPref = vbCrLf
                    p_FormatReturn = False
                ElseIf CDate(g_FixDate(sLast)) > CDate(g_FixDate(Me("txtDateT" & lIdx))) Then
                    sMsg = sMsg & sPref & "'" & sName & "' can't be greater than 'To Date " & lIdx & "'"
                    sPref = vbCrLf
                    p_FormatReturn = False
                End If
                sField = sField & sSuff & g_FixDate(Me("txtDateF" & lIdx), CBA_DM2Y) & "-" & g_FixDate(Me("txtDateT" & lIdx), CBA_DM2Y)
                sSuff = " "
                sName = "'To Date " & lIdx & "'": sLast = g_FixDate(Me("txtDateT" & lIdx))
            ElseIf g_IsDate(Me("txtDateF" & lIdx), True) And g_IsDate(Me("txtDateT" & lIdx), True) = False Then
                sField = sField & sSuff & g_FixDate(Me("txtDateF" & lIdx), CBA_DM2Y) & "-"
                sSuff = " "
                sMsg = sMsg & sPref & "No 'To Date " & lIdx & "' has been entered"
                sPref = vbCrLf
                p_FormatReturn = False
            ElseIf g_IsDate(Me("txtDateF" & lIdx), True) = False And g_IsDate(Me("txtDateT" & lIdx), True) Then
                sField = sField & sSuff & "-" & g_FixDate(Me("txtDateT" & lIdx), CBA_DM2Y)
                sSuff = " "
                sMsg = sMsg & sPref & "No 'From Date " & lIdx & "' has been entered"
                sPref = vbCrLf
                p_FormatReturn = False
            ElseIf g_IsDate(Me("txtDateF" & lIdx), True) = False And g_IsDate(Me("txtDateT" & lIdx), True) = False Then
                sField = sField & sSuff
                sSuff = " "
            End If
        Next
        sField = Trim(sField)
        If sField = "" Then sField = "N/A"
    ElseIf pbProducts = True Then
        sSuff = "": sField = "": lPLen = PRODLEN
        ' Set the len of the text b4 the CR
        If pbCann Then
            lPLen = CANNLEN
        ElseIf pbComp Then
            lPLen = COMPLEN
        End If
        If Me.lstProducts.ListCount > -1 And sNA = "" Then
            ' Add all the selected Products that have been selected
            For lIdx = 0 To Me.lstProducts.ListCount - 1
                If Me.lstProducts.Selected(lIdx) = True Then
                    sField = sField & sSuff & Left(Me.lstProducts.List(lIdx), lPLen)
                    sSuff = vbCrLf
                End If
            Next
        Else
            sField = "N/A"
        End If
    End If
    ' Set up the Returns field
    Me.txtReturn.Value = sField
    If pbSetupIP = False Then Me.cmdSave.Visible = True
    ' If text is too long...
'    If Len(sField) > 255 And p_FormatReturn = True Then
'        p_FormatReturn = False
'        sMsg = "Number of characters in text should be less than 256 - please unselect item/s "
'    End If
    ' Set up the errors or not...
    If p_FormatReturn = True Then
        Me.txtReturn.BackColor = CBA_Grey
        Me.txtMsg.Visible = False
    Else
        Me.txtReturn.BackColor = CBA_Red
        Me.txtMsg.Visible = True
        Me.txtMsg = sMsg
    End If
    Exit Function
    
End Function

Private Sub p_lstProdListChange(sSearch_Change As String)
    ' This routine will further format the lstProduct list to include items that haven't yet been selected
    Dim sSQL As String, bPassOK As Boolean, sAnd As String, lIdx As Long, a As Long, b As Long, sSearch As String
    Dim laList() As Long, lListIdx As Long, vTemp
    Static bFlag As Boolean, lProdCode As Long
    If pbSetupIP = False Then
        If lProdCode = 0 Then
            lProdCode = g_DLookup("PD_Product_Code", "L2_Products", "PD_ID=" & plProductID, "PD_ID", g_GetDB("ASYST"), -1)
        End If
        If sSearch_Change = "Search" Then
            If Trim(NZ(Me.txtProdSearch, "")) = "" And Me.cboCG.ListIndex = -1 Then
                MsgBox "At least one selection must be made"
                Exit Sub
            End If
            ' Remove all of the Products that haven't been selected
            lListIdx = -1
            ReDim laList(0 To 0)
            
            lIdx = Me.lstProducts.ListCount - 1
            While lIdx >= 0
                If Me.lstProducts.Selected(lIdx) = False Then
                    Me.lstProducts.RemoveItem (lIdx)                      '''(lIdx)
                Else
                    lListIdx = lListIdx + 1
                    ReDim laList(0 To lListIdx)
                    laList(lListIdx) = Val(Me.lstProducts.Column(0, lIdx))
                End If
                lIdx = lIdx - 1
            Wend
            ' Reset the products
            Me.cmdProdAccept.Caption = "Search Selection"
            Me.lstProducts.Locked = True
            ''Me.lstProducts.BackColor = CBA_Grey

''        ElseIf bFlag = True And pbProductsIP = False Then
''            Me.cmdProdAccept.Caption = "Accept Selection"
''            Me.lstProducts.Locked = False
''            Me.lstProducts.BackColor = CBA_White
''        End If
''        ' Wait for the Apply Selection to run the following
''        If pbProductsIP = False Then
            Me.cmdProdAccept.Caption = "Accept Selection"
            Me.lstProducts.Locked = False
            ''Me.lstProducts.BackColor = CBA_White
            ' Set up the SQL for the current selected query
            sAnd = " AND ": sSearch = LCase(Me.txtProdSearch)
            sSQL = "SELECT ProductCode,Description,ProductClass " & _
                "FROM cbis599p.dbo.Product " & _
                "WHERE Con_ProductCode IS NULL "
            If NZ(Me.txtProdSearch, "") > "" Then
                sSQL = sSQL & sAnd & "( (CHARINDEX('" & sSearch & "', ProductCode)>0) OR (CHARINDEX('" & sSearch & "', lower(Description))>0))"
                sAnd = " AND "
            End If
            If NZ(Me.cboCG, "") > "" Then
                sSQL = sSQL & sAnd & "CGNo=" & Left(Me.cboCG, 2) & " "
                sAnd = " AND "
            End If
            If NZ(Me.cboSCG, "") > "" Then
                sSQL = sSQL & sAnd & "SCGNo=" & Left(Me.cboSCG, 2) & " "
                sAnd = " AND "
            End If
            If sAnd = "" Then Exit Sub 'If nothing entered then exit
            CBA_DBtoQuery = 599
            bPassOK = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", "CBIS_QUERY", CBA_BasicFunctions.TranslateServerName(CBA_DBtoQuery, Date), "SQLNCLI10", sSQL, 120, , , False)
            ' Append the results of the query into the Products Listbox
            If bPassOK = True Then
                ' Do a sort of the array
                For a = 0 To UBound(CBA_CBISarr, 2) - 1
                    For b = a + 1 To UBound(CBA_CBISarr, 2)
                        ' If the sort is in the wrong place then swap it over
                        If (CBA_CBISarr(0, a) > CBA_CBISarr(0, b)) Then
                            ' Swap the ProductCode values
                             vTemp = NZ(CBA_CBISarr(0, b), 0)
                             CBA_CBISarr(0, b) = NZ(CBA_CBISarr(0, a), 0)
                             CBA_CBISarr(0, a) = vTemp
                            ' Swap the Description values
                             vTemp = NZ(CBA_CBISarr(1, b), "?")
                             CBA_CBISarr(1, b) = NZ(CBA_CBISarr(1, a), "?")
                             CBA_CBISarr(1, a) = vTemp
                        End If
                    Next b
                Next a
                ' Get rid of any items that have already been selected
                For a = 0 To UBound(CBA_CBISarr, 2)
                    ' If the item is the selected Product Code then delete it
                    If (CBA_CBISarr(0, a) = lProdCode) Then
                        CBA_CBISarr(0, a) = 0
                    Else
                        For b = 0 To lListIdx
                            ' If the item is already in the list, delete it
                            If (CBA_CBISarr(0, a) = laList(b)) Then
                                 CBA_CBISarr(0, b) = 0
                            End If
                        Next b
                    End If
                Next a
                ' Add the rest of the items to the list
                For b = LBound(CBA_CBISarr, 2) To UBound(CBA_CBISarr, 2)
                    If Not NZ(CBA_CBISarr(0, b), 0) = 0 Then
                        Me.lstProducts.AddItem CBA_CBISarr(0, b) & "-" & CBA_CBISarr(1, b)
                    End If
                Next
            Else
                MsgBox "Search data not found"
            End If
            Me.txtProdSearch = ""
            Me.cboCG = Null
            Me.cboSCG = Null
        Else
            Me.cmdProdAccept.Visible = True
        End If
        bFlag = pbProductsIP
    End If

End Sub

Private Function p_GetProdRegionClass(Read_Write As String) As String
    ' Get the POS for each region (Read_Write="Read") and update the records in the class for each region (Read_Write="Write")
    Dim tUnitCost As Scripting.Dictionary
    Dim tPrice As Scripting.Dictionary
    Dim tMult As Scripting.Dictionary
    Dim SNo As String, lIdx As Long, lRegion As Long, sSQL As String
    Dim CN As ADODB.Connection, RS As ADODB.Recordset
    Dim cls_Mod As CBA_AST_Product
    Dim DivNo As Long, tTotSto As Long, MaxNo As Long
    Dim tCalcSales As Currency, tSupCost As Currency
    Dim curVal As Variant
    Dim oKey
    
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    ' Read from the class
    If Read_Write = "Read" Then
        If CBA_ASyst.CBA_AST_getProdClassMod(cls_Mod) = True Then
            CBA_ErrTag = ""
            For lRegion = 501 To 509
                If lRegion <> 508 Then
                    SNo = Format(lRegion - 500, "00")
                    If SNo = "09" Then SNo = "08"
                    Me("txt" & SNo & "ID").Value = NZ(cls_Mod.plPDID, 0)
                    Me("txt" & SNo & "UnitCost").Value = Format(NZ(cls_Mod.pUnit_CostDiv(lRegion), 0), "0.00")
                    Me("txt" & SNo & "SupplierCostSupport").Value = Format(NZ(cls_Mod.pSupplier_Cost_SupportDiv(lRegion), 0), "0.00")
                    Me("txt" & SNo & "PriorSales").Value = Format(NZ(cls_Mod.pcPrior_SalesDiv(lRegion), 0), "#,0")
                    Me("txt" & SNo & "ExpectedSales").Value = Format(NZ(cls_Mod.pcEstSalesDiv(lRegion), 0), "#,0")
                    Me("txt" & SNo & "CalculatedSales").Value = Format(NZ(cls_Mod.pCalcSalesDiv(lRegion), 0), "#,0")
                    Me("txt" & SNo & "UPSPW").Value = Format(NZ(cls_Mod.pcUPSPWDiv(lRegion), 0), "#,0")
                    Me("txt" & SNo & "FillQty").Value = Format(NZ(cls_Mod.pFill_QtyDiv(lRegion), 0), "#,0")
                    Me("txt" & SNo & "CurrRetailPrice").Value = Format(NZ(cls_Mod.pCurr_Retail_PriceDiv(lRegion), 0), "0.00")
                    Me("txt" & SNo & "RetailPrice").Value = Format(NZ(cls_Mod.pRetail_PriceDiv(lRegion), 0), "0.00")
                    Me("txt" & SNo & "SalesMultiplier").Value = Format(NZ(cls_Mod.pEstMultiplierDiv(lRegion), 0), "0.00")
                    Me("txt" & SNo & "Status").Value = cls_Mod.pStatusDiv(lRegion)
                    Me("txt" & SNo & "Upd").Value = "N"
                    If Me("txt" & SNo & "Status").Value < 4 Then
                        Me("chk" & SNo & "Regions").Value = True
                    Else
                        Me("chk" & SNo & "Regions").Value = False
                    End If
                End If
            Next
        Else
            Err.Raise 513, , "No cls_mod (CBA_AST_Product) Module found"
        End If
    Else            ' Update the class
        If CBA_ASyst.CBA_AST_getProdClassMod(cls_Mod) = True Then
            For lIdx = 1 To 8
                SNo = Format(lIdx, "00")
                If lIdx < 8 Then DivNo = 500 + lIdx Else DivNo = 509
                If UCase(Me("txt" & SNo & "Upd").Value) = "Y" Or (pbAllUpd = True And Me("txt" & SNo & "UnitCost").Visible = True) Then
                    cls_Mod.SetpCalcSalesDiv g_UnFmt(Me("txt" & SNo & "CalculatedSales").Value, "sng"), DivNo
                    cls_Mod.SetpUnit_CostDiv g_UnFmt(Me("txt" & SNo & "UnitCost").Value, "cur"), DivNo
                    cls_Mod.SetpSupplier_Cost_SupportDiv g_UnFmt(Me("txt" & SNo & "SupplierCostSupport").Value, "cur"), DivNo
                    cls_Mod.SetpcPrior_SalesDiv g_UnFmt(Me("txt" & SNo & "PriorSales").Value, "sng"), DivNo
                    cls_Mod.SetpcEstSalesDiv g_UnFmt(Me("txt" & SNo & "ExpectedSales").Value, "sng"), DivNo
                    cls_Mod.SetpcUPSPWDiv g_UnFmt(Me("txt" & SNo & "UPSPW").Value, "lng"), DivNo
                    cls_Mod.SetpFill_QtyDiv g_UnFmt(Me("txt" & SNo & "FillQty").Value, "lng"), DivNo
                    cls_Mod.SetpCurr_Retail_PriceDiv g_UnFmt(Me("txt" & SNo & "CurrRetailPrice").Value, "cur"), DivNo
                    cls_Mod.SetpRetail_PriceDiv g_UnFmt(Me("txt" & SNo & "RetailPrice").Value, "cur"), DivNo
                    cls_Mod.SetpEstMultiplierDiv g_UnFmt(Me("txt" & SNo & "SalesMultiplier").Value, "sng"), DivNo
                    If Me("chk" & SNo & "Regions").Value = True Then
                        cls_Mod.SetpStatusDiv 1, DivNo
                    Else
                        cls_Mod.SetpStatusDiv 4, DivNo
                    End If
                 End If
            Next
            Set tUnitCost = New Scripting.Dictionary
            Set tPrice = New Scripting.Dictionary
            Set tMult = New Scripting.Dictionary
            For DivNo = 501 To 509
                If DivNo = 508 Then DivNo = 509
                    tCalcSales = tCalcSales + cls_Mod.pCalcSalesDiv(DivNo)
                    tTotSto = tTotSto + cls_Mod.pcaCurRegStoNoDiv(DivNo)
                    tSupCost = tSupCost + (cls_Mod.pCalcSalesDiv(DivNo) * cls_Mod.pSupplier_Cost_SupportDiv(DivNo))
                    
                    If tPrice.Exists(cls_Mod.pRetail_PriceDiv(DivNo)) Then tPrice(cls_Mod.pRetail_PriceDiv(DivNo)) = tPrice(cls_Mod.pRetail_PriceDiv(DivNo)) + 1 Else tPrice(cls_Mod.pRetail_PriceDiv(DivNo)) = 1
                    If tUnitCost.Exists(cls_Mod.pUnit_CostDiv(DivNo)) Then tUnitCost(cls_Mod.pUnit_CostDiv(DivNo)) = tUnitCost(cls_Mod.pUnit_CostDiv(DivNo)) + 1 Else tUnitCost(cls_Mod.pUnit_CostDiv(DivNo)) = 1
            Next
            cls_Mod.SetpUPSPW tCalcSales / tTotSto / cls_Mod.pWoS
            cls_Mod.SetpOrigCalcSales tCalcSales
            If tCalcSales = 0 Then
                cls_Mod.SetpSupplier_Cost_Support 0
            Else
                cls_Mod.SetpSupplier_Cost_Support tSupCost / tCalcSales
            End If
            curVal = "": MaxNo = 0
            For Each oKey In tPrice.Keys
                If tPrice(oKey) > MaxNo Or curVal = "" Then
                    MaxNo = tPrice(oKey)
                    curVal = oKey
                End If
            Next oKey
            cls_Mod.SetpRetail_Price curVal
            curVal = "": MaxNo = 0
            For Each oKey In tUnitCost.Keys
                If tUnitCost(oKey) > MaxNo Or curVal = "" Then
                    MaxNo = tUnitCost(oKey)
                    curVal = oKey
                End If
            Next oKey
            cls_Mod.SetpUnit_Cost curVal
            If NZ(cls_Mod.pExp_Sales, 0) = 0 Then
                cls_Mod.SetpEstMultiplier 1
            Else
                cls_Mod.SetpEstMultiplier tCalcSales / cls_Mod.pExp_Sales
            End If
        Else
            MsgBox "Data not Saved!"
            Err.Raise 513, , "No cls_mod (CBA_AST_Product) Module found"
        End If
    End If
Exit_Routine:
    On Error Resume Next
    Set RS = Nothing
    Set CN = Nothing
    Exit Function
Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-p_GetProdRegionClass", 3)
    CBA_Error = " Error -" & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & sSQL
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "ASYST", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
    
End Function

Private Function p_UpdProdRegionRecs(Read_Write As String) As String
    ' Get the POS for each region (Read_Write="Read") and update the records in the database for each region (Read_Write="Write")
    
    Dim SNo As String, lIdx As Long, lRegion As Long, sSQL As String
    Dim CN As ADODB.Connection, RS As ADODB.Recordset
    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    Set CN = New ADODB.Connection
    Set RS = New ADODB.Recordset
    CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("ASYST") & ";"
    If Read_Write = "Read" Then             ' If a read from the database
        ' Get the products concerned...
        CBA_ErrTag = "SQL"
        sSQL = "SELECT * FROM L3_ProductRegions WHERE PV_PD_ID=" & plProductID & " ORDER BY PV_Region"
        RS.Open sSQL, CN
        Do While Not RS.EOF
            CBA_ErrTag = ""
            lRegion = RS!PV_Region
            SNo = Format(lRegion - 500, "00")
            If SNo = "09" Then SNo = "08"
            Me("txt" & SNo & "ID").Value = NZ(RS!PV_ID, 0)
            Me("txt" & SNo & "UnitCost").Value = Format(NZ(RS!PV_Unit_Cost, 0), "0.00")
            Me("txt" & SNo & "SupplierCostSupport").Value = Format(NZ(RS!PV_Supplier_Cost_Support, 0), "0.00")
            Me("txt" & SNo & "PriorSales").Value = Format(NZ(RS!PV_Prior_Sales, 0), "#,0")
            Me("txt" & SNo & "ExpectedSales").Value = Format(NZ(RS!PV_OrigEstSales, 0), "#,0")
            Me("txt" & SNo & "CalculatedSales").Value = Format(NZ(RS!PV_OrigCalcSales, 0), "#,0")
            Me("txt" & SNo & "UPSPW").Value = Format(NZ(RS!PV_UPSPW, 0), "#,0")
            Me("txt" & SNo & "FillQty").Value = Format(NZ(RS!PV_Fill_Qty, 0), "#,0")
            Me("txt" & SNo & "CurrRetailPrice").Value = Format(NZ(RS!PV_Curr_Retail_Price, 0), "0.00")
            Me("txt" & SNo & "RetailPrice").Value = Format(NZ(RS!PV_Retail_Price, 0), "0.00")
            Me("txt" & SNo & "SalesMultiplier").Value = Format(NZ(RS!PV_OrigEstMultiplier, 0), "0.00")
            Me("txt" & SNo & "Status").Value = NZ(RS!PV_Status, 1)
            Me("txt" & SNo & "Upd").Value = "N"
            If NZ(RS!PV_Status, 1) < 4 Then
                Me("chk" & SNo & "Regions").Value = True
            Else
                Me("chk" & SNo & "Regions").Value = False
            End If
            RS.MoveNext
        Loop
    Else                                   ' If a write to the database
        For lIdx = 1 To 8
            SNo = Format(lIdx, "00")
            If UCase(Me("txt" & SNo & "Upd").Value) = "Y" Or (pbAllUpd = True And Me("txt" & SNo & "UnitCost").Visible = True) Then
                sSQL = ""
                sSQL = sSQL & "UPDATE L3_ProductRegions SET "
                sSQL = sSQL & "PV_Unit_Cost = " & g_UnFmt(Me("txt" & SNo & "UnitCost").Value, "cur") & ","
                sSQL = sSQL & "PV_Supplier_Cost_Support = " & g_UnFmt(Me("txt" & SNo & "SupplierCostSupport").Value, "cur") & ","
                sSQL = sSQL & "PV_Prior_Sales = " & g_UnFmt(Me("txt" & SNo & "PriorSales").Value, "sng") & ","
                sSQL = sSQL & "PV_OrigEstSales = " & g_UnFmt(Me("txt" & SNo & "ExpectedSales").Value, "sng") & ","
                sSQL = sSQL & "PV_OrigCalcSales = " & g_UnFmt(Me("txt" & SNo & "CalculatedSales").Value, "sng") & ","
                sSQL = sSQL & "PV_UPSPW = " & g_UnFmt(Me("txt" & SNo & "UPSPW").Value, "lng") & ","
                sSQL = sSQL & "PV_Fill_Qty = " & g_UnFmt(Me("txt" & SNo & "FillQty").Value, "lng") & ","
                sSQL = sSQL & "PV_Curr_Retail_Price = " & g_UnFmt(Me("txt" & SNo & "CurrRetailPrice").Value, "cur") & ","
                sSQL = sSQL & "PV_Retail_Price = " & g_UnFmt(Me("txt" & SNo & "RetailPrice").Value, "cur") & ","
                sSQL = sSQL & "PV_OrigEstMultiplier = " & g_UnFmt(Me("txt" & SNo & "SalesMultiplier").Value, "sng") & ","
                sSQL = sSQL & "PV_LastUpd = " & g_GetSQLDate(Now(), CBA_DMYHN) & ","
                If Me("chk" & SNo & "Regions").Value = True Then
                    sSQL = sSQL & "PV_Status = 1"
                Else
                    sSQL = sSQL & "PV_Status = 4"
                End If
                sSQL = sSQL & " WHERE (((PV_ID)=" & Me("txt" & SNo & "ID").Value & " ));"
                RS.Open sSQL, CN
            End If
        Next
    End If
Exit_Routine:
    On Error Resume Next
    Set RS = Nothing
    Set CN = Nothing
    Exit Function
Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-p_UpdProdRegionRecs", 3)
    CBA_Error = " Error -" & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & sSQL
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "ASYST", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
    
End Function


Private Sub p_FormatFlds(ctl As MSForms.TextBox, Optional sOR As String = "True")
    ' Will format the sub fld as per locked etc
    If sOR = "True" Then
        ctl.Locked = True
    Else
        ctl.Locked = (AST_FillTagArrays(ctl.Tag, plFrmID, plAuth, "Lock") = "lock")
    End If
    If ctl.Locked = True Then
        ctl.BackColor = CBA_Grey
    Else
        ctl.BackColor = CBA_EntryYellow
    End If
End Sub

Public Function p_CalcRegionValues(ByVal vVal, ByVal SNo As String, ByVal Mult_Calc_USW As String, bLine As Boolean) As Boolean
    ' Calculate the other values from the input values, if one value is changed
    ' IT WAS***********************************************************************************************************************
    ' ********I.e. Initially, change CalculatedSales or if change Multiplier
    ' (CalculatedSales = (ExpectedSales * Multiplier)) and (USW = CalculatedSales / WksOnSale / NumberOfStores)
    ' ********If change USW then
    ' (CalculatedSales = (USW * WksOnSale * NumberOfStores)
    
    ' BUT NOW ********************************************************************************************************************
    
    ''The way that I designed it to work at the moment is :-
    ''    If 'Exp Sales' >0 and  'Sales Multiplier' is changed>0 then
    ''            'Calced Sales' = 'Exp Sales' * 'Sales Multiplier'
    ''            Therefore 'USW' =  'Calced Sales' /  * 'Number of Stores for Region'  /  'Weeks on Sale'
    ''    Else If 'Exp Sales' >0 and  'Calced Sales'  is changed >0 then
    ''            'Sales Multiplier' =  'Calced Sales'  / 'Exp Sales'
    ''            Therefore 'USW' =  'Calced Sales' /  * 'Number of Stores for Region'  /  'Weeks on Sale'
    ''    else If  'USW' is changed >0
    ''            'Calced Sales' ='USW' * 'Number of Stores for Region'  * 'Weeks on Sale'
    ''            'Sales Multiplier' =  'Calced Sales'  / 'Exp Sales'
    ''            (if 'Exp Sales'  = 0 ( or 'Prior Sales'=0) then 'Exp Sales' ='Calced Sales')

    Dim lReg As Long, bOutput As Boolean, aRegs() As String, sFmt As String
    Dim lExp As Long, lCalced As Long, lUPSPW As Long, sngMult As Single '', lWeeks As Long
    Dim cMod As CBA_AST_Product
    Const pCREGS As String = "500,501,502,503,504,505,506,507,509"
    Static lNoStores(1 To 8) As Long, lNoWeeks As Long
    On Error GoTo Err_Routine
    
    p_CalcRegionValues = True
    ' Get the field type
    Select Case Mult_Calc_USW
        Case "SalesMultiplier"
            Mult_Calc_USW = "Mult"
        Case "CalculatedSales"
            Mult_Calc_USW = "Calc"
        Case "UPSPW"
            Mult_Calc_USW = "USW"
        Case Else
            Mult_Calc_USW = ""
            If bLine = False Then p_CalcRegionValues = False
            GoTo Exit_Routine
    End Select
    
    CBA_ErrTag = ""
    aRegs = Split(pCREGS, ",")
    lReg = aRegs(Val(SNo))
    ' Get the number of stores for the region
    If lNoStores(Val(SNo)) = 0 Then
        bOutput = CBA_COM_MATCHGenPullSQL("PORT_NOSTORE", , Date, , , , CStr(lReg))
        If bOutput = True Then
            lNoStores(Val(SNo)) = NZ(CBA_CBISarr(0, 0), 0)
        Else
            lNoStores(Val(SNo)) = 1
        End If
    End If
    ' Get the number of weeks the promotion goes for
    If lNoWeeks = 0 Then
        If CBA_ASyst.CBA_AST_getProdClassMod(cMod) = True Then
            lNoWeeks = cMod.pWoS
        Else
            lNoWeeks = CBA_AST_frm_Products.cboWeeksOfSale
        End If
        'lNoWeeks = g_DLookup("PD_Weeks_Of_Sale", "L2_Products", "PD_ID=" & plProductID, "PD_ID", g_getDB("ASYST"), 4)
    End If
    ' Transfer the fields
    lExp = NZ(Me("txt" & SNo & "ExpectedSales"))
    lCalced = NZ(Me("txt" & SNo & "CalculatedSales"))
    lUPSPW = NZ(Me("txt" & SNo & "UPSPW"))
    sngMult = NZ(Me("txt" & SNo & "SalesMultiplier"))
        ' Get the field type
    Select Case Mult_Calc_USW
        Case "Mult"
            sngMult = NZ(vVal, 0)
        Case "Calc"
            lCalced = NZ(vVal, 0)
        Case "USW"
            lUPSPW = NZ(vVal, 0)
        Case Else
    End Select
 
    ' Calc the values
    If Mult_Calc_USW = "Mult" Then
        If lExp > 0 Then lCalced = Round(lExp * sngMult, 0)
        GoSub CalcUPSPW
    ElseIf Mult_Calc_USW = "Calc" Then
        sngMult = Round(lCalced / IIf(lExp > 0, lExp, lCalced), 4)
        GoSub CalcUPSPW
    ElseIf Mult_Calc_USW = "USW" Then
        lCalced = Round(lUPSPW * lNoWeeks * lNoStores(Val(SNo)), 0)
        sngMult = Round(lCalced / IIf(lExp > 0, lExp, lCalced), 4)
    End If
        
''WAS''''''    If Mult_Calc_USW = "USW" Then
'''''''''''        lCalced = Round(lNoWeeks * sngMult * lUPSPW * lNoStores(Val(sNo)), 0)
'''''''''''    ElseIf Mult_Calc_USW = "Mult" Then
'''''''''''        If lExp > 0 Then
'''''''''''            lCalced = Round(lExp * sngMult, 0)
'''''''''''        Else
'''''''''''            lCalced = Round(lNoWeeks * sngMult * lUPSPW * lNoStores(Val(sNo)), 0)
'''''''''''        End If
'''''''''''        ''lCalced = Round(lNoWeeks * sngMult * lExp * lNoStores(Val(sNo)), 0)
'''''''''''        ''lUPSPW = Round(lCalced / lNoWeeks / lNoStores(Val(sNo)), 0)
'''''''''''''    Else
'''''''''''''        If lExp > 0 Then lCalced = Round(lExp * sngMult, 0)
'''''''''''''        ''lUPSPW = Round(lCalced / lNoWeeks / lNoStores(Val(sNo)), 0)
'''''''''''    End If
    ' Transfer the values back
    sFmt = AST_FillTagArrays("ExpectedSales", plFrmID, plAuth, "Format")
    Me("txt" & SNo & "ExpectedSales") = Format(lExp, sFmt)
    sFmt = AST_FillTagArrays("CalculatedSales", plFrmID, plAuth, "Format")
    Me("txt" & SNo & "CalculatedSales") = Format(lCalced, sFmt)
    sFmt = AST_FillTagArrays("UPSPW", plFrmID, plAuth, "Format")
    Me("txt" & SNo & "UPSPW") = Format(lUPSPW, sFmt)
    sFmt = AST_FillTagArrays("SalesMultiplier", plFrmID, plAuth, "Format")
    Me("txt" & SNo & "SalesMultiplier") = Format(sngMult, sFmt)
      
Exit_Routine:
    On Error Resume Next
    Exit Function
    
    ' Calc UPSPW
CalcUPSPW:
    If lNoWeeks = 0 Or lNoStores(Val(SNo)) = 0 Then
        lUPSPW = 0
    Else
        lUPSPW = Round(lCalced / lNoWeeks / lNoStores(Val(SNo)), 0)
    End If
    Return
    
Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-p_CalcRegionValues", 3)
    CBA_Error = " Error -" & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "ASYST", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
'    GoTo Exit_Routine
    Resume Next
End Function

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Dim lReturn As Long
    If pbHasBeenSaved = False And Me.cmdSave.Visible = True Then
        lReturn = MsgBox("Exit without saving?", vbYesNo + vbDefaultButton2, "Exit warning")
        If lReturn <> vbYes Then
            Cancel = True
            Exit Sub
        End If
        CBA_strAldiMsg = ""
    ElseIf pbHasBeenSaved = False Then
        CBA_strAldiMsg = ""
    End If
End Sub

