VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CBA_AST_frm_Promotions 
   Caption         =   "Super Saver Promotions "
   ClientHeight    =   9630
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15195
   OleObjectBlob   =   "CBA_AST_frm_Promotions.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "CBA_AST_frm_Promotions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit     ' @CBA_ASyst 190426

Private psSQL As String, pbPassOk As Boolean, bNoChg As Boolean
Private Const TABLE_FLDS_A = "PG_ID,PG_Promo_Desc,PG_Status,PG_On_Sale_Date,PG_Weeks_Of_Sale,PG_End_Date," & _
                             "PG_Theme,PG_GBDM_Date,PG_UpdUser,PG_CrtUser"
Private Const TABLE_FLDS_U = "PG_ID,PG_Promo_Desc,PG_Status,PG_On_Sale_Date,PG_Weeks_Of_Sale,PG_End_Date," & _
                             "PG_Theme,PG_GBDM_Date,PG_LastUpd,PG_UpdUser"
Private Const TABLE_FLDS_F = "PG_ID,PG_Promo_Desc,PG_Status,PG_On_Sale_Date,PG_Weeks_Of_Sale,PG_End_Date," & _
                             "PG_Theme,PG_GBDM_Date,PG_LastUpd,PG_CrtDate,PG_UpdUser,PG_CrtUser"
                             
Private Const MAIN_TABLE = "L1_Promotions", MAIN_QUERY = "qry_L1_Promotions"
Private Const SUB_TABLE = "L2_Products", SUB_QUERY = "qry_L2_Products"
Private Const FIELD_PREFIX = "PG_", plFrmID As Long = 1
Private pbAddIP As Boolean, plAllSts As Long, pbGBDMDateChanged As Boolean, psOldGBDMDate As String, plAuth As Long
Private pbUpdIP As Boolean, psLockedIP As String, pbSent As Boolean

'#RW Added new mousewheel routines 190701
Private Sub lstPromo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call mw_SetBoxHook(lstPromo)
End Sub
Private Sub lstProducts_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call mw_SetBoxHook(lstProducts)
End Sub

Private Sub cboAllStatus_Change()
    ' On change of selection Status
    If Me.cboAllStatus.ListIndex > -1 Then
        plAllSts = Me.cboAllStatus.Value
    Else
        plAllSts = 0
    End If
End Sub

Private Sub cboStatus_Change()
    If bNoChg = True Then Exit Sub
    Call p_UpdIP(cboStatus)
End Sub

Private Sub cmdAddPromotion_Click()
    Call g_SetupIP("Promos", 2, True)
    ' Wiil add a promotion to the list
    CBA_lPromotion_ID = 0
    Call AST_FillForm(Me.Controls, TABLE_FLDS_F, FIELD_PREFIX, plFrmID, plAuth, True)
    ' Set the locked and vis status -- should be Visible at this stage...
    Call p_SetSaveVis("Add")
    Me.lstProducts.Clear
    Me.cboStatus.Value = 1
    Me.cboStatus.Locked = True
    Me.cboStatus.BackColor = CBA_Grey
    psOldGBDMDate = ""
    On Error Resume Next
    pbAddIP = True
    Call g_SetupIP("Promos", 2, False)
End Sub

Private Sub cmdSavePromotion_Click()
    ' Validation
    Dim sValidate As String
    sValidate = p_Validate
    If sValidate = "Yes" Then
        Call g_SetupIP("Promos", 2, True)
        If NZ(Me.cboAllStatus.Value, 0) <> plAllSts Then Me.cboAllStatus.Value = Null  ' ?RWAS Try out
        If pbAddIP = True Then
            CBA_lPromotion_ID = CBA_LongHiVal
            Call AST_WriteTable(Me.Controls, TABLE_FLDS_A, FIELD_PREFIX, MAIN_TABLE, plFrmID, plAuth, 0)
            Call p_SetSaveVis("Reset")
            Call p_FillPromoListBox(plAllSts)
            Call p_SetSaveVis("Init")
            CBA_lPromotion_ID = g_DLookup("PG_ID", "L1_Promotions", "PG_ID>0", "PG_ID DESC", g_GetDB("ASYST"), 0)
''            Call lstPromo_Click
        Else
            If pbGBDMDateChanged Then
                ' Change all the product lines to have a GBDM Date as above
                If AST_UpdGBDMDate(Me.txtGBDMDate.Value, CBA_lPromotion_ID) = False Then Exit Sub
                DoEvents
            End If
            Call AST_WriteTable(Me.Controls, TABLE_FLDS_U, FIELD_PREFIX, MAIN_TABLE, plFrmID, plAuth, CBA_lPromotion_ID)               'Me.txtID.Value)
            pbUpdIP = False: psLockedIP = "No"  ' Reset the maintain flags
            Call p_SetSaveVis("Reset")
            Call p_FillPromoListBox
            Call p_SetSaveVis("Init")
        End If
        pbAddIP = False
        pbGBDMDateChanged = False
        Me.lstPromo = Null
        DoEvents
        Me.lstPromo = CStr(CBA_lPromotion_ID)
        ''Me.lstPromo.Selected(Me.lstPromo.ListIndex) = True
        DoEvents
        Call lstPromo_Click
    ElseIf sValidate = "Reset" Then
        Call lstPromo_Click
        pbAddIP = False
    End If
    Call g_SetupIP("Promos", 2, False)


End Sub

Private Sub cmdResetPromotion_Click()
    ' Reset
    If CBA_lPromotion_ID > 0 And CBA_lPromotion_ID < CBA_LongHiVal Then
        If pbUpdIP = True Then
            pbUpdIP = False
            psLockedIP = AST_Lock_Table("ASYST", "PG_", MAIN_TABLE, "PG_ID=" & CBA_lPromotion_ID, plAuth, "SetN")
        End If

        Me.lstPromo.Value = CBA_lPromotion_ID
    Else
        Me.lstPromo.Value = Null
    End If
    Call lstPromo_Click
    pbAddIP = False
End Sub

Private Sub cmdRegion_Click()
    ' Send the prepared Promotions to the regions
    Dim lRegion As Long, dtRegionsDate As Date, bWillbe_Was_Deleted As Boolean
    Const CFLDS As String = "PD_Product_Code,PD_Future_Prod_Code,PD_CGSCG,PD_CGDesc,PD_GBD,PD_BD,Freq_Items,PD_Product_Desc,PD_On_Sale_Date,PD_End_Date," & _
                            "Actual_Marketing_Support,PD_TV,PD_Radio,PD_Out_Of_Home,PD_Press_Dates,PD_Cover_Date,PD_Cannabalised,PD_Complementary," & _
                            "PD_Table_Merch,PD_EndCap_Merch,PV_Fill_Qty,PV_Curr_Retail_Price,PV_Retail_Price,PV_Unit_Cost,PV_Supplier_Cost_Support," & _
                            "PV_UPSPW,Sales_Multiplier,Expected_Sales,PD_ID,PV_ID,PG_Promo_Desc,PG_Theme,Region,Sts_Seq,PD_PG_ID"
    If pbSent = False Then
        ' Set up the prior send date
        If g_IsDate(Me.txtRegionDate, True) Then
            dtRegionsDate = g_FixDate(Me.txtRegionDate, CBA_DMYHN)
        Else
            dtRegionsDate = CDate("00:00")
        End If
        ' Create the xls file to go to the regions
        For lRegion = 501 To 509
            If lRegion <> 508 Then
                dtRegionsDate = CDate(AST_CrtProductRegionSS(CFLDS, "PD_,PV_,PG_", "qry_L3_ProductRegions", "PD_ID,PD_PG_ID,PV_ID", lRegion, ",PD_GBDM_Approved_Date,PDLastUpd,PVLastUpd,PDStatus,PVStatus"))
            End If
        Next lRegion
        ' Flag setup as being IP whilst the field is loaded
        Call g_SetupIP("Promos", 2, True)
    ''    ' Flag the date to be sent into the Promo record - note you have about 90 minutes to change your mind and resend before it is sent to the regions
    ''    Me.txtRegionDate = g_FixDate(dtRegionsDate, CBA_D3DMYHN)
        ' Flag the date to be sent into a Config record
        Me.txtRegionDate = g_FixDate(AST_RegionDate(CStr(dtRegionsDate)), CBA_D3DMYHN)
        
        ' Set up the Send Date Captions
''        Call p_SetUpSendDate
        Call g_SetupIP("Promos", 2, False)
'''        pbSent = True
        MsgBox "Excel files have been created successfully - they will be sent to the regions in approx 30 minutes", vbOKOnly
        Call p_GBDMSts(False, "True")
''        Me.cmdRegion.Caption = "Cancel Send"
    Else
        bWillbe_Was_Deleted = AST_ResetProdRegionSS(CStr(g_FixDate(Me.txtRegionDate, CBA_DMYHN)), False)
        If bWillbe_Was_Deleted = True Then
            If MsgBox("Are you sure you want to delete the prior Region Send?", vbYesNo) = vbYes Then
                Call AST_ResetProdRegionSS(CStr(g_FixDate(Me.txtRegionDate, CBA_DMYHN)), True)
            Else
                bWillbe_Was_Deleted = False
            End If
         Else
             MsgBox "No Region Files were found or Region Files have already been sent"
        End If
''        pbSent = False
        ' Set the cmd key status
        Call p_GBDMSts(bWillbe_Was_Deleted)
    End If
End Sub

Private Sub cmdGBDM_Click()
    ' Send the prepared Promotions to the GBDM
    Dim dtGBDMDate As Date

    Const CFLDS As String = "PD_Product_Code,PD_Future_Prod_Code,PD_CGSCG,PD_CGDesc,PD_GBD,PD_BD,Freq_Items,PD_Product_Desc,PD_On_Sale_Date,PD_End_Date," & _
                            "Actual_Marketing_Support,PD_TV,PD_Radio,PD_Out_Of_Home,PD_Press_Dates,PD_Cover_Date,PD_Cannabalised,PD_Complementary," & _
                            "PD_Table_Merch,PD_EndCap_Merch,PD_Fill_Qty,PD_Curr_Retail_Price,PD_Retail_Price,PD_Unit_Cost,PD_Supplier_Cost_Support," & _
                            "PD_UPSPW,PD_Sales_Multiplier,PD_Expected_Sales,PD_ID,PG_Promo_Desc,PG_Theme,Sts_Seq,PD_PG_ID"
                            
     ' Set up the prior send date
    If g_IsDate(Me.txtGBDMDate, True) Then
        dtGBDMDate = g_FixDate(Me.txtGBDMDate, CBA_DMYHN)
    Else
        dtGBDMDate = CDate("00:00")
    End If
    ' Create the xls file to go to the GBDM
    dtGBDMDate = CDate(AST_CrtProductGBDMSS(CFLDS, "PD_,PG_", "qry_L3_ProductGBDM", "PD_ID,PD_PG_ID"))
    ' Flag setup as being IP whilst the field is loaded
    Call g_SetupIP("Promos", 2, True)
    ' Flag the date to be sent into a Config record
    Me.txtGBDMDate = g_FixDate(AST_GBDMDate(CStr(dtGBDMDate)), CBA_D3DMYHN)
  ''  Call g_SetupIP("Promos", 2, False)
    
    MsgBox "Excel file has been created successfully - press ok to continue" & vbCrLf & "(Promo management will be closed to make the Excel file Visible)", vbOKOnly
    Unload Me
End Sub

Private Sub lblProducts_Click()
    If Me.txtPromoDesc.BackColor = CBA_Grey Then Exit Sub
    CBA_lProduct_ID = 0         ' User hasn't selected a product (clicked on a label) so set to zero
    CBA_AST_frm_Products.Show
End Sub

Private Sub lstProducts_Click()
    If Me.txtPromoDesc.BackColor = CBA_Grey Then Exit Sub
    ' Take this out for now as it is not called for and forms will be called modeless now
''    If Me.lstProducts.ListIndex > -1 Then
''        ' Show the Products form for the selection
''        CBA_lProduct_ID = Me.lstProducts.Column(0, Me.lstProducts.ListIndex)
''        CBA_AST_frm_Products.Show
''    Else
''        CBA_lProduct_ID = 0
''        CBA_AST_frm_Products.Show
''    End If
End Sub

Private Sub lstPromo_Click()
    ' Get the Promo specified
    If bNoChg = True Then Exit Sub
    ' Flag setup as being IP whilst the form is loaded
    Call g_SetupIP("Promos", 2, True)
    If Me.lstPromo.ListIndex > -1 Then
        CBA_lPromotion_ID = Me.lstPromo.Column(0, Me.lstPromo.ListIndex)
        ' Fill the fields
        CBA_DBtoQuery = 1
        psSQL = "SELECT " & TABLE_FLDS_F & " FROM [" & MAIN_TABLE & "] WHERE PG_ID=" & CBA_lPromotion_ID & ";"
        pbPassOk = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", MAIN_TABLE, g_GetDB("ASYST"), CBA_MSAccess, psSQL, 120, , , False)
        If pbPassOk = True Then
            Call AST_FillForm(Me.Controls, TABLE_FLDS_F, FIELD_PREFIX, plFrmID, plAuth)
            ' Fill the Theme / Promotion List Box
            CBA_DBtoQuery = 1
            psSQL = "SELECT PD_ID,PD_Product_Code,PD_Product_Desc FROM [" & SUB_TABLE & "] WHERE PD_PG_ID=" & CBA_lPromotion_ID & ";"
            pbPassOk = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", SUB_TABLE, g_GetDB("ASYST"), CBA_MSAccess, psSQL, 120, , , False)
            Me.lstProducts.Clear
            If pbPassOk = True Then
                Call AST_FillListBox(Me.lstProducts, 3, , True)
            End If
            CBA_DBtoQuery = 3
            psOldGBDMDate = Me.txtGBDMDate
            ' Set the locked and vis status -- fields may be shown at this stage...
            Call p_SetSaveVis("Reset")
''            DoEvents
''            Debug.Print Me.cboStatus.Value & "-";
            If Me.cboStatus.Value = 8 Then Me.cboStatus.Locked = True
            ' Set up the Send Captions
''            Call p_SetUpSendDate
        End If
    Else
        psOldGBDMDate = ""
        Call p_SetSaveVis("Init")
    End If
    Call g_SetupIP("Promos", 2, False)
End Sub

Private Sub txtGBDMDate_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim bHasDates As Boolean
    If Me.txtGBDMDate.BackColor = CBA_Grey Then Exit Sub
    If Me.cboStatus.Value >= 5 Then
        MsgBox "GBDM Date cannot be set as the Promotion has the wrong status to be changed", vbOKOnly
        Exit Sub
    End If
    ' Test to see if the date can be applied - some fields may not be filled in the l2_products table....
    If g_IsDate(Me.txtGBDMDate.Value, True) = False Then
        If AST_UpdGBDMDate(g_FixDate(Me.txtGBDMDate.Value), CBA_lPromotion_ID, True, , "Fatal") <> "True" Then Exit Sub
    Else
        bHasDates = True
    End If
    Call CBA_BT_CalendarShow(Me.Top, Me.Left, Me.txtGBDMDate, CBA_D3DMY)
    If varCal.bCalValReturned = False Then Exit Sub
    If bHasDates = True Then Call AST_UpdGBDMDate(g_FixDate(Me.txtGBDMDate.Value), CBA_lPromotion_ID, , , "All")  ' Don't call as if it comes to here, the lines already have GBDM dates
    Application.SendKeys "{TAB}"
    bNoChg = True
    Call p_UpdIP(Me.txtGBDMDate, True)
    If g_IsDate(Me.txtGBDMDate.Value, True) = True And Me.cboStatus.Value >= 3 Then
    ElseIf g_IsDate(Me.txtGBDMDate.Value, True) = True And Me.cboStatus.Value < 3 Then
        Me.cboStatus.Value = 3                          ' Set to 'GBDM Approved'
    ElseIf Me.cboStatus.Value = 3 Then
        Me.cboStatus.Value = 1                          ' Set to 'In Dev...'
    End If
    'Debug.Print Me.cboStatus.Value
    bNoChg = False
End Sub

''Private Sub txtRegionDate_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
''    If Me.txtRegionDate.BackColor = CBA_Grey Then Exit Sub
''    Call CBA_BT_CalendarShow(Me.Top, Me.Left, Me.txtRegionDate, CBA_D3DMYHN)
''    Call p_UpdIP(Me.txtRegionDate, True)
''End Sub

Private Sub txtOnSaleDate_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Me.txtOnSaleDate.BackColor = CBA_Grey Then Exit Sub
    Call CBA_BT_CalendarShow(Me.Top, Me.Left, Me.txtOnSaleDate, CBA_D3DMY, False)
    Call AST_setEndDate(Me.Name, Me.Controls, Me.txtOnSaleDate, CBA_D3DMY)
    Call p_UpdIP(Me.txtOnSaleDate, True)
End Sub

Private Sub cboWeeksOfSale_Change()
    Call AST_setEndDate(Me.Name, Me.Controls, Me.cboWeeksOfSale, CBA_D3DMY)
    Call p_UpdIP(Me.cboWeeksOfSale)
End Sub

Private Sub txtTheme_Change()
    Call p_UpdIP(Me.txtTheme)
End Sub

Private Sub txtPromoDesc_Change()
    Call p_UpdIP(Me.txtPromoDesc)
End Sub

Private Sub UserForm_Initialize()

    ' Flag setup as being IP
    Call g_SetupIP("Promos", 1, True, True)
    Me.StartUpPosition = 0
    Me.Top = Application.Top + (Application.Height / 2) - (Me.Height / 2)
    Me.Left = Application.Left + (Application.Width / 2) - (Me.Width / 2)
    ' Capture the Top, Width and Left positions
    Call g_PosForm(Me.Top, Me.Width, Me.Left, , True)
    ' Get the authority level
    plAuth = CBA_lAuthority
    Me.lblVersion.Caption = CBA_getVersionStatus(g_GetDB("ASYST"), CBA_AST_Ver, "Super Saver Tool", "AST") & "-" & plAuth       ' Get the latest version
    CBA_lPromotion_ID = 0
    ' Set Weeks of Sale combo
    Call AST_WeeksOfSale(cboWeeksOfSale)
    ' Fill the Current Status DDBox
    CBA_DBtoQuery = 1
    psSQL = "SELECT Sts_ID,Sts_Desc FROM [L0_Statuses] WHERE Sts_PromotionValid='Y' ORDER BY Sts_ID"
    pbPassOk = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", "L0_Statuses", g_GetDB("ASYST"), CBA_MSAccess, psSQL, 120, , , False)
    If pbPassOk = True Then
        Call AST_FillDDBox(Me.cboStatus, 2)
        Call AST_FillDDBox(Me.cboAllStatus, 2)
    End If
    CBA_DBtoQuery = 3
    ' Fill the Theme / Promotion List Box
    Call p_FillPromoListBox
    ' Set the date of the last GBDM Report
    Me.txtGBDM = g_FixDate(AST_GBDMDate(""), CBA_D3DMYHN)
    Me.txtGBDM.BackColor = CBA_Grey
    ' Set the date of the last Region Report
    Me.txtRegionDate = g_FixDate(AST_RegionDate(""), CBA_D3DMYHN)
    Me.txtRegionDate.BackColor = CBA_Grey
   ' Set the locked and vis status -- should be hidden at this stage...
    Call p_SetSaveVis("Init", True)
    Call p_GBDMSts(AST_ResetProdRegionSS(CStr(g_FixDate(Me.txtRegionDate, CBA_DMYHN)), False), "Test")
    On Error Resume Next
    Call g_SetupIP("Promos", 1, False)
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Dim lReturn As Long
    If Me.cmdSavePromotion.Visible = True Then
        lReturn = MsgBox("Exit without saving?", vbYesNo + vbDefaultButton2, "Exit warning")
        If lReturn <> vbYes Then
            Cancel = True
            Exit Sub
        End If
    End If
    ' Null global values
    CBA_lPromotion_ID = 0
    CBA_lProduct_ID = 0
End Sub

Private Function p_Validate() As String
    ' This routine will validate the entries made or not made
    Dim sCrtDate As String, sEndDate As String, sGBDMDate As String, sOnSaleDate As String, sName As String, sMsg As String, sDate As String
    sCrtDate = IIf(g_IsDate(Me.txtCrtDate, True) = True, g_FixDate(Me.txtCrtDate), "")
    sEndDate = IIf(g_IsDate(Me.txtEndDate, True) = True, g_FixDate(Me.txtEndDate), "")
    sGBDMDate = IIf(g_IsDate(Me.txtGBDMDate, True) = True, g_FixDate(Me.txtGBDMDate), "")
    sOnSaleDate = IIf(g_IsDate(Me.txtOnSaleDate, True) = True, g_FixDate(Me.txtOnSaleDate), "")
    p_Validate = "Yes": sName = "": sMsg = ""
    If Len(NZ(Me.txtPromoDesc, "")) < 8 Then
        sMsg = "Promotion needs a title of at least 8 characters"
        sName = "txtPromoDesc"
    ElseIf Val(Me.cboWeeksOfSale) < 1 Then
        sMsg = "Weeks of Sale must be 1 or greater"
        sName = "WeeksOfSale"
    ElseIf CDate(IIf(g_IsDate(Me.txtGBDMDate.Value, True), g_FixDate(Me.txtGBDMDate.Value), Date)) > Date Then
        sMsg = "'GBDM Approval Date' can't be a date in the future"
        sName = "txtGBDMDate"
    ElseIf g_IsDate(Me.txtGBDMDate.Value, True) = True And Me.cboStatus.Value < 3 Then
        sMsg = "If 'GBDM Approval Date' is set, Status must be 'GBDM Approved', 'Suspended' or 'Completed'"
        sName = "cboStatus"
    ElseIf g_IsDate(Me.txtGBDMDate.Value, True) = False And Me.cboStatus.Value = 3 Then
        sMsg = "If Status is 'GBDM Approved', 'GBDM Approval Date' must be set"
        sName = "cboStatus"
    ElseIf pbAddIP = False Then
        sDate = g_FixDate(g_DLookup("PG_LastUpd", MAIN_TABLE, "PG_ID=" & CBA_lPromotion_ID, "", g_GetDB("ASYST"), "01/01/1900"), CBA_DMYHN)
        If sDate <> g_FixDate(Me.txtLastUpd.Value, CBA_DMYHN) Then
            sMsg = "Record has been updated by another while this record was being prepared - update will now cancel"
            sName = "txtLastUpd"
        End If
    End If
    ' If an error
    If sMsg > "" Then
        p_Validate = "No"
        If MsgBox(sMsg & vbCrLf & "Update will cancel", vbOKOnly, "Validation Warning") = vbOK Then
''            p_Validate = "Reset"
''        Else
            On Error Resume Next
            Me(sName).SetFocus
        End If
        Exit Function
    Else
        ' If the GBDM Date has changed, flag the record so that the lines can be updated too.
        If psOldGBDMDate <> CStr(Me.txtGBDMDate) Then pbGBDMDateChanged = True
    End If

End Function

Sub p_FillPromoListBox(Optional lWhere = 0)
    ' Fill the Theme / Promotion List Box
    Dim lPos As Long
    lPos = CBA_lPromotion_ID
    bNoChg = True
    CBA_DBtoQuery = 1
    If lWhere = 0 Then
        psSQL = "SELECT PG_ID,PG_Promo_Desc,PG_Status1 FROM [" & MAIN_QUERY & "] WHERE PG_Status > 0 ORDER BY Sts_Seq,PG_ID"
    Else
        psSQL = "SELECT PG_ID,PG_Promo_Desc, PG_Status1 FROM [" & MAIN_QUERY & "] WHERE PG_Status = " & lWhere & " ORDER BY Sts_Seq,PG_ID"
    End If
    pbPassOk = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", MAIN_QUERY, g_GetDB("ASYST"), CBA_MSAccess, psSQL, 120, , , False)
    If pbPassOk = True Then
        Call AST_FillListBox(Me.lstPromo, 3)
    End If
    CBA_DBtoQuery = 3
    bNoChg = False
End Sub

Sub p_SetSaveVis(Add_Reset_Init As String, Optional bInit As Boolean = False)
    
    Dim bLocked As Boolean, bAddVis As Boolean, bSaveVis As Boolean, bResetVis As Boolean, bLockOR As Boolean
    Static sLast As String
    ' if bLockOR = true will lock and grey everything
    Select Case Add_Reset_Init
    Case "Add"                                                  ' After an Add...
        bAddVis = False: bSaveVis = True: bResetVis = bSaveVis: bLockOR = False
    Case "Save"                                                 ' After a Save...
        bAddVis = False: bSaveVis = True: bResetVis = bSaveVis: bLockOR = False
    Case "Reset"                                                ' After a reset
        bAddVis = True: bSaveVis = False: bResetVis = bSaveVis: bLockOR = False
    Case "Init"                                                 ' Upon first entry - if there are Promos existing
        bAddVis = True: bSaveVis = False: bResetVis = bSaveVis: bLockOR = True
    Case Else
        MsgBox "no data"
    End Select
''    if nz(me.
''    If bInit Then bAddVis = False
    If Add_Reset_Init <> sLast Then
        ' Set the Lock / Visibility of the fields on the form
        Call AST_SetLockVis(Me.Controls, bLockOR, plFrmID, plAuth)
        Add_Reset_Init = sLast
    End If
    ' Set the  Visibility of the Cmd Keys and the ListBoxes on the form
    bLocked = bSaveVis
    ' If there has been an error (someone is already maintaining this record) then don't display the save key
    If psLockedIP = "AlreadyLocked" Then bSaveVis = False
    Call AST_SetCmdVis(Me.cmdSavePromotion, bSaveVis, plFrmID, plAuth)
    Call AST_SetCmdVis(Me.cmdAddPromotion, bAddVis, plFrmID, plAuth)
    Call AST_SetCmdVis(Me.cmdResetPromotion, bResetVis, plFrmID, plAuth)
    Me.lstProducts.Locked = bLocked
    Me.lstProducts.BackColor = IIf(Me.lstProducts.Locked = True, CBA_Grey, CBA_White)
    Me.lstPromo.Locked = bLocked
    ''Me.lstPromo.BackColor = IIf(Me.lstPromo.Locked = True, CBA_Grey, CBA_White)
    'If g_IsDate(Me.txtGBDMDate, True) And CBA_lAuthority = 1 And bSaveVis = False And psLockedIP <> "AlreadyLocked" Then
    If plAuth = 1 Then Me.cmdRegion.Visible = True
    'Else
        'cmdRegion.Visible = False
    'End If
    If plAuth = 1 Then Me.cmdGBDM.Visible = True
    
End Sub

Sub p_UpdIP(cActCtl As Control, Optional bIsDate As Boolean = False)
    Static bRecursed As Boolean
    Dim sClr As String
    If bRecursed = True Then Exit Sub
    bRecursed = True

    ' Flag the record as being updated
    If g_SetupIP("Promos") = False Then
        If pbUpdIP = False Then
            pbUpdIP = True
            ' Test to ensure the record is not already held and set psLockedIP = "AlreadyLocked" id it is
            psLockedIP = AST_Lock_Table("ASYST", "PG_", MAIN_TABLE, "PG_ID=" & CBA_lPromotion_ID, CBA_lAuthority, "Get_SetY")
        End If
            
        ' Check to see if the colour needs changing
        sClr = AST_FillTagArrays(cActCtl.Tag, plFrmID, plAuth, "Clr")
        Call AST_FillYellow(cActCtl, sClr)
        ' If a date, check to see if it has changed, and hop out if hasn't
        If bIsDate = True Then
            If varCal.bCalValReturned = False Then GoTo Exit_Routine
        End If
        Call AST_getFieldTag(cActCtl, plFrmID, plAuth, , "ApdUpd") ' Will append the 'IsUpdated' flag "~" to tag
        Call p_SetSaveVis("Save")
    End If
Exit_Routine:
    bRecursed = False
End Sub

Private Sub p_GBDMSts(ByVal bDeleted As Boolean, Optional ByVal sSent As String = "False")
    ' Will set the Me.cmdRegion.Caption according to its status
    If sSent = "Test" And bDeleted = False Then                                     ' Find out if the last date sent was less than an hour ago
        sSent = "False"
        If CDate(g_FixDate(NZ(Me.txtRegionDate, "01/01/2000"), CBA_DMYHN)) > CDate(g_FixDate(DateAdd("n", -60, Now()), CBA_DMYHN)) Then sSent = "True": bDeleted = False
    End If
    ' Set Caption
    If sSent = "True" And bDeleted = False Then             ' If it looks to be just sent...
        Me.cmdRegion.Caption = "Cancel Send"
        pbSent = True
    ElseIf bDeleted = True Then                             ' If it has only been just sent and then deleted...
        Me.cmdRegion.Caption = "Generate and Re-Send to Regions"
        pbSent = True
    Else                                                    ' Will be sent in a new unit of time...
        Me.cmdRegion.Caption = "Generate and Send to Regions"
        pbSent = False
    End If
End Sub


''Private Sub p_SetUpSendDate()
''''    ' Set up the Send Date Captions
''''    Dim bSent As Boolean, bReSend As Boolean
''''    bSent = False: bReSend = False
''''    If g_IsDate(Me.txtRegionDate, True) Then
''''        bReSend = True
''''        If CDate(g_FixDate(Me.txtRegionDate, CBA_DMYHN)) > CDate(g_FixDate(Now(), CBA_DMYHN)) Then bSent = True
''''    End If
''End Sub





