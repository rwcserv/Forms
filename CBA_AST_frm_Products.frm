VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CBA_AST_frm_Products 
   Caption         =   "Super Saver Products"
   ClientHeight    =   9255
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   17010
   OleObjectBlob   =   "CBA_AST_frm_Products.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "CBA_AST_frm_Products"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit   ' @CBA_ASyst 190625

Private Const TABLE_FLDS_A = "PD_ID,PD_PG_ID,PD_Status,PD_On_Sale_Date,PD_Weeks_Of_Sale,PD_End_Date,PD_Product_Code,PD_Future_Prod_Code,PD_Prior_UPSPW,PD_UPSPW," & _
                            "PD_Product_Desc,PD_Retail_Price,PD_Orig_Unit_Cost,PD_Unit_Cost,PD_Supplier_Cost_Support,PD_Retail_Discount,PD_Theme,PD_Cannabalised,PD_Calculated_Sales," & _
                            "PD_Sales_Multiplier,PD_Super_Saver_Type,PD_Sub_Basket,PD_Elasticity_Flag,PD_Freq_Items,PD_Curr_Retail_Price," & _
                            "PD_Elasticity_Confidence_Flag,PD_Expected_Marketing_Support,PD_Actual_Marketing_Support,PD_Product_Approval_Date," & _
                            "PD_GBDM_Approval_Date,PD_GBDM_Approved_Date,PD_Approval_Req_Wks,PD_Approval_Req_Date,PD_Approval_Req_Cmt,PD_Complementary,PD_GST," & _
                            "PD_Out_Of_Home,PD_TV,PD_Radio,PD_Table_Merch,PD_EndCap_Merch,PD_Regions,PD_Press_Dates,PD_Cover_Date," & _
                            "PD_Standee,PD_CGSCG,PD_CGDesc,PD_ACatCGSCG,PD_ACatDesc,PD_Prior_Sales,PD_Expected_Sales,PD_Comment,PD_GBD,PD_BD,PD_UpdUser,PD_CrtUser"
Private Const TABLE_FLDS_U = "PD_ID,PD_PG_ID,PD_Status,PD_On_Sale_Date,PD_Weeks_Of_Sale,PD_End_Date,PD_Product_Code,PD_Future_Prod_Code,PD_Prior_UPSPW,PD_UPSPW," & _
                            "PD_Product_Desc,PD_Retail_Price,PD_Orig_Unit_Cost,PD_Unit_Cost,PD_Supplier_Cost_Support,PD_Retail_Discount,PD_Theme,PD_Cannabalised,PD_Calculated_Sales," & _
                            "PD_Sales_Multiplier,PD_Super_Saver_Type,PD_Sub_Basket,PD_Elasticity_Flag,PD_Freq_Items," & _
                            "PD_Elasticity_Confidence_Flag,PD_Expected_Marketing_Support,PD_Actual_Marketing_Support,PD_Product_Approval_Date," & _
                            "PD_GBDM_Approval_Date,PD_GBDM_Approved_Date,PD_Approval_Req_Wks,PD_Approval_Req_Date,PD_Approval_Req_Cmt,PD_Complementary,PD_GST," & _
                            "PD_Out_Of_Home,PD_TV,PD_Radio,PD_Table_Merch,PD_EndCap_Merch,PD_Regions,PD_Press_Dates,PD_Cover_Date," & _
                            "PD_Standee,PD_Prior_Sales,PD_Expected_Sales,PD_Comment,PD_GBD,PD_BD,PD_LastUpd,PD_UpdUser"
Private Const TABLE_FLDS_F = "PD_ID,PD_PG_ID,PD_Status,PD_On_Sale_Date,PD_Weeks_Of_Sale,PD_End_Date,PD_Product_Code,PD_Future_Prod_Code,PD_Prior_UPSPW,PD_UPSPW," & _
                            "PD_Product_Desc,PD_Retail_Price,PD_Orig_Unit_Cost,PD_Unit_Cost,PD_Supplier_Cost_Support,PD_Retail_Discount,PD_Theme,PD_Cannabalised,PD_Calculated_Sales," & _
                            "PD_Sales_Multiplier,PD_Super_Saver_Type,PD_Sub_Basket,PD_Elasticity_Flag,PD_Freq_Items,PD_Curr_Retail_Price," & _
                            "PD_Elasticity_Confidence_Flag,PD_Expected_Marketing_Support,PD_Actual_Marketing_Support,PD_Product_Approval_Date," & _
                            "PD_GBDM_Approval_Date,PD_GBDM_Approved_Date,PD_Approval_Req_Wks,PD_Approval_Req_Date,PD_Approval_Req_Cmt,PD_Complementary,PD_GST," & _
                            "PD_Standee,PD_Out_Of_Home,PD_TV,PD_Radio,PD_Table_Merch,PD_EndCap_Merch,PD_Regions,PD_Press_Dates,PD_Cover_Date," & _
                            "PD_CGSCG,PD_CGDesc,PD_ACatCGSCG,PD_ACatDesc,PD_Prior_Sales,PD_Expected_Sales,PD_Comment,PD_GBD,PD_BD,PD_LastUpd,PD_UpdUser,PD_CrtDate,PD_CrtUser"



Private Const MAIN_TABLE = "L2_Products" '''', MAIN_QUERY = "qry_L2_Products"
Private Const FIELD_PREFIX = "PD_", plFrmID As Long = 2
''Private Const READ_ONLY_EXCEPTS As String = "OrigUnitCost,UnitCost,SupplierCostSupport,PriorSales, ExpectedSales,UPSPW,CurrRetailPrice,RetailPrice,SalesMultiplier"

Private psSQL As String, pbPassOk As Boolean, psProductCode As String, pbNullFields As Boolean, pbHasRows As Boolean, psValidate As String
Private pbAddIP As Boolean, pbValidateAsOK As Boolean
Private pdblElastic As Single, pdblAvg_Price As Single, psElasticity_State As String, psConfidence_Level As String, plCGno As Long, plSCGNo As Long
Private pbSalesChanged As Boolean, pbIsGBDMDate As Boolean
Private plAllSts As Long, plAuth As Long
Private pbUpdIP As Boolean, psLockedIP As String, psReadOnlyExcepts As String
Private Sub cboAllStatus_Change()
    ' On change of selection Status
    If Me.cboAllStatus.ListIndex > -1 Then
        plAllSts = Me.cboAllStatus.Value
    Else
        plAllSts = 0
    End If
End Sub
Private Sub Exp_Prods_Click()
    p_ExportLBox Me.lstProducts ' If p_ExportLBox(Me.lstProducts) = True Then MsgBox "Data Exported"
End Sub
Private Sub Exp_Cannibalised_Click()
    p_ExportLBox Me.txtCannabalised ' If p_ExportLBox(Me.txtCannabalised) = True Then MsgBox "Data Exported"
End Sub
Private Sub Exp_Compliment_Click()
    p_ExportLBox Me.txtComplementary 'If p_ExportLBox(Me.txtComplementary) = True Then MsgBox "Data Exported"
End Sub
Private Sub Exp_Specials_Click()
    p_ExportLBox Me.lstSpecial 'If p_ExportLBox(Me.lstSpecial) = True Then MsgBox "Data Exported"
End Sub
Private Sub Exp_Theme_Click()
    p_ExportLBox Me.lstTheme 'If p_ExportLBox(Me.lstTheme) = True Then MsgBox "Data Exported"
End Sub
Private Function p_ExportLBox(ByRef LB As Object) As Boolean
Dim l As Variant
Dim ls As ListBox
Dim wbk As Workbook
Dim dta As Variant
    If TypeName(LB) = "ListBox" Then
        If LB.ListCount = 0 Then
            MsgBox "Nothing to Export"
            p_ExportLBox = False
            Exit Function
        End If
    ElseIf TypeName(LB) = "TextBox" Then
        If LB.Value = "" Then
            MsgBox "Nothing to Export"
            p_ExportLBox = False
            Exit Function
        End If
    End If
    Set wbk = Workbooks.Add
    With ActiveSheet
    Range(.Cells(1, 1), .Cells(5, 79)).Interior.ColorIndex = 49
    .Cells(1, 1).Select
    .Pictures.Insert CBA_BSA & "VBA Development Tools\IMAGES\ALDI Logo NEW mod HighRes.png"
    .Cells.Font.Name = "ALDI SUED Office"
    .Cells(4, 3).Font.Size = 24
    .Cells(4, 3).Font.ColorIndex = 2
    Range(.Cells(4, 8), .Cells(4, 9)).Font.ColorIndex = 2
        If LB.Name = Me.lstProducts.Name Then
            .Cells(4, 3).Value = "Promotional Products - " & Me.cboPromotionID.text
            .Cells(6, 1).Value = "Productcode"
            .Cells(6, 2).Value = "Product Description"
            For l = 0 To LB.ListCount - 1
                .Cells(l + 7, 1).Value = LB.List(l, 1)
                .Cells(l + 7, 2).Value = LB.List(l, 2)
            Next
            Range(.Cells(1, 1), .Cells(1, 2)).EntireColumn.AutoFit
        End If
        If LB.Name = Me.lstSpecial.Name Then
            .Cells(4, 3).Value = "Specials from same CG during SuperSaver - " & Me.cboPromotionID.text
            .Cells(6, 1).Value = "Productcode"
            .Cells(6, 2).Value = "On Sale Date"
            .Cells(6, 3).Value = "Product Description"
            For l = 0 To LB.ListCount - 1
                .Cells(l + 7, 1).Value = LB.List(l, 1)
                .Cells(l + 7, 2).Value = LB.List(l, 2)
                .Cells(l + 7, 3).Value = LB.List(l, 3)
            Next
            Range(.Cells(1, 1), .Cells(1, 3)).EntireColumn.AutoFit
        End If
        If LB.Name = Me.lstTheme.Name Then
            .Cells(4, 3).Value = "Themes during Super Saver Period - " & Me.cboPromotionID.text
            .Cells(6, 1).Value = "Theme"
            .Cells(6, 2).Value = "On Sale Date"
            For l = 0 To LB.ListCount - 1
                .Cells(l + 7, 1).Value = LB.List(l, 0)
                .Cells(l + 7, 2).Value = LB.List(l, 1)
            Next
            Range(.Cells(1, 1), .Cells(1, 2)).EntireColumn.AutoFit
        End If
        If LB.Name = Me.txtCannabalised.Name Then
            .Cells(4, 3).Value = Me.txtProductCode & " - " & Me.txtProductDesc & " Cannibalised Products - Promo: " & Me.cboPromotionID.text
            .Cells(6, 1).Value = "Productcode"
            .Cells(6, 2).Value = "Product Description"
            dta = Split(Me.txtCannabalised.Value, Chr(10))
            For l = LBound(dta) To UBound(dta)
                .Cells(l + 7, 1).Value = Mid(dta(l), 1, InStr(1, dta(l), "-") - 1)
                .Cells(l + 7, 2).Value = Mid(dta(l), InStr(1, dta(l), "-") + 1, 99)
            Next
            Range(.Cells(1, 1), .Cells(1, 2)).EntireColumn.AutoFit
        End If
        If LB.Name = Me.txtComplementary.Name Then
            .Cells(4, 3).Value = Me.txtProductCode & " - " & Me.txtProductDesc & " Complimentary Products - Promo: " & Me.cboPromotionID.text
            .Cells(6, 1).Value = "Productcode"
            .Cells(6, 2).Value = "Product Description"
            dta = Split(Me.txtComplementary.Value, Chr(10))
            For l = LBound(dta) To UBound(dta)
                .Cells(l + 7, 1).Value = Mid(dta(l), 1, InStr(1, dta(l), "-") - 1)
                .Cells(l + 7, 2).Value = Mid(dta(l), InStr(1, dta(l), "-") + 1, 99)
            Next
            Range(.Cells(1, 1), .Cells(1, 2)).EntireColumn.AutoFit
        End If
    End With
    p_ExportLBox = True
End Function
Private Sub cboMedium_Change()
    Static sLastName As String, sName As String
    ' Set the prior field to invisible
    If g_SetupIP("Products") Then Exit Sub
    If sLastName > "" Then
        Me("lbl" & sLastName).Visible = False
        Me("txt" & sLastName).Visible = False
    End If
    ' Which field are we going to look at
    If Me.cboMedium.ListIndex > -1 Then
        Select Case Me.cboMedium
        Case "Out Of Home"
            sName = "OutOfHome"
        Case "Radio"
            sName = "Radio"
        Case "Television"
            sName = "TV"
        Case "Press"
            sName = "PressDates"
        Case "Standee"
            sName = "Standee"
        Case Else
            sName = ""
        End Select
        ' Set the selected field to Visible
        If sName > "" Then
            Me("lbl" & sName).Visible = True
            Me("txt" & sName).Visible = True
        End If
        sLastName = sName
    Else
        sLastName = ""
        sName = "OutOfHome"
        GoSub GSFields
        sName = "Radio"
        GoSub GSFields
        sName = "TV"
        GoSub GSFields
        sName = "PressDates"
        GoSub GSFields
        sName = "Standee"
        GoSub GSFields
    End If
    Exit Sub

GSFields:
    Me("lbl" & sName).Visible = False
    Me("txt" & sName).Visible = False
    Return

End Sub

Private Sub cboPromotionID_Change()
    ' Change of Promo
    If g_SetupIP("Products") = False Then
        Call g_SetupIP("Products", 4, True)
        ' Fill the Product List Box if the Promotion has changed
        If Me.cboPromotionID.ListIndex > -1 Then
            ' Call p_Progbar("cboPromotionID", , ">0")
            CBA_lPromotion_ID = Me.cboPromotionID.Column(0, Me.cboPromotionID.ListIndex)
            Me.txtPGID = CBA_lPromotion_ID
            CBA_lProduct_ID = 0
            plAuth = CBA_lAuthority
            Me.lstProducts.Clear
            Call p_FillLstProducts
        Else
            Me.lstProducts.Clear
            CBA_lProduct_ID = 0
        End If
        ' Re-Select the product
        Call p_SelectProduct
        Call g_SetupIP("Products", 4, False)
    End If
End Sub

Private Sub cboPromotionID_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call mw_SetBoxHook(cboPromotionID)
End Sub

Private Sub cboSuperSaverType_Change()
    Call p_UpdIP(Me.cboSuperSaverType)
    Me.txtApprovalReqDate.Value = g_FixDate(CDate(g_FixDate(Me.txtOnSaleDate)) - p_LeadWeeks, CBA_D3DMY)
    Call p_TestApprovals
End Sub

Private Sub cmdAddProduct_Click()
    ' Will add a product to the list
    Dim lProdCode As Long, lIdx As Long
    CBA_strAldiMsg = "Product_Code"
    CBA_AST_frm_AssocFld.Show vbModal
    If CBA_strAldiMsg > "" Then lProdCode = Val(CBA_strAldiMsg)
    If lProdCode = 0 Then Exit Sub
    ' Call p_Progbar("CmdAdd", 10, "PC=" & lProdCode)
    ' Check to ensure the product exists in CBis
    ' Call p_Progbar("pGetProdDets", 10, "Test")
    If p_GetProdDets(CStr(lProdCode), True) = False Then                           ' Get the product details
        MsgBox "Product " & lProdCode & " Can't be found", vbOKOnly
        Exit Sub
    End If
    ' Check to ensure the product doesn't already exist in the promotion
    For lIdx = 0 To lstProducts.ListCount - 1
        If lProdCode = Me.lstProducts.Column(1, lIdx) Then
            If MsgBox("Product " & lProdCode & " already exits in this Promotion... Press Yes to add it anyway", vbYesNo, "Product Add") = vbNo Then
                Exit Sub
            End If
        End If
    Next lIdx
    plAuth = CBA_lAuthority
    ' Call p_Progbar("Init")
    Call g_SetupIP("Products", 3, True)  ' To stop prior updating of fields, setup is on
    psProductCode = lProdCode
    pbAddIP = True
    ' Addrec setup
    CBA_lProduct_ID = 0
''    Call p_SetupForm("AddRec")
    ' Call p_Progbar("ASTFillForm", 10, "NULL")
    ' ******* Fill the form and null all the fields
    pbNullFields = True: pbIsGBDMDate = False
    Call AST_FillForm(Me.Controls, TABLE_FLDS_F, FIELD_PREFIX, plFrmID, plAuth, pbNullFields)      ' Null the fields on the form
    ''Call p_NullLists("Init") done later
    ' Call p_Progbar("SetSaveVis", 10, "Add")
    Call p_SetSaveVis("Add")                                                                        ' Set the form cmd key visibility
    Me.txtID = 0
    Me.txtOrigUnitCost = 0
    Me.txtUnitCost = 0
    Me.txtCurrRetailPrice = 0
    Me.txtOnSaleDate = g_RevDate(Me.cboPromotionID.Column(3, Me.cboPromotionID.ListIndex), , CBA_D3DMY)
    Me.cboWeeksOfSale = Me.cboPromotionID.Column(4, Me.cboPromotionID.ListIndex)
    Me.txtEndDate = g_RevDate(Me.cboPromotionID.Column(5, Me.cboPromotionID.ListIndex), , CBA_D3DMY)
    Me.txtTheme = Me.cboPromotionID.Column(7, Me.cboPromotionID.ListIndex)
    Call AST_FillYellow(Me.txtTheme, "y")
    Me.txtRegions = "All"
    Me.txtProductCode = psProductCode
    Me.cboStatus.Value = 1
    Me.txtSalesMultiplier.Value = 1
    Me.txtRetailDiscount = 0
    Me.txtUPSPW = 0
    Me.txtPriorUPSPW = 0
''    Me.cboStatus.Locked = True
''    Me.cboStatus.BackColor = CBA_Grey
''    Call txtUnitCost_AfterUpdate
    Me.txtSupplierCostSupport = 0
''    Call txtSupplierCostSupport_AfterUpdate
    Me.chkEndCapMerch = False
    Me.chkTableMerch = False
''    Me.txtCoverDate = "N/A"
    Me.txtOutOfHome = "N/A"
    Me.txtFutureProdCode = "N/A"
    Me.txtPressDates = "N/A"
    Me.txtRadio = "N/A"
    Me.txtTV = "N/A"
    Me.txtStandee = "N/A"
    pbHasRows = False
    ' Get and format the rest of the data
    ' Call p_Progbar("pGetProdDets", 10, "PCode")
    Call p_GetProdDets(psProductCode)                                                   ' Get the product details

    Call g_SetupIP("Products", 3, False)

    ' Call p_Progbar("End")

End Sub

Private Sub cmdSaveProduct_Click()
    ' Validation
    psValidate = p_Validate          ' Validate the entries

    If psValidate = "Yes" Then
        If NZ(Me.cboAllStatus.Value, 0) <> plAllSts Then Me.cboAllStatus.Value = Null  ' ?RWAS try out
        Call g_SetupIP("Products", 2, True)  ' To stop prior updating of fields, set up is on
        Me.txtPGID = CBA_lPromotion_ID
        ' If the GBDM date has been entered, update it's approval date
        If g_IsDate(Me.txtGBDMApprovalDate, True) = True And pbIsGBDMDate = False Then
            pbIsGBDMDate = True
            Me.txtGBDMApprovedDate = Now()
        End If
        ' If an add...
        If pbAddIP = True Then
            Call AST_WriteTable(Me.Controls, TABLE_FLDS_A, FIELD_PREFIX, MAIN_TABLE, plFrmID, plAuth, 0)
            CBA_lProduct_ID = g_DLookup("PD_ID", "L2_Products", "PD_ID>0", "PD_ID DESC", g_GetDB("ASYST"), 0)
            p_SendCModToDB

        Else   ' Is an Update
            Call AST_WriteTable(Me.Controls, TABLE_FLDS_U, FIELD_PREFIX, MAIN_TABLE, plFrmID, plAuth, CBA_lProduct_ID)
            p_SendCModToDB True

            psLockedIP = "No"
''            pbUpdIP = False
        End If
        pbUpdIP = False ''Was Removed in order to not force a requery of the SalesData
        ' Call p_Progbar("FillLstProducts", 10, "Call")
        'If Me.lstProducts.Column(2) <> Me.txtProductDesc Then
            Call p_FillLstProducts ' Refill the products in the list as the desc may have changed or it may have been added
        'End If
''        Call p_ResetProduct("Reset")

        'If pbUpdIP = False Then
        Call p_SelectProduct
        'pbUpdIP = False

        Call g_SetupIP("Products", , , True)
        pbAddIP = False
    ElseIf psValidate = "Reset" Then
''        Call cmdResetProduct_Click
    End If
End Sub
Private Function p_SendCModToDB(Optional ByVal b_Update As Boolean = False) As Boolean
Dim cMod As CBA_AST_Product
Dim sSQL As String
Dim DivNo As Long
Dim CN As ADODB.Connection
Dim RS As ADODB.Recordset

    If CBA_AST_getProdClassMod(cMod) = True Then
        'On Error GoTo Err_Routine
        CBA_ErrTag = ""
        Set CN = New ADODB.Connection
        Set RS = New ADODB.Recordset
        CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("ASYST") & ";"
        CBA_ErrTag = "SQL"
        If b_Update = True Then
          For DivNo = 501 To 509
              If DivNo = 508 Then DivNo = 509
              Set RS = New ADODB.Recordset
              sSQL = "UPDATE L3_ProductRegions "
              sSQL = sSQL & "Set PV_Unit_Cost = " & cMod.pUnit_CostDiv(DivNo) & Chr(10)
              sSQL = sSQL & ",PV_Supplier_Cost_Support = " & cMod.pSupplier_Cost_SupportDiv(DivNo) & Chr(10)
              sSQL = sSQL & ",PV_Prior_Sales = " & cMod.pcPrior_SalesDiv(DivNo) & Chr(10)
              sSQL = sSQL & ",PV_OrigEstSales = " & cMod.pcEstSalesDiv(DivNo) & Chr(10)
              sSQL = sSQL & ",PV_OrigCalcSales = " & cMod.pCalcSalesDiv(DivNo) & Chr(10)
              sSQL = sSQL & ",PV_UPSPW = " & cMod.pcUPSPWDiv(DivNo) & Chr(10)
              sSQL = sSQL & ",PV_Fill_Qty = " & cMod.pFill_QtyDiv(DivNo) & Chr(10)
              sSQL = sSQL & ",PV_Curr_Retail_Price = " & cMod.pCurr_Retail_PriceDiv(DivNo) & Chr(10)
              sSQL = sSQL & ",PV_Retail_Price = " & cMod.pRetail_PriceDiv(DivNo) & Chr(10)
              sSQL = sSQL & ",PV_OrigEstMultiplier = " & cMod.pEstMultiplierDiv(DivNo) & Chr(10)
              sSQL = sSQL & ",PV_LastUpd=" & g_GetSQLDate(Now, CBA_DMYHN) & Chr(10)
              sSQL = sSQL & "Where PV_PD_ID = " & cMod.plPDID & Chr(10)
              sSQL = sSQL & "And PV_Region = " & DivNo & Chr(10)
              'Debug.Print sSQL
              RS.Open sSQL, CN
          Next
        Else
            For DivNo = 501 To 509
                If DivNo = 508 Then DivNo = 509
                Set RS = New ADODB.Recordset
                sSQL = "INSERT INTO L3_ProductRegions(" & Chr(10)
                sSQL = sSQL & "PV_PD_ID," & Chr(10)
                sSQL = sSQL & "PV_Region," & Chr(10)
                sSQL = sSQL & "PV_Unit_Cost," & Chr(10)
                sSQL = sSQL & "PV_Supplier_Cost_Support," & Chr(10)
                sSQL = sSQL & "PV_Prior_Sales," & Chr(10)
                sSQL = sSQL & "PV_OrigEstSales," & Chr(10)
                sSQL = sSQL & "PV_OrigCalcSales," & Chr(10)
                sSQL = sSQL & "PV_UPSPW," & Chr(10)
                sSQL = sSQL & "PV_Fill_Qty," & Chr(10)
                sSQL = sSQL & "PV_Curr_Retail_Price," & Chr(10)
                sSQL = sSQL & "PV_Retail_Price," & Chr(10)
                sSQL = sSQL & "PV_OrigEstMultiplier)" & Chr(10)

                sSQL = sSQL & " VALUES (" & Chr(10)
                sSQL = sSQL & CBA_lProduct_ID & Chr(10)
                sSQL = sSQL & "," & DivNo & Chr(10)
                sSQL = sSQL & "," & cMod.pUnit_CostDiv(DivNo) & Chr(10)
                sSQL = sSQL & "," & cMod.pSupplier_Cost_SupportDiv(DivNo) & Chr(10)
                sSQL = sSQL & "," & cMod.pcPrior_SalesDiv(DivNo) & Chr(10)
                sSQL = sSQL & "," & cMod.pcEstSalesDiv(DivNo) & Chr(10)
                sSQL = sSQL & "," & cMod.pCalcSalesDiv(DivNo) & Chr(10)
                sSQL = sSQL & "," & cMod.pcUPSPWDiv(DivNo) & Chr(10)
                sSQL = sSQL & "," & cMod.pFill_QtyDiv(DivNo) & Chr(10)
                sSQL = sSQL & "," & cMod.pCurr_Retail_PriceDiv(DivNo) & Chr(10)
                sSQL = sSQL & "," & cMod.pRetail_PriceDiv(DivNo) & Chr(10)
                sSQL = sSQL & "," & cMod.pEstMultiplierDiv(DivNo) & Chr(10)
                sSQL = sSQL & ")"
                'Debug.Print sSQL
                RS.Open sSQL, CN
            Next
        End If
    End If
End Function
Private Sub cmdResetProduct_Click()
    ' Reset
    If pbAddIP Then CBA_lProduct_ID = 0
''    Call p_ResetProduct("Reset")
    Call g_SetupIP("Products", 2, True)
    Call p_SelectProduct
    pbAddIP = False: pbUpdIP = False
    Call g_SetupIP("Products", , , True)
End Sub
Private Sub lstProducts_Click()
    ' Handle any lst Products click
''    If g_SetupIP("Products") = True Then    ' If hasn't reset properly ... spent couple of hours trying to find out why
''        If Me.cmdSaveProduct.Visible = False And Me.lstProducts.ListIndex > -1 Then Call g_SetupIP("Products", , , True)
''    End If
    If g_SetupIP("Products") = False Then
        plAuth = CBA_lAuthority
''        ' If the product has been GBDM Approved and not an Admin then give Read-Only Authority
''        If Me.cboStatus <> 1 And g_IsDate(Me.txtGBDMApprovalDate, True) And plAuth <> 1 Then plAuth = 0
        ' Call p_Progbar("lstProducts_Clck", 10, "Rtne")
        Call g_SetupIP("Products", 2, True)  ' To stop prior updating of fields, set up is on
        ' Get the Products for the selected Promotion
        If Me.lstProducts.ListIndex > -1 Then
            ' Call p_Progbar("lstProducts_Clck", 10, ">-1")
            ' Call p_Progbar("End")
            pbSalesChanged = False
        Else
            CBA_lProduct_ID = 0
        End If
        ' Call p_Progbar("SetProducts", 10, "Call")
        Call p_SelectProduct
        Call g_SetupIP("Products", 2, False)
    End If
End Sub

Private Sub cboFreqItems_Change()
    Call p_UpdIP(Me.cboFreqItems)
End Sub

Private Sub chkTableMerch_Click()
    Call p_UpdIP(Me.chkTableMerch)
End Sub

Private Sub chkEndCapMerch_Click()
    Call p_UpdIP(Me.chkEndCapMerch)
End Sub

Private Sub lstProducts_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call mw_SetBoxHook(lstProducts)
End Sub

Private Sub lstSpecial_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call mw_SetBoxHook(lstSpecial)
End Sub

Private Sub lstTheme_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call mw_SetBoxHook(lstTheme)
End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub txtComplementary_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim sOld As String
    If Me.txtComplementary.BackColor = CBA_Grey Then Exit Sub
    sOld = Replace(Me.txtComplementary, " ", "_")
    sOld = Replace(sOld, vbCrLf, " ")
    CBA_strAldiMsg = "Complementary~" & sOld
''    varCal.sDate = g_FixDate(Me.txtOnSaleDate)
    CBA_AST_frm_AssocFld.Show vbModal
    If CBA_strAldiMsg > "" Then Me.txtComplementary = CBA_strAldiMsg
    If sOld <> Me.txtComplementary Then Call p_UpdIP(Me.txtComplementary)
End Sub

Private Sub txtCannabalised_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim sOld As String
    If Me.txtCannabalised.BackColor = CBA_Grey Then Exit Sub
    sOld = Replace(Me.txtCannabalised, " ", "_")
    sOld = Replace(sOld, vbCrLf, " ")
    CBA_strAldiMsg = "Cannibalised~" & sOld
''    varCal.sDate = g_FixDate(Me.txtOnSaleDate)
    CBA_AST_frm_AssocFld.Show vbModal
    If CBA_strAldiMsg > "" Then Me.txtCannabalised = CBA_strAldiMsg
    If sOld <> Me.txtCannabalised Then Call p_UpdIP(Me.txtCannabalised)
End Sub

Private Sub txtPressDates_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim sOld As String
    If Me.txtPressDates.BackColor = CBA_Grey Then Exit Sub
    sOld = Me.txtPressDates
    CBA_strAldiMsg = "Press_Dates~" & sOld
    varCal.sDate = g_FixDate(Me.txtOnSaleDate)
    CBA_AST_frm_AssocFld.Show vbModal
    If CBA_strAldiMsg > "" Then Me.txtPressDates = CBA_strAldiMsg
    If sOld <> Me.txtPressDates Then Call p_UpdIP(Me.txtPressDates)
End Sub

Private Sub txtOutOfHome_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim sOld As String
    If Me.txtOutOfHome.BackColor = CBA_Grey Then Exit Sub
    sOld = Me.txtOutOfHome
    CBA_strAldiMsg = "Out_Of_Home_Dates~" & sOld
    varCal.sDate = g_FixDate(Me.txtOnSaleDate)
    CBA_AST_frm_AssocFld.Show vbModal
    If CBA_strAldiMsg > "" Then Me.txtOutOfHome = CBA_strAldiMsg
    If sOld <> Me.txtOutOfHome Then Call p_UpdIP(Me.txtOutOfHome)
End Sub

Private Sub txtProductCode_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Exit Sub            ' @RWAST Is V2 functionality to change this over
    If Me.cmdSaveProduct.Visible = True Or (plAuth <> 1 And plAuth <> 2) Then Exit Sub
    varFldVars.lFldWidth = 120
    varFldVars.lFldHeight = 0
    varFldVars.lFrmLeft = g_PosForm(0, (varFldVars.lFldWidth * 2), 0, "Left")
    varFldVars.lFrmTop = g_PosForm(0, 0, 0, "Top")
    varFldVars.sHdg = "Select Promotion to transfer this product to"
    varFldVars.sSQL = "SELECT PG_ID, PG_Promo_Desc FROM [qry_L1_Promotions] WHERE PG_Status < 4 AND PG_ID <> " & CBA_lPromotion_ID & " ORDER BY Sts_Seq,PG_ID"
    varFldVars.sDB = "ASYST"
    varFldVars.bAllowNullOfField = False
    varFldVars.lCols = 2
    varFldVars.sType = "ComboBox"
    CBA_frmEntryField.Show vbModal

End Sub

Private Sub txtRegions_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim sOld As String
    If Me.txtRegions.BackColor = CBA_Grey Then Exit Sub
    sOld = Me.txtRegions
    CBA_strAldiMsg = "Regions~" & sOld
    CBA_AST_frm_AssocFld.Show vbModal
    If CBA_strAldiMsg > "" Then Me.txtRegions = CBA_strAldiMsg
    If sOld <> Me.txtRegions Then Call p_UpdIP(Me.txtRegions)
End Sub

Private Sub txtCalculatedSales_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call p_RegionClick
End Sub

Private Sub txtCurrRetailPrice_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call p_RegionClick
End Sub

Private Sub txtRetailPrice_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call p_RegionClick
End Sub

Private Sub txtSalesMultiplier_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call p_RegionClick
End Sub

Private Sub txtPriorSales_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call p_RegionClick
End Sub

Private Sub txtUPSPW_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call p_RegionClick
End Sub

Private Sub txtPriorUPSPW_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call p_RegionClick
End Sub

Private Sub txtExpectedSales_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call p_RegionClick
End Sub

Private Sub txtStandee_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim sOld As String
    If Me.txtStandee.BackColor = CBA_Grey Then Exit Sub
    sOld = Me.txtStandee
    CBA_strAldiMsg = "Standee_Dates~" & sOld
    varCal.sDate = g_FixDate(Me.txtOnSaleDate)
    CBA_AST_frm_AssocFld.Show vbModal
    If CBA_strAldiMsg > "" Then Me.txtStandee = CBA_strAldiMsg
    If sOld <> Me.txtStandee Then Call p_UpdIP(Me.txtStandee)
End Sub

Private Sub txtComments_Change()
    Call p_UpdIP(Me.txtComments)
End Sub

Private Sub txtFutureProdCode_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim sOld As String
    If Me.txtFutureProdCode.BackColor = CBA_Grey Then Exit Sub
    sOld = Me.txtFutureProdCode
    CBA_strAldiMsg = "Future_ProductCode~" & sOld
    CBA_AST_frm_AssocFld.Show vbModal
    If CBA_strAldiMsg > "" Then Me.txtFutureProdCode = CBA_strAldiMsg
    If sOld <> Me.txtFutureProdCode Then Call p_UpdIP(Me.txtFutureProdCode)
End Sub

Private Sub txtRadio_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim sOld As String
    If Me.txtRadio.BackColor = CBA_Grey Then Exit Sub
    sOld = Me.txtRadio
    CBA_strAldiMsg = "Radio_Dates~" & sOld
    varCal.sDate = g_FixDate(Me.txtOnSaleDate)
    CBA_AST_frm_AssocFld.Show vbModal
    If CBA_strAldiMsg > "" Then Me.txtRadio = CBA_strAldiMsg
    If sOld <> Me.txtRadio Then Call p_UpdIP(Me.txtRadio)
End Sub

Private Sub txtSupplierCostSupport_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call p_RegionClick
End Sub

Private Sub txtTV_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim sOld As String
    If Me.txtTV.BackColor = CBA_Grey Then Exit Sub
    sOld = Me.txtTV
    CBA_strAldiMsg = "TV_Dates~" & sOld
    varCal.sDate = g_FixDate(Me.txtOnSaleDate)
    CBA_AST_frm_AssocFld.Show vbModal
    If CBA_strAldiMsg > "" Then Me.txtTV = CBA_strAldiMsg
    If sOld <> Me.txtTV Then Call p_UpdIP(Me.txtTV)
End Sub

Private Sub txtUnitCost_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call p_RegionClick
End Sub

Private Sub txtOrigUnitCost_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call p_RegionClick
End Sub

Private Sub UserForm_Activate()
''    Set CBA_MW = New CBA_clsMouseWheel
''    Call CBA_MW.HookMouseWheelScroll(Me, Me.Name) 'Hook MouseWheel Scroll of Form and of all its controls
End Sub

Private Sub UserForm_Initialize()
    Dim sName As String
    Const LBLTOP = 36, LBLH = 18, FLDLEFT = 186, TXTTOP = 49, TXTH = 46, FLDW = 96
    ' Call p_Progbar("Initialise")
    ' To stop prior updating of upd fields, set up is on
    Call g_SetupIP("Products", 1, True, True)
    Me.StartUpPosition = 0
    Me.Top = Application.Top + (Application.Height / 2) - (Me.Height / 2)
    Me.Left = Application.Left + (Application.Width / 2) - (Me.Width / 2)
    Me.Width = 855.75
    ' Capture the Top, Width and Left positions
    Call g_PosForm(Me.Top, Me.Width, Me.Left, , True)


    'REMOVING THE REFRAMING OF THE MARKETING FRAME
    'Me.frmeMark.Width = 294
    
    
    Call AST_FillTagArrays("ID", 0, 0, "Clr")
''    ' Get the Authority level
''    CBA_lAuthority = AST_getUserASystAuthority(CBA_SetUser, True, True)
    plAuth = CBA_lAuthority
    ' Get the latest version - Test to see if it is the latest
    Me.lblVersion.Caption = CBA_getVersionStatus(g_GetDB("ASYST"), CBA_AST_Ver, "Super Saver Tool", "AST") & "-" & plAuth
    ' Call p_Progbar("Cont", 10)
    ' Fill the Promotions DDBox
    CBA_DBtoQuery = 1
    psSQL = "SELECT * FROM [qry_L1_Promotions] ORDER BY Sts_Seq,PG_ID"
    pbPassOk = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", "qry_L1_Promotions", g_GetDB("ASYST"), CBA_MSAccess, psSQL, 120, , , False)
    If pbPassOk = True Then
        Call AST_FillDDBox(Me.cboPromotionID, 8)
    End If
    ' Call p_Progbar("Cont", 10)

    ' Fill the Tier DDBox
    CBA_DBtoQuery = 1
    psSQL = "SELECT * FROM [L0_Tiers] ORDER BY TR_ID"
    pbPassOk = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", "L0_Tiers", g_GetDB("ASYST"), CBA_MSAccess, psSQL, 120, , , False)
    If pbPassOk = True Then
        Call AST_FillDDBox(Me.cboActualMarketingSupport, 2)
        Call AST_FillDDBox(Me.cboExpectedMarketingSupport, 2)
    End If
    ' Call p_Progbar("Cont", 10)

    ' Fill the Freq/Item DDBox
    CBA_DBtoQuery = 1
    psSQL = "SELECT * FROM [L0_Freq_Items] ORDER BY FI_ID"
    pbPassOk = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", "L0_Freq_Items", g_GetDB("ASYST"), CBA_MSAccess, psSQL, 120, , , False)
    If pbPassOk = True Then
        Call AST_FillDDBox(Me.cboFreqItems, 2)
    End If
    ' Call p_Progbar("Cont", 10)

   ' Fill the Sub-Basket DDBox
    CBA_DBtoQuery = 1
    psSQL = "SELECT SB_ID, SB_Classification FROM [L0_SubBasket] ORDER BY SB_ID"
    pbPassOk = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", "L0_SubBasket", g_GetDB("ASYST"), CBA_MSAccess, psSQL, 120, , , False)
    If pbPassOk = True Then
        Call AST_FillDDBox(Me.cboSubBasket, 2)
    End If
    CBA_DBtoQuery = 3
    ' Call p_Progbar("Cont", 10)

   ' Fill the Current Status DDBox
    CBA_DBtoQuery = 1
    psSQL = "SELECT Sts_ID,Sts_Desc FROM [L0_Statuses] WHERE Sts_ProductValid='Y' ORDER BY Sts_ID"
    pbPassOk = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", "L0_Statuses", g_GetDB("ASYST"), CBA_MSAccess, psSQL, 120, , , False)
    If pbPassOk = True Then
        Call AST_FillDDBox(Me.cboStatus, 2)
        Call AST_FillDDBox(Me.cboAllStatus, 2)
    End If
    CBA_DBtoQuery = 3
    ' Call p_Progbar("Cont", 10)

   ' Fill the SuperSaverType DDBox
    CBA_DBtoQuery = 1
    psSQL = "SELECT SST_ID,SST_Desc,SST_LeadWeeks FROM [L0_Super_Saver_Type] ORDER BY SST_ID"
    pbPassOk = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", "L0_Super_Saver_Type", g_GetDB("ASYST"), CBA_MSAccess, psSQL, 120, , , False)
    If pbPassOk = True Then
        Call AST_FillDDBox(Me.cboSuperSaverType, 3)
    End If
    CBA_DBtoQuery = 3
    ' Call p_Progbar("Cont", 10)

   ' Fill the Medium DDBox
    CBA_DBtoQuery = 1
    psSQL = "SELECT MD_Desc FROM [L0_Medium] WHERE MD_AST_Valid='Y' ORDER BY MD_ID"
    pbPassOk = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", "L0_Medium", g_GetDB("Gen"), CBA_MSAccess, psSQL, 120, , , False)
    If pbPassOk = True Then
        Call AST_FillDDBox(Me.cboMedium, 1)
    End If
    CBA_DBtoQuery = 3
    ' Call p_Progbar("Cont", 10)

    ' Get the Promo ID if it hasn't been gotton
    If CBA_lPromotion_ID = CBA_LongHiVal Then
        CBA_lPromotion_ID = g_DLookup("PG_ID", "L1_Promotions", "PG_ID>0", "PG_ID DESC", g_GetDB("ASYST"), 0)
    End If
    ' Call p_Progbar("Cont", 10)
    ' Set Weeks of Sale combo
    Call AST_WeeksOfSale(cboWeeksOfSale)

    ' Call p_Progbar("End")
    Me.lstSpecial.Locked = False
    Me.lstTheme.Locked = False
    Me.lstSpecial.BackColor = CBA_Grey
    Me.lstTheme.BackColor = CBA_Grey
    ' Set up the other AssocFld form fields
    sName = "OutOfHome"
    GoSub GSFields
    sName = "Radio"
    GoSub GSFields
    sName = "TV"
    GoSub GSFields
    sName = "PressDates"
    GoSub GSFields
    sName = "Standee"
    GoSub GSFields
    Me.cboMedium = Null
    ' If the Promotion_ID has been entered fill the Products list box with the values for it
    If CBA_lPromotion_ID > 0 Then
        Me.cboPromotionID.Value = CBA_lPromotion_ID
        ' Call p_Progbar("FillLstProducts", 10, "Call" & CBA_lPromotion_ID)
        Call p_FillLstProducts
    Else
        Call p_NullLists
    End If

    ' Set up the product if one has been specified
    ' Call p_Progbar("SetProducts", 10, "Call")
    Call p_SelectProduct

    ' Setup ended
    Call g_SetupIP("Products", 1, False)
    Exit Sub

GSFields:
'    Me("lbl" & SName).Width = FLDW
'    Me("lbl" & SName).Top = LBLTOP
'    Me("lbl" & SName).Height = LBLH
'    Me("lbl" & SName).Left = FLDLEFT
    Me("lbl" & sName).Visible = False
'    Me("txt" & SName).Width = FLDW
'    Me("txt" & SName).Top = TXTTOP
'    Me("txt" & SName).Height = TXTH
'    Me("txt" & SName).Left = FLDLEFT
    Me("txt" & sName).Visible = False
    Return

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Dim lReturn As Long

    If Me.cmdSaveProduct.Visible = True Then
        lReturn = MsgBox("Exit without saving?", vbYesNo + vbDefaultButton2, "Exit warning")
        If lReturn <> vbYes Then
                Cancel = True
            Exit Sub
        End If
    End If
    CBA_lProduct_ID = 0
    On Error Resume Next
''    Call CBA_MW.RemoveMouseWheelHook(Me.Name)
''    Set CBA_MW = Nothing
End Sub

Private Sub p_RegionClick()
    Dim sOld As String, sFields As String
    Static bRecursed As Boolean
    If bRecursed Then Exit Sub
    If pbAddIP Then
        CBA_lProduct_ID = CBA_LongHiVal
'        If MsgBox("The product needs to be saved before distributing to the regions - Save now?", vbYesNo) = vbYes Then
'            pbValidateAsOK = True
'            Call cmdSaveProduct_Click
'            pbValidateAsOK = False
'            If psValidate <> "Yes" Then Exit Sub
'        Else
'            Exit Sub
'        End If
    End If
    bRecursed = True
    Call AST_CrtProdRegionClass(CBA_lProduct_ID)
    pbHasRows = True
    sOld = Me.txtRegion
    CBA_lAuth = plAuth
    CBA_strAldiMsg = "Region1~" & sOld
    CBA_AST_frm_AssocFld.Show vbModal
    ' Global flag says regions have changed
    If CBA_bDataChg = True Then
''        Call AST_AveProdRegionRecs(CBA_lProduct_ID, CBA_strAldiMsg)
        If CBA_strAldiMsg > "" Then
            '''Me.txtRegion = CBA_strAldiMsg
''            CBA_DBtoQuery = 1
''            sFields = "PD_Regions,PD_Retail_Price,PD_Unit_Cost,PD_Supplier_Cost_Support,PD_Calculated_Sales,PD_Sales_Multiplier,PD_Curr_Retail_Price,PD_Prior_Sales,PD_Expected_Sales,PD_Retail_Discount,PD_UPSPW"
''            psSQL = "SELECT " & sFields & " FROM [" & MAIN_TABLE & "] WHERE PD_ID=" & CBA_lProduct_ID & ";"
''            pbPassOk = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", MAIN_TABLE, g_getDB("ASYST"), CBA_MSAccess, psSQL, 120, , , False)
''            If pbPassOk = True Then
''                pbNullFields = False: pbIsGBDMDate = False
''                Call AST_FillForm(Me.Controls, sFields, FIELD_PREFIX, plFrmID, plAuth, pbNullFields)
                ' Set the Visible and locked setting


            'If Me.cmdSaveProduct.Visible = False Then
                Call p_SetSaveVis("Save")
                pbUpdIP = True
                Call p_RegionDataUpd






            'End If

            If g_IsDate(Me.txtGBDMApprovalDate, True) = True Then pbIsGBDMDate = True
''            End If
        End If
        ''Call p_UpdIP(Me.txtRegion)   ' Has already been saved so no need to save again as part of this deal
    End If
    bRecursed = False
End Sub
Private Function p_RegionDataUpd() As Boolean
Dim cMod As CBA_AST_Product
    ' Re-insert the recalculated Class module values
    If CBA_ASyst.CBA_AST_getProdClassMod(cMod) = True Then
        Me.txtUPSPW = cMod.pUPSPW
        Me.txtCalculatedSales = cMod.pOrigCalcSales
        Me.txtUnitCost = cMod.pUnit_Cost
        Me.txtSupplierCostSupport = cMod.pSupplier_Cost_Support
        Me.txtRetailPrice = cMod.pRetail_Price
        Me.txtRetailDiscount = Round(((g_DivZero(Me.txtCurrRetailPrice - Me.txtRetailPrice, Me.txtCurrRetailPrice)) * 100), 1)
        Me.txtSalesMultiplier = Round(cMod.pEstMultiplier, 2)
        p_RegionDataUpd = True
    End If

End Function



Private Function p_Validate() As String
    Dim sOnSaleDate As String, sName As String, sMsg As String, sDate As String
    Dim sCrtDate As String, sEndDate As String
    On Error GoTo Err_Routine

    sOnSaleDate = IIf(g_IsDate(Me.txtOnSaleDate, True) = True, g_FixDate(Me.txtOnSaleDate), "")
    p_Validate = "Yes": sName = "": sMsg = ""
    If pbValidateAsOK = True Then Exit Function
    If Len(NZ(Me.txtProductDesc, "")) < 5 Then
        sMsg = "Product needs a title of more than 4 characters"
        sName = "txtProductDesc"
    ElseIf Val(Me.cboWeeksOfSale) < 1 Then
        sMsg = "Weeks of Sale must be 1 or greater"
        sName = "cboWeeksOfSale"
    ElseIf NZ(Me.cboStatus, 0) <> 1 And NZ(Me.cboStatus, 0) <> 2 And pbAddIP Then
        sMsg = "Only 'In Development' or 'Concept' can be selected for a new Product"
        sName = "cboStatus"
    ElseIf NZ(Me.cboActualMarketingSupport, "") = "" And p_ValidByUser("ActualMarketingSupport", plAuth, "IfDate") Then
        sMsg = "Actual Support hasn't been selected"
        sName = "cboActualMarketingSupport"
    ElseIf NZ(Me.cboFreqItems, "") = "" And p_ValidByUser("FreqItems", plAuth, "IfDate") Then
        sMsg = "Freq or Item/Basket hasn't been selected"
        sName = "cboFreqItems"
    ElseIf CDate(IIf(g_IsDate(Me.txtProductApprovalDate.Value, True), g_FixDate(Me.txtProductApprovalDate.Value), Date)) > Date Then
        sMsg = "Product Approval Date can't be a date in the future"
        sName = "txtProductApprovalDate"
    ElseIf CDate(IIf(g_IsDate(Me.txtGBDMApprovalDate.Value, True), g_FixDate(Me.txtGBDMApprovalDate.Value), Date)) > Date Then
        sMsg = "GBDM Approval Date can't be a date in the future"
        sName = "txtGBDMApprovalDate"
    ElseIf Me.cboPromotionID.Column(6, Me.cboPromotionID.ListIndex) = 3 And g_IsDate(Me.txtProductApprovalDate.Value, True) = False And p_ValidByUser("ProductApprovalDate", plAuth) Then
        sMsg = "Product Approval should have been obtained by this stage as the Promotion itself has GBDM Approval"
        sName = "txtProductApprovalDate"
    ElseIf Me.cboPromotionID.Column(6, Me.cboPromotionID.ListIndex) = 3 And g_IsDate(Me.txtGBDMApprovalDate.Value, True) = False And p_ValidByUser("GBDMApprovalDate", plAuth) Then
        sMsg = "GBDM Approval should have been obtained by this stage as the Promotion itself has GBDM Approval"
        sName = "txtGBDMApprovalDate"
    ElseIf g_IsDate(Me.txtGBDMApprovalDate, True) = False And CDate(g_FixDate(Me.txtApprovalReqDate)) < Now() _
                And Me.txtGBDMApprovalDate.BackColor <> CBA_Grey And p_ValidByUser("GBDMApprovalDate", plAuth) And Not pbAddIP Then
        sMsg = "GBDM Approval should be given at this stage"
        sName = "txtGBDMApprovalDate"
    ElseIf g_IsDate(Me.txtGBDMApprovalDate, True) = True And Me.cboStatus = 1 And p_ValidByUser("GBDMApprovalDate", plAuth) Then
        sMsg = "Status should reflect GBDM Approval"
        sName = "cboStatus"
    ElseIf g_IsDate(Me.txtGBDMApprovalDate, True) = False And Me.cboStatus = 3 And p_ValidByUser("GBDMApprovalDate", plAuth) Then
        sMsg = "Status should reflect 'In Development', 'Cancelled', 'Suspended' etc"
        sName = "cboStatus"
    ElseIf g_IsDate(Me.txtProductApprovalDate, True) = False And CDate(g_FixDate(Me.txtApprovalReqDate)) < Now() _
            And Me.txtProductApprovalDate.BackColor <> CBA_Grey And p_ValidByUser("ProductApprovalDate", plAuth) And Not pbAddIP Then
        sMsg = "Product Approval should have been obtained by this stage"
        sName = "txtProductApprovalDate"
    ElseIf pbHasRows = False And p_ValidByUser("ProductApprovalDate", plAuth, "IfDate") Then ' If no row data but has a ProdApproval date
        sMsg = "Region data should have been entered by this stage - Please click on any Cost, Retail or Sales field"
        sName = "txtProductApprovalDate"
    ElseIf pbAddIP = False Then
        sDate = g_FixDate(g_DLookup("PD_LastUpd", MAIN_TABLE, "PD_ID=" & CBA_lProduct_ID, "", g_GetDB("ASYST"), "01/01/1900"), CBA_DMYHN)
        If sDate <> g_FixDate(Me.txtLastUpd.Value, CBA_DMYHN) Then
        sMsg = "Record has been updated by another while this record was being prepared - Please press Reset and enter data again"
        sName = "txtLastUpd"
        End If
    End If
    ' If an error
    If sMsg > "" Then
        p_Validate = "No"
        If MsgBox(sMsg & vbCrLf & "Press OK to continue", vbOKOnly, "Validation Warning") = vbOK Then
            p_Validate = "Reset"
''        Else
            Me(sName).SetFocus
        End If
        Exit Function
    End If

Exit_Routine:

    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-frm_Products-p_Validate", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "ASYST", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Function

Private Function p_ValidByUser(sFieldTag As String, lAuth As Long, Optional IfDate_IfNotDate As String = "") As Boolean
    ' Will form a validation by user i.e. Marketing can't enter Buyer data (p_ValidByUser=False), but Admin can
    Const MRKTING As String = "ActualMarketingSupport,FreqItems"
    Const BUYING As String = "UnitCost,SupplierCostSupport,RetailPrice,SalesMultiplier,CalculatedSales"
    p_ValidByUser = False
    ' Check by Authority
    If lAuth = 1 Then
        GoSub GSDate
    ElseIf lAuth = 2 And InStr(1, BUYING, sFieldTag) > 0 Then
        GoSub GSDate
    ElseIf lAuth = 3 And InStr(1, MRKTING, sFieldTag) > 0 Then
        GoSub GSDate
    End If
    Exit Function
GSDate:
    ' Check by date or not
    If IfDate_IfNotDate = "IfDate" And g_IsDate(Me.txtProductApprovalDate, True) = True Then
        p_ValidByUser = True
    ElseIf IfDate_IfNotDate = "IfNotDate" And g_IsDate(Me.txtProductApprovalDate, True) = False Then
        p_ValidByUser = True
    ElseIf IfDate_IfNotDate = "" Then
        p_ValidByUser = True
    End If
''    Return

End Function

Private Sub p_FillLstProducts()
    ' Fill the Products combo box with product values for the Selected Promotion
    On Error GoTo Err_Routine

    ' Call p_Progbar("FillLstProducts", 10, "Rtne")
    CBA_DBtoQuery = 1
    If plAllSts = 0 Then
        psSQL = "SELECT PD_ID,PD_Product_Code,PD_Product_Desc FROM [" & MAIN_TABLE & "] WHERE PD_PG_ID=" & CBA_lPromotion_ID & " ORDER BY PD_ID;"
    Else
        psSQL = "SELECT PD_ID,PD_Product_Code,PD_Product_Desc FROM [" & MAIN_TABLE & "] WHERE PD_PG_ID=" & CBA_lPromotion_ID & " AND PD_Status = " & plAllSts & " ORDER BY PD_ID;"
    End If
    pbPassOk = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", MAIN_TABLE, g_GetDB("ASYST"), CBA_MSAccess, psSQL, 120, , , False)
    If pbPassOk = True Then
        Call AST_FillListBox(Me.lstProducts, 3, IIf(CBA_lProduct_ID > 0, CBA_lProduct_ID, ""))
''    Else
''        Call p_NullLists
''        CBA_lProduct_ID = 0
    End If
    CBA_DBtoQuery = 3

Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-p_FillFieldAuth", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "ASYST", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Sub

Private Sub p_SelectProduct()
    ' Get the selected Product for the promotion, or zero it out
    On Error GoTo Err_Routine
    Dim lProdID As Long

    CBA_ErrTag = ""
    pbHasRows = False
    ' Select whether the product has been selected or not
    ' Call p_Progbar("SetProducts", , "Rtne")
    If Me.lstProducts.ListIndex > -1 Then
        lProdID = NZ(Me.lstProducts.Column(0, Me.lstProducts.ListIndex), -1)
    Else
        lProdID = -1
    End If
    If lProdID > -1 Then
        CBA_lProduct_ID = lProdID
    ElseIf CBA_lProduct_ID > 0 Then
        Me.lstProducts.Value = CBA_lProduct_ID
    Else
        CBA_lProduct_ID = 0
    End If
    ' If a selected product
    If CBA_lProduct_ID > 0 Then
        ' Set the variables
        psProductCode = NZ(Me.lstProducts.Column(1, Me.lstProducts.ListIndex), 0)
        ' Fill the Product Form fields
        CBA_DBtoQuery = 1
        CBA_ErrTag = "SQL"
        psSQL = "SELECT " & TABLE_FLDS_F & " FROM [" & MAIN_TABLE & "] WHERE PD_ID=" & CBA_lProduct_ID & ";"
        pbPassOk = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", MAIN_TABLE, g_GetDB("ASYST"), CBA_MSAccess, psSQL, 120, , , False)
        If pbPassOk = True Then
            ' Call p_Progbar("ASTFillForm", 10, "Full")
            pbNullFields = False: pbIsGBDMDate = False
            Call AST_FillForm(Me.Controls, TABLE_FLDS_F, FIELD_PREFIX, plFrmID, plAuth, pbNullFields)
            If g_IsDate(Me.txtGBDMApprovalDate, True) = True Then pbIsGBDMDate = True
        End If
        CBA_ErrTag = ""
        CBA_DBtoQuery = 3
        ' Test to see if there are region records
        If g_DLookup("PV_ID", "L3_ProductRegions", "PV_PD_ID=" & CBA_lProduct_ID, "PV_PD_ID", g_GetDB("ASYST"), 0) > 0 Then pbHasRows = True
        ' Call p_Progbar("pGetProdDets", 10, "PCode")
        Call p_GetProdDets(psProductCode)                                   ' Get the product details
        ' Call p_Progbar("SetLockVis", 10)
        ' If the product has been GBDM Approved and not an Admin then give Read-Only Authority
        If Me.cboStatus <> 1 And g_IsDate(Me.txtGBDMApprovalDate, True) And plAuth <> 1 Then plAuth = 0
        Call AST_SetLockVis(Me.Controls, False, plFrmID, plAuth)           ' Set the visable/locked fields according to the authority
        Call p_SetSaveVis("Reset")
        Me.txtACatCGSCG.ControlTipText = Me.txtACGDesc
        Me.txtCGSCG.ControlTipText = Me.txtCGDesc

    Else
        psProductCode = 0
        ' Call p_Progbar("ASTFillForm", 10, "NULL")
        pbNullFields = True: pbIsGBDMDate = False
        Call AST_FillForm(Me.Controls, TABLE_FLDS_F, FIELD_PREFIX, plFrmID, plAuth, pbNullFields, ",PGID,") ' Null the fields on the form
        Call p_NullLists("Init")
        ' Call p_Progbar("SetSaveVis", 10, "Init")
        Call p_SetSaveVis("Init")                                                  ' Set the form visibility
    End If
    ' Finally set the lstProducts
    Me.lstProducts = Null
    If CBA_lProduct_ID > 0 Then Me.lstProducts = CBA_lProduct_ID
Exit_Routine:

    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-p_SelectProduct", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "ASYST", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Sub

Private Function p_GetProdDets(sProdCode As String, Optional bTestOnly As Boolean = False) As Boolean
    ' Apply the product details - return false if not found
    Dim psSQL As String, sElasticity_State As String, sConfidence_Level As String
    Dim scg As String, sSCg As String
    On Error GoTo Err_Routine

    p_GetProdDets = False
    ' Call p_Progbar("pGetProdDets", 10, "Rtne-Desc-CG-BD-Test?")
    ' Get the Product details - PC,PCDesc,CG,CGDesc,SCG,SCGDesc, BD, GBD, Prod Desc etc  - Note Changed Firstname for EmpSign (initials) and didn't include Name
                        ' 0           1                        2       3                           4             5
    psSQL = "SELECT p.ProductCode, p.Description AS ProdDesc,cg.CGNo, cg.Description AS CGDesc, scg.SCGNo, scg.Description AS SCGDesc, " & _
            "       emp.EmpSign AS BDInits, emp.Name as BDLName, em.EmpSign AS GBDInits, em.Name AS GBDLName, p.TaxID, " & _
            "       en.CatNo,en.CGNo AS ACG,en.SCGNo AS ASCG,en.Category AS ACat,en.CommodityGroup AS ACG,en.SubCommodityGroup AS ASCG " & _
            "FROM cbis599p.dbo.Product p " & _
            "LEFT JOIN cbis599p.dbo.tf_acgmap() en ON en.ACGEntityID = p.ACGEntityID " & _
            "LEFT JOIN cbis599p.dbo.CommodityGroup cg on cg.CGNo = p.CGNo " & _
            "LEFT JOIN cbis599p.dbo.SubCommodityGroup scg on scg.CGNo = p.cgno and scg.SCGNo = p.SCGNo " & _
            "LEFT JOIN cbis599p.dbo.EMPLOYEE emp on emp.EmpNo = p.Empno " & _
            "LEFT JOIN cbis599p.dbo.EMPLOYEE em on emp.EmpNo_Grp = em.EmpNo " & _
            "       WHERE p.ProductCode=" & sProdCode & "; "
    CBA_DBtoQuery = 599
    pbPassOk = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", "CBIS_QUERY", CBA_BasicFunctions.TranslateServerName(CBA_DBtoQuery, Date), "SQLNCLI10", psSQL, 120, , , False)  'Runs CBA_DB_Connection module to create connection to dtabase and run query
    If pbPassOk = True Then
        p_GetProdDets = True
        If g_Empty(Me.txtProductDesc.Value) = 0 Then Me.txtProductDesc.Value = CBA_CBISarr(1, 0)
        scg = CBA_CBISarr(2, 0) & IIf(g_Empty(CBA_CBISarr(3, 0)) > 0, "-" & CBA_CBISarr(3, 0), "")
        sSCg = CBA_CBISarr(4, 0) & IIf(g_Empty(CBA_CBISarr(5, 0)) > 0, "-" & CBA_CBISarr(5, 0), "")
        Me.txtCGSCG.Value = CBA_CBISarr(2, 0) & IIf(g_Empty(CBA_CBISarr(4, 0)) > 0, "-" & CBA_CBISarr(4, 0), "")
        Me.txtCGDesc.Value = scg & IIf(g_Empty(sSCg) > 0, "-" & sSCg, "")
        ' ACat data
        Me.txtACatCGSCG.Value = Format(CBA_CBISarr(11, 0), "00") & "-" & Format(CBA_CBISarr(12, 0), "00") & "-" & Format(CBA_CBISarr(13, 0), "00")
        Me.txtACGDesc.Value = Format(CBA_CBISarr(11, 0), "00") & "-" & CBA_CBISarr(14, 0) & "/" & Format(CBA_CBISarr(12, 0), "00") & "-" & CBA_CBISarr(15, 0) & "/" & Format(CBA_CBISarr(13, 0), "00") & "-" & CBA_CBISarr(16, 0)
        Me.txtBD.Value = CBA_CBISarr(6, 0)              ' & ", " & CBA_CBISarr(6, 0)
        Me.txtGBD.Value = CBA_CBISarr(8, 0)             ' & ", " & CBA_CBISarr(8, 0)
        plCGno = NZ(CBA_CBISarr(2, 0), 0): plSCGNo = NZ(CBA_CBISarr(4, 0), 0)
        Me.txtGST = IIf(NZ(CBA_CBISarr(10, 0), 1) = 2, "Y", "N")
    End If

    ' If Product not found or just a test, hop out
    If p_GetProdDets = False Or bTestOnly Then Exit Function
    ' Get the SubBasket and SuperSaverType
    ' Call p_Progbar("pGetProdDets", 10, "SBskt-Elast+Clrs")
    CBA_DBtoQuery = 1
    psSQL = "SELECT SB_ID, SST_ID, SST_LeadWeeks FROM qry_L2_ProductDets WHERE CG_CGNo = " & plCGno & ";"
    pbPassOk = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", "qry_L2_ProductDets", g_GetDB("ASYST"), CBA_MSAccess, psSQL, 120, , , False)
    If pbPassOk = True Then
        Me.cboSubBasket.Value = NZ(CBA_ABIarr(0, 0), "")
        If pbAddIP Or Me.cboSuperSaverType.ListIndex = -1 Then Me.cboSuperSaverType.Value = NZ(CBA_ABIarr(1, 0), "")
    Else
        Me.cboSubBasket.Value = ""
        Me.cboSuperSaverType.Value = ""
        ''plLeadWeeks = 0
    End If
    Me.txtApprovalReqDate.Value = g_FixDate(CDate(g_FixDate(Me.txtOnSaleDate)) - p_LeadWeeks, CBA_D3DMY)
    ''' Call p_Progbar("Cont", 10)

    ' Get Price Elasticity details
    CBA_DBtoQuery = 1
    psSQL = "SELECT CGNo, SCGNo, Elastic, Avg_Price, Elasticity_State, Confidence_Level FROM L0_Price_Elasticity_Sheet WHERE CGNo = " & plCGno & " AND SCGNo=" & plSCGNo & ";"
    pbPassOk = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", "L0_Price_Elasticity_Sheet", g_GetDB("ASYST"), CBA_MSAccess, psSQL, 120, , , False)
    If pbPassOk = True Then
        pdblElastic = NZ(CBA_ABIarr(2, 0), 0)
        pdblAvg_Price = NZ(CBA_ABIarr(3, 0), 0)
        psElasticity_State = NZ(CBA_ABIarr(4, 0), "Low")
        psConfidence_Level = NZ(CBA_ABIarr(5, 0), "Low")
    Else
        pdblElastic = 0
        pdblAvg_Price = 0
        psElasticity_State = "Low"
        psConfidence_Level = "Low"
    End If
    ''' Call p_Progbar("Cont", 10)

    ' Set the Product Elasticity Flag
    sElasticity_State = psElasticity_State
    Select Case sElasticity_State
    Case "High"
        Me.txtElasticityFlag.Value = sElasticity_State
        Me.txtElasticityFlag.BackColor = CBA_Green
    Case "Medium"
        Me.txtElasticityFlag.Value = sElasticity_State
        Me.txtElasticityFlag.BackColor = vbYellow
    Case "Low"
        Me.txtElasticityFlag.Value = sElasticity_State
        Me.txtElasticityFlag.BackColor = CBA_Red
    End Select
    ' Set the Product Elasticity Confidence Flag
    sConfidence_Level = psConfidence_Level
    Select Case sConfidence_Level
    Case "High", "H"
        Me.txtElasticityConfidenceFlag.Value = "High"
        Me.txtElasticityConfidenceFlag.BackColor = CBA_Green
    Case "Medium", "M"
        Me.txtElasticityConfidenceFlag.Value = "Medium"
        Me.txtElasticityConfidenceFlag.BackColor = vbYellow
    Case "Low", "L"
        Me.txtElasticityConfidenceFlag.Value = "Low"
        Me.txtElasticityConfidenceFlag.BackColor = CBA_Red
    End Select
    ''' Call p_Progbar("Cont", 10)
    ' Get the prior sales data and extrapolate the new
    ' Call p_Progbar("pGetSalesData", 10, "Call")
    Call p_GetSalesData
    ' Call p_Progbar("Cont", 10)
Exit_Routine:

    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-p_GetProdDets", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "ASYST", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Function

Sub p_SetSaveVis(Add_Reset_Init As String)

    ' Set the Visible / lock status of the fields as they change through updating??/

    Dim bAddVis As Boolean, bSaveVis As Boolean, bResetVis As Boolean, bLockOR As Boolean, bLocked As Boolean, bAVis As Boolean
    Dim lAuth As Long
''                psLockedIP = "No"
''            pbUpdIP = False
    ' Call p_Progbar("SetSaveVis", 10, "Rtne")
    lAuth = plAuth
    ' if bLockOR = true will lock and grey everything
    Select Case Add_Reset_Init
    Case "Add"                                                  ' After an Add
        bAddVis = False: bSaveVis = True: bResetVis = bSaveVis: bLockOR = False: bAVis = True
    Case "Save"                                                 ' After a Save has been initiated...
        bAddVis = False: bSaveVis = True: bResetVis = bSaveVis: bLockOR = False: bAVis = True
    Case "Reset"                                                ' For a Reset
        bAddVis = True: bSaveVis = False: bResetVis = bSaveVis: bLockOR = False: bAVis = True
    Case "Init"                                                 ' Upon first entry - if there are Products existing
        bAddVis = True: bSaveVis = False: bResetVis = bSaveVis: bLockOR = True: bAVis = False
        If Me.cboPromotionID.ListIndex = -1 Then bAddVis = False
        Me.cboMedium = Null
        Call cboMedium_Change  ' Nulling it will call the routine????
        lAuth = 0
    Case Else
        MsgBox "no data"
    End Select
''    ' For fields that need to be displayed on a save but not add
''    If pbAddIP Then bAVis = False
    ' If the Promotion is at GBDM level doen't allow an add unless Admin
    If Me.cboPromotionID.ListIndex > -1 Then
        If Me.cboPromotionID.Column(6, Me.cboPromotionID.ListIndex) = 3 And lAuth <> 1 Then
            bAddVis = False
        End If
    End If

    If Add_Reset_Init <> "Save" Then
        ' Call p_Progbar("SetLockVis", 10)
        ' Set the Lock / Visibility of the fields on the form
        Call AST_SetLockVis(Me.Controls, bLockOR, plFrmID, lAuth)
    End If
    ' Set the  Visibility of the Cmd Keys and ListBoxes on the form
    bLocked = bSaveVis
    ' If there has been a Lock error then don't allow a Save
    If psLockedIP = "AlreadyLocked" Then bSaveVis = False
    ' Call p_Progbar("SetCmdVis", 10)
    Call AST_SetCmdVis(Me.cmdSaveProduct, bSaveVis, plFrmID, plAuth)
    Call AST_SetCmdVis(Me.cmdAddProduct, bAddVis, plFrmID, plAuth)
    Call AST_SetCmdVis(Me.cmdResetProduct, bResetVis, plFrmID, plAuth)

    Me.lstProducts.Locked = bLocked
''    Me.lstProducts.BackColor = IIf(Me.lstProducts.Locked = True, CBA_Grey, CBA_White)
    Me.cboPromotionID.Locked = bLocked
    Me.cboPromotionID.BackColor = IIf(Me.cboPromotionID.Locked = True, CBA_Grey, CBA_White)
''    If Me.txtSalesMultiplier.BackColor = CBA_Grey Then
''        Me.cmdRegion.Visible = False
''    Else
''        Me.cmdRegion.Visible = True
''    End If
    Call p_TestApprovals
    If bLockOR Or pbAddIP Then Exit Sub

End Sub

Sub p_UpdIP(cActCtl As Control, Optional bIsDate As Boolean = False, Optional bFormat As Boolean = False, Optional bClr As Boolean = True)
    ' Flag the record as being updated
    Static bRecursed As Boolean
    Dim sFormat As String, sClr As String
    If bRecursed = True Then Exit Sub

    Call CBA_ProcI("s-p_UpdIP", 3)
    bRecursed = True
    ' If just a format...
    If bFormat Then
        sFormat = AST_FillTagArrays(cActCtl.Tag, plFrmID, plAuth, "Format")
        cActCtl.Value = Format(cActCtl.Value, sFormat)
    End If
    ' If a Clr check...
    If bClr Then
        ' Check to see if the colour needs changing
        sClr = AST_FillTagArrays(cActCtl.Tag, plFrmID, plAuth, "Clr")
        Call AST_FillYellow(cActCtl, sClr)
    End If
    ' If an actual update is going to take place
    If g_SetupIP("Products") = False Then
        ' Call p_Progbar("p_UpdIP", 10, cActCtl.Name)
        If pbUpdIP = False Then
            pbUpdIP = True
            ' Call p_Progbar("Lock_Table", 10, "Lock")
            psLockedIP = AST_Lock_Table("ASYST", "PD_", MAIN_TABLE, "PD_ID=" & CBA_lProduct_ID, CBA_lAuthority, "Get_SetY")
        End If
'        ' Check to see if the colour needs changing
'        sClr = AST_FillTagArrays(cActCtl.Tag, plFrmID, plAuth, "Clr")
'        Call AST_FillYellow(cActCtl, sClr)
        ' If a date, check to see if it has changed, if not then hop out of update
        If bIsDate = True And varCal.bCalValReturned = False Then GoTo GTLeave
        ' Will append the 'IsUpdated' flag "~" to tag and bring back format
        sFormat = AST_getFieldTag(cActCtl, plFrmID, plAuth, , "ApdUpd")
        If Not bIsDate And bFormat Then
            cActCtl.Value = Format(cActCtl.Value, sFormat)
        End If
        ' Set the Visible and locked setting
        If Me.cmdSaveProduct.Visible = False Then
            Call p_SetSaveVis("Save")
        End If
    End If
GTLeave:
    bRecursed = False
End Sub

Private Sub p_GetSalesData(Optional bReCalculate As Boolean = False)
    ' Get the prior sales data and the comparision value (* the expected multiplier)
    Dim sOnSaleDate As String, sEndDate As String, sOnSaleDate12 As String, sOnSaleDate1 As String, sEndDate12 As String, sngMult As Single, lIdx As Long, lUPSPW As Long, sngRetail As Single
    Dim lStoreCnt As Long, lPriorStoreCnt As Long, lPUPSPW As Long

    On Error GoTo Err_Routine
    CBA_ErrTag = ""
    sOnSaleDate = g_FixDate(Me.txtOnSaleDate.Value)
    sEndDate = g_FixDate(Me.txtEndDate.Value)
    sOnSaleDate12 = DateAdd("m", -12, g_FixDate(Me.txtOnSaleDate.Value))
    sOnSaleDate1 = DateAdd("d", -1, g_FixDate(Me.txtOnSaleDate.Value))
    sEndDate12 = DateAdd("m", -12, g_FixDate(Me.txtEndDate.Value))
    If psProductCode = "" Then psProductCode = "0"
    ' If has no rows
    If (pbHasRows = False And NZ(Me.txtUnitCost, 0) = 0 And NZ(Me.txtCurrRetailPrice, 0) = 0) Or bReCalculate = True Then
        sngMult = NZ(Me.txtSalesMultiplier.Value, 1)
        ' Call p_Progbar("GetSalesData", 10, "SalesVal-Specials-Themes")
        If g_IsDate(sOnSaleDate12) And g_IsDate(sEndDate12) Then
            Me.txtPriorSales.Value = AST_getCompareSalesCBis(sOnSaleDate12, sEndDate12, psProductCode, "Prior")
            Me.txtExpectedSales.Value = AST_getCompareSalesCBis(sOnSaleDate12, sEndDate12, psProductCode, "Now", , sngMult)
        Else
            Me.txtPriorSales.Value = "0"
            Me.txtExpectedSales.Value = "0"
        End If
        Me.txtCalculatedSales = Format(NZ(Me.txtPriorSales, 0) * sngMult, "#,000")
        Me.cboExpectedMarketingSupport.Value = AST_GetEstTierData(psProductCode, plCGno)

        ', lUPSPW, sngRetail


        CBA_SQL_Queries.CBA_GenPullSQL "CBA_AST_ALLStoreNos", DateAdd("m", -12, sEndDate)
        lStoreCnt = CBA_CBISarr(0, 0)
        CBA_SQL_Queries.CBA_GenPullSQL "CBA_AST_ALLStoreNos", sEndDate
        lPriorStoreCnt = CBA_CBISarr(0, 0)
        lPUPSPW = Me.txtPriorSales.Value / lPriorStoreCnt / Me.cboWeeksOfSale
        lUPSPW = Me.txtExpectedSales.Value / lStoreCnt / Me.cboWeeksOfSale
        Me.txtPriorUPSPW.Value = Format(lPUPSPW, "#0")
        Me.txtUPSPW.Value = Format(lUPSPW, "#0")
        ' Fill the Current Retail Value and the Costs for the item
        CBA_DBtoQuery = 599
        pbPassOk = CBA_SQL_Queries.CBA_GenPullSQL("CBIS_Retail and Cost", Date, , psProductCode)
        If pbPassOk = False Then
            pbPassOk = CBA_SQL_Queries.CBA_GenPullSQL("CBIS_Retail and Cost2", Date, , psProductCode)
        End If
        If pbPassOk Then
            Me.txtOrigUnitCost = Format(NZ(CBA_CBISarr(2, 0), 0), "0.00")
            Me.txtUnitCost = Format(NZ(CBA_CBISarr(2, 0), 0), "0.00")
            Me.txtCurrRetailPrice = Format(NZ(CBA_CBISarr(1, 0), 0), "0.00")
            Me.txtRetailPrice = Format(NZ(CBA_CBISarr(1, 0), 0), "0.00")
            Me.txtRetailDiscount = (Me.txtCurrRetailPrice - Me.txtRetailPrice) / Me.txtCurrRetailPrice
        Else
            Me.txtOrigUnitCost = 0
            Me.txtUnitCost = 0
            Me.txtCurrRetailPrice = 0
            Me.txtRetailPrice = 0
        End If
    End If
    Call p_NullLists("Init")
    'Create Region Class Module
    If CBA_AST_DestroyClassModule = True Then
        'l_Product_ID
        Call AST_CrtProdRegionClass(IIf(pbAddIP, CBA_LongHiVal, CBA_lProduct_ID), True)
    End If
    ' Fill the Specials list box
    If g_IsDate(sOnSaleDate) And g_IsDate(sEndDate) Then
        CBA_DBtoQuery = 599
        pbPassOk = CBA_SQL_Queries.CBA_GenPullSQL("CBIS_Specials", sOnSaleDate, sEndDate, psProductCode)
        If pbPassOk Then
            CBA_ABIarr = CBA_CBISarr
            For lIdx = 0 To UBound(CBA_ABIarr, 2)
                If NZ(CBA_ABIarr(0, lIdx), "") = "" Then CBA_ABIarr(0, lIdx) = "NoCode"
                CBA_ABIarr(1, lIdx) = g_RevDate(CBA_ABIarr(1, lIdx), "yyyy-mm-dd", "dd/mm/yy")
            Next
            Call AST_FillListBox(Me.lstSpecial, 4)
        End If
    End If
    ' Fill the Themes list box
    If g_IsDate(sOnSaleDate) And g_IsDate(sEndDate) Then
        CBA_DBtoQuery = 599
        pbPassOk = CBA_SQL_Queries.CBA_GenPullSQL("CBIS_Themes", sOnSaleDate, sEndDate)
        If pbPassOk Then
            CBA_ABIarr = CBA_CBISarr
            For lIdx = 0 To UBound(CBA_ABIarr, 2)
''                If NZ(CBA_ABIarr(0, lIdx), "") = "" Then CBA_ABIarr(0, lIdx) = "NoCode"
                CBA_ABIarr(1, lIdx) = g_RevDate(CBA_ABIarr(1, lIdx), "yyyy-mm-dd", "dd/mm/yy")
            Next
            Call AST_FillListBox(Me.lstTheme, 2)
        End If
    End If

Exit_Routine:
    On Error Resume Next
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("p_GetSalesData", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0) & "-" & CBA_ErrTag
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "AST", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Sub

Private Function p_LeadWeeks() As Long
    ' Will deliver back the lead time in weeks
    If Me.cboSuperSaverType.ListIndex > -1 Then
        p_LeadWeeks = Me.cboSuperSaverType.Column(2) * 7
    Else
        p_LeadWeeks = 0
    End If
End Function

''Private Sub p_ResetProduct(sReset_Init As String)
''    ' Will reset the product...
''    ' Call p_Progbar("ResetProduct", 10, "Rtne-" & sReset_Init)
''    ' Reset
''    If sReset_Init = "Init" Then
''        ' Call p_Progbar("ASTFillForm", 10, "NULL")
''        pbNullFields = True
''        Call AST_FillForm(Me.Controls, TABLE_FLDS_F, FIELD_PREFIX, plFrmID, plAuth, pbNullFields)          ' Null the fields on the form
''        ' Call p_Progbar("SetSaveVis", 10, "Init")
''        Call p_SetSaveVis("Init")
''        Me.lstProducts.Value = Null
''    Else
''        If pbUpdIP = True Then
''            ' Call p_Progbar("Lock_Table", 10, "UnLock")
''            psLockedIP = AST_Lock_Table("ASYST", "PD_", MAIN_TABLE, "PD_ID=" & CBA_lProduct_ID, plAuth, "SetN")
''            ' Call p_Progbar("SetSaveVis", 10, "Init")
''            Call p_SetSaveVis("Reset")
''            ' Call p_Progbar("lstProducts_Clck", 10, "Null")
''            Me.lstProducts.Value = Null
''            Call lstProducts_Click                                         ' Reset the current or new product
''            pbUpdIP = False
''        Else
''            ' Call p_Progbar("lstProducts_Clck", 10, "Reset?")
''            Call lstProducts_Click                                         ' Reset the current or new product
''            ' Call p_Progbar("SetSaveVis", 10, "Reset")
''            Call p_SetSaveVis("Reset")                                     ' Set the form visibility
''        End If
''    End If
''End Sub

Private Sub p_TestApprovals()
    Static lColour As Long

    ' Flag the GBDM date as overdue if it is past the lead time (Approval Req Date)
    If g_IsDate(Me.txtGBDMApprovalDate, True) = False And CDate(g_FixDate(Me.txtApprovalReqDate)) < Now() And pbNullFields = False Then
        txtGBDMApprovalDate.BackColor = vbRed
    Else
        If txtGBDMApprovalDate.BackColor <> vbRed Then
            lColour = txtGBDMApprovalDate.BackColor
        Else
            If lColour = 0 Then lColour = CBA_White
            txtGBDMApprovalDate.BackColor = lColour
        End If
    End If
    ' Flag the Product Approval Date as overdue if it is past the lead time (Approval Req Date)
    If g_IsDate(Me.txtProductApprovalDate, True) = False And CDate(g_FixDate(Me.txtApprovalReqDate)) < Now() And pbNullFields = False Then
        txtProductApprovalDate.BackColor = vbRed
    Else
        If txtProductApprovalDate.BackColor <> vbRed Then
            lColour = txtProductApprovalDate.BackColor
        Else
            If lColour = 0 Then lColour = CBA_White
            txtProductApprovalDate.BackColor = lColour
        End If
    End If

End Sub

Private Sub p_NullLists(Optional InitAll As String = "InitAll")
    If InitAll = "InitAll" Then Me.lstProducts.Clear
    Me.lstSpecial.Clear
    Me.lstTheme.Clear
End Sub

Private Sub p_Progbar(Optional ByVal Init_Cont As String = "Cont", Optional lInc As Long = 1, Optional sType As String)
    Static bHasRun As Boolean, lSecs As Long, sMsg As String, sMsg2 As String, sLast As String, sSuff As String, sPath As String
    Dim sTemp As String
    If Init_Cont = "Init" Then
            Exit Sub
''        frmSplashWait.Show vbModeless
''        Me.ProgBar.Width = lInc
''        stcWidth = lInc
''        Me.ProgBar.Visible = True
    ElseIf Init_Cont = "End" Or Init_Cont = "Cont" Then
        Exit Sub
''        Me.ProgBar.Visible = False
    Else
''        stcWidth = stcWidth + lInc
''        If stcWidth > Me.ProgBar.Max Then stcWidth = Me.ProgBar.Max
''        Me.ProgBar.Width = stcWidth
''        Unload frmSplashWait
    End If
    If bHasRun = False Then lSecs = 20
Redo:
    If lSecs > 2 Then
        If sMsg > "" Then
            If sPath = "" Then
                sPath = g_GetDB("ASYST", True)
                sPath = Replace(sPath, "ASYSTErrors.txt", "ASYSTLogs.txt")
            End If
            sMsg = "": sSuff = ">>>>>": sLast = ""
        End If
        Call g_SecsElapsed("Init")
        lSecs = 0
        GoTo Redo
    Else
        lSecs = g_SecsElapsed("")
    End If

''    If sLast = Init_Cont Then
''        sTemp = Init_Cont & "**"
''    Else
    If InStr(1, sMsg2, Init_Cont) > 0 Then
        sTemp = Init_Cont & "*"
    Else
        sTemp = Init_Cont
    End If
    sMsg2 = sMsg
    ' Accum the Msg string
    sMsg = sMsg & sSuff & sTemp & IIf(sType > "", "(" & sType & ")", "")
    bHasRun = True: sSuff = ";": sLast = Init_Cont

End Sub

Private Sub cboActualMarketingSupport_Change()
    Call p_UpdIP(Me.cboActualMarketingSupport)
End Sub

Private Sub cboExpectedMarketingSupport_Change()
    Call p_UpdIP(Me.cboExpectedMarketingSupport)
End Sub

Private Sub cboStatus_Change()
    Call p_UpdIP(Me.cboStatus)
End Sub

Private Sub txtTheme_Change()
    Call p_UpdIP(Me.txtTheme)
End Sub

Private Sub txtProductDesc_Change()
    Call p_UpdIP(Me.txtProductDesc)
End Sub

Private Sub txtCoverDate_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim bUseCtlDate As Boolean
    If Me.txtCoverDate.BackColor = CBA_Grey Then Exit Sub
    If g_IsDate(Me.txtCoverDate, True) Then
        bUseCtlDate = True
    Else
        bUseCtlDate = False
        varCal.sDate = g_FixDate(Me.txtOnSaleDate)
    End If
    Call CBA_BT_CalendarShow(Me.Top, Me.Left, Me.txtCoverDate, CBA_D3DMY, , bUseCtlDate)
    Call p_UpdIP(Me.txtCoverDate, True)
End Sub

Private Sub txtOnSaleDate_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Static b_recurse As Boolean
Dim yn As Long
Dim c_mod As CBA_AST_Product

    If Me.txtOnSaleDate.BackColor = CBA_Grey Then Exit Sub
    Call CBA_BT_CalendarShow(Me.Top, Me.Left, Me.txtOnSaleDate, CBA_D3DMY, False)
    If varCal.bCalValReturned = False Then Exit Sub

    If g_SetupIP("Products") = False And b_recurse = False Then
        yn = MsgBox("By changing the Start Date, all the sales data will be recalculated and will need to be reentered" & Chr(10) & Chr(10) & "Do you wish to proceed?", vbYesNo)
        If yn = 6 Then
            Call AST_setEndDate(Me.Name, Me.Controls, Me.txtOnSaleDate, CBA_D3DMY)
            Me.txtApprovalReqDate.Value = g_FixDate(CDate(g_FixDate(Me.txtOnSaleDate)) - p_LeadWeeks, CBA_D3DMY)
            Call p_GetSalesData(True)
            Call p_UpdIP(Me.txtOnSaleDate, True)
            If Not pbAddIP And pbHasRows = True Then
                If CBA_AST_DestroyClassModule = True Then
                    Call AST_CrtProdRegionClass(Me.lstProducts.Column(0), True)
                End If
            End If
        Else
            b_recurse = True
            txtOnSaleDate = Format(CDate(Trim(Left(Mid(txtOnSaleDate.Tag, InStr(1, txtOnSaleDate.Tag, " ") + 1, 19), Len(Mid(txtOnSaleDate.Tag, InStr(1, txtOnSaleDate.Tag, " ") + 1, 19)) - 1))), CBA_D3DMY)
            b_recurse = False
        End If
    End If


End Sub


Private Sub cboWeeksOfSale_Change()
Dim yn As Long
Dim WoS As String
Dim c_mod As CBA_AST_Product
Static b_recurse As Boolean
    If g_SetupIP("Products") = False And b_recurse = False Then
        yn = MsgBox("By changing the weeks of sale, all the sales data will be recalculated and will need to be reentered" & Chr(10) & Chr(10) & "Do you wish to proceed?", vbYesNo)
        If yn = 6 Then
            Call AST_setEndDate(Me.Name, Me.Controls, Me.cboWeeksOfSale, CBA_D3DMY)
            Call p_GetSalesData(True)
            Call p_UpdIP(Me.cboWeeksOfSale)
            If Not pbAddIP And pbHasRows = True Then
                If CBA_AST_DestroyClassModule = True Then
                    Call AST_CrtProdRegionClass(Me.lstProducts.Column(0), True)
                End If
            End If

        ElseIf yn = 7 Then
            WoS = Trim(Val(Mid(CBA_AST_frm_Products.cboWeeksOfSale.Tag, InStr(1, CBA_AST_frm_Products.cboWeeksOfSale.Tag, "~") + 1, 9)))
            b_recurse = True
            Me.cboWeeksOfSale.Value = WoS
            b_recurse = False
        End If
    End If
End Sub

Private Sub txtProductApprovalDate_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ''If Me.txtProductApprovalDate.BackColor = CBA_Grey Then Exit Sub
    If AST_FillTagArrays(Me.txtProductApprovalDate.Tag, plFrmID, plAuth, "Lock") = "lock" Then Exit Sub
    If g_IsDate(txtProductApprovalDate.Value, True) = False Then
        If AST_UpdGBDMDate(g_FixDate(txtProductApprovalDate.Value), CBA_lPromotion_ID, True, CStr(txtID.Value), "Fatal", "Product Approval Date") <> "True" Then Exit Sub
    End If
    Call CBA_BT_CalendarShow(Me.Top, Me.Left, Me.txtProductApprovalDate, CBA_D3DMY)
    If varCal.bCalValReturned = False Then Exit Sub
    Call p_UpdIP(Me.txtProductApprovalDate, True)
    Call p_TestApprovals
End Sub

Private Sub txtGBDMApprovalDate_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim bHasDate As Boolean
    ''If Me.txtGBDMApprovalDate.BackColor = CBA_Grey Then Exit Sub
    If AST_FillTagArrays(Me.txtGBDMApprovalDate.Tag, plFrmID, plAuth, "Lock") = "lock" Then Exit Sub
    If g_IsDate(txtGBDMApprovalDate.Value, True) = False Then
        If AST_UpdGBDMDate(g_FixDate(txtGBDMApprovalDate.Value), CBA_lPromotion_ID, True, CStr(txtID.Value), "Fatal") <> "True" Then Exit Sub
    Else
        bHasDate = True
    End If
    Call CBA_BT_CalendarShow(Me.Top, Me.Left, Me.txtGBDMApprovalDate, CBA_D3DMY)
    If varCal.bCalValReturned = False Then Exit Sub
    Call AST_UpdGBDMDate(g_FixDate(txtGBDMApprovalDate.Value), CBA_lPromotion_ID, True, CStr(txtID.Value), "All")
    If g_IsDate(Me.txtGBDMApprovalDate, True) And Me.cboStatus = 1 Then
        Me.cboStatus = 3
    ElseIf g_IsDate(Me.txtGBDMApprovalDate, True) = False And Me.cboStatus = 3 Then
        Me.cboStatus = 1
    End If
    Call p_UpdIP(Me.txtGBDMApprovalDate, True)
    Call p_TestApprovals
End Sub



