VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CBA_AST_frm_ProductRows 
   Caption         =   "Super Saver Products"
   ClientHeight    =   9900
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   26040
   OleObjectBlob   =   "CBA_AST_frm_ProductRows.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "CBA_AST_frm_ProductRows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit   ' @CBA_ASyst 190716
Private plStartLine As Long, plMaxLines As Long, plEffLines As Long, plPage As Long, psMrk_Buy As String
Private pbAsc1 As Boolean, pbAsc2 As Boolean
Private psDateFmt As String     ', plSrchCol1 As Long, plSrchVal1 As String, plSrchFmt1 As String, plSrchCol2 As Long, plSrchVal2 As String, plSrchFmt2 As String
Private plAryIdx As Long, plActIdx As Long
Private pbUpdIP As Boolean, plAuth As Long, plPageIdx As Long, psFldNames() As String, ps00Ctls() As String, pl00Ctls As Long

Private ptProdRows(0 To 19) As CBA_AST_clsProdRows, psSQL As String, pbPassOk As Boolean, pbPromoChg As Boolean, sTABLE_FLDS_F As String, plNoCols As Long
Private plCGDesc As Long, plACatDesc As Long, plActMS As Long, plFreq As Long

Private Const plFrmID = 4, TABLE_PREF = "PD_", TABLE_NAME = "L2_Products", FORMTAG = "ProdRows"
Private Const plMaxSrch As Long = 10

Private Const plFORMLINES As Long = 19, plStart As Long = 1, plClr1 As Long = 12648384, plClr2 As Long = 12648447
Private Const FORM_FLDS = "ProductCode,FutureProdCode,ProductDesc,BD,GBD,Status,SubBasket,SuperSaverType,ACatCGSCG,CGSCG,TableMerch,EndCapMerch,OnSaleDate,WeeksOfSale,EndCatDate," & _
                          "UCostExpMSupp,SuppCostActMSupp,CurRetFreq,NewRetOOH,DiscStandee,SalesMultMedRegs,UPSPWPress,PriorSalesTV,CalcRadio,ProductApprovalDate,GBDMApprovalDate,CGDesc,ACatDesc,Upd,ID,"

Private Const TABLE_FLDS_B = "PD_Product_Code,PD_Future_Prod_Code,PD_Product_Desc,PD_BD,PD_GBD,Status,Sub_Basket,Super_Saver_Type,PD_ACatCGSCG,PD_CGSCG,PD_Table_Merch,PD_EndCap_Merch," & _
                             "PD_On_Sale_Date,PD_Weeks_Of_Sale,PD_End_Date,PD_Unit_Cost,PD_Supplier_Cost_Support,PD_Curr_Retail_Price,PD_Retail_Price,PD_Retail_Discount,PD_Sales_Multiplier,PD_UPSPW," & _
                             "PD_Prior_Sales,PD_Calculated_Sales,PD_Product_Approval_Date,PD_GBDM_Approval_Date,PD_CGDesc,PD_ACatDesc,Upd,PD_ID,Actual_Marketing_Support,Freq_Items"
Private Const TABLE_FLDS_M = "PD_Product_Code,PD_Future_Prod_Code,PD_Product_Desc,PD_BD,PD_GBD,Status,Sub_Basket,Super_Saver_Type,PD_ACatCGSCG,PD_CGSCG,PD_Table_Merch,PD_EndCap_Merch," & _
                             "PD_On_Sale_Date,PD_Weeks_Of_Sale,PD_Cover_Date,Expected_Marketing_Support,Actual_Marketing_Support,Freq_Items,PD_Out_Of_Home,PD_Standee,PD_Media_Regions," & _
                             "PD_Press_Dates,PD_TV,PD_Radio,PD_Product_Approval_Date,PD_GBDM_Approval_Date,PD_CGDesc,PD_ACatDesc,Upd,PD_ID"
                       
Private Sub cboPromotionID_Change()
    Dim lIdx As Long, bVis As Boolean
    ' Change of Promo
    If g_SetupIP(FORMTAG) = False Then
        Call g_SetupIP(FORMTAG, 2, True)
        ' Fill the Product List Box if the Promotion has changed
        If Me.cboPromotionID.ListIndex > -1 Then
            CBA_lPromotion_ID = Me.cboPromotionID.Column(0, Me.cboPromotionID.ListIndex)
            plAuth = CBA_lAuthority
            pbPromoChg = True: bVis = True
        Else
            bVis = False
        End If
        ' Set the visibility of the ApplyAll fields
        For lIdx = 0 To pl00Ctls
            If (ps00Ctls(1, lIdx) = "lock") Or Not bVis Then
                Me(ps00Ctls(0, lIdx)).Visible = False
            Else
                Me(ps00Ctls(0, lIdx)).Visible = True
            End If
''            Me(ps00Ctls(0, lIdx)).Visible = True
''            Debug.Print ">" & ps00Ctls(0, lIdx) & "-" & ps00Ctls(1, lIdx);
        Next
        Call p_SetSearchParms("", "", "Init")
        ' Unsetup the form
        Call g_SetupIP(FORMTAG, 2, False)
        ' Reset
        Call cmdResetProduct_Click
    End If
End Sub

Private Sub cboPromotionID_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call mw_SetBoxHook(cboPromotionID)
End Sub
                        
Private Sub cmdExcel_Click()
    ' Will create the Excel File for the user
    Dim wbk_Tmp As Workbook, wks_Tmp As Worksheet, lAryIdx As Long, lIdx As Long, lcol As Long, rngR As Range, rngA As Range
    Dim sFmt As String, sBaseFmt As String, sHdg As String, lWidth As Long, lRow As Long '', aTabFlds() As String
    Const a_HdgRow As Long = 7
    Const a_COLS As String = "A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z,AA,AB,AC,AB,AC,AD,AE,AF,AG,AH"
   
    If CBA_BasicFunctions.isRunningSheetDisplayed = False Then CBA_BasicFunctions.CBA_Running "Create ASYST Super Saver Report"
    Application.ScreenUpdating = False
    
    On Error GoTo Err_Routine
    ' Add to the workbooks
    Set wbk_Tmp = Application.Workbooks.Add
    ' Add a Worksheet
    Set wks_Tmp = ActiveSheet
    wks_Tmp.Name = Left("Promo-" & g_KeepReplace(Me.cboPromotionID.Column(1), "AlphaN", "", "-"), 30)
    wks_Tmp.Cells.Locked = True
    ' Set up the initials for the Summary sheet
    Range(wks_Tmp.Cells(1, 1), wks_Tmp.Cells(5, AST_TF("GBDMApprovalDate"))).Interior.ColorIndex = 49
    wks_Tmp.Cells(1, 1).Select
    wks_Tmp.Pictures.Insert CBA_BSA & "VBA Development Tools\IMAGES\ALDI Logo NEW mod HighRes.png"
    Cells.Font.Name = "ALDI SUED Office"
    wks_Tmp.Cells(3, 4).Value = "Promotion " & Me.cboPromotionID.Column(1) & " as of : " & g_FixDate(Now(), CBA_D3DMYHN)
    wks_Tmp.Cells(3, 4).Font.Size = 22
    wks_Tmp.Cells(3, 4).Font.ColorIndex = 2
    ' For each field in the headings array, set up the headings
    Set rngR = Range(wks_Tmp.Cells(a_HdgRow, 1), wks_Tmp.Cells(a_HdgRow, AST_TF("GBDMApprovalDate") + 1))
    rngR.Interior.ColorIndex = 6
    rngR.BorderAround xlContinuous, xlThick, xlColorIndexNone
    rngR.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    rngR.Borders(xlInsideVertical).LineStyle = xlContinuous
    rngR.HorizontalAlignment = xlHAlignCenter
    rngR.VerticalAlignment = xlVAlignCenter
    wks_Tmp.Activate
    Rows(a_HdgRow).Select
    
    Selection.Rows.WrapText = True
    Selection.Rows.AutoFit
    
    For lcol = 0 To AST_TF("GBDMApprovalDate")
        sHdg = AST_FillTagArrays(AST_TF(lcol), 2, 0, "Hdg")
        lWidth = Val(AST_FillTagArrays(AST_TF(lcol), 2, 0, "Width"))
        wks_Tmp.Cells(a_HdgRow, lcol + 1).Value = sHdg
        Range(wks_Tmp.Cells(a_HdgRow, lcol + 1), wks_Tmp.Cells(a_HdgRow, lcol + 1)).EntireColumn.ColumnWidth = lWidth
    Next lcol
    ' Cycle through all the lines
    lRow = a_HdgRow '': lProds = 0: lPG_ID = -1: sSep = "": sPromo = ""
    For lAryIdx = plStart To plMaxLines
''        lRow = lRow + 1
        For lcol = 0 To plNoCols - 4
            ' Determine if the line is to be omitted
            If p_SetSearchParms("", "", "Return", lAryIdx) = False Then GoTo SkipID
''            If plSrchCol1 > -1 And plSrchVal1 > "" Then
''                If InStr(1, LCase(CBA_PRa(0, plSrchCol1, lAryIdx)), LCase(plSrchVal1)) = 0 Then GoTo SkipID
''                If plSrchCol2 > -1 And plSrchVal2 > "" Then
''                    If InStr(1, LCase(CBA_PRa(0, plSrchCol2, lAryIdx)), LCase(plSrchVal2)) = 0 Then GoTo SkipID
''                End If
''            End If
            ' Increment row ....
            If lcol = 0 Then lRow = lRow + 1
            ' Get formats etc
            sFmt = AST_FillTagArrays(AST_TF(lcol), 2, 0, "Format")
            sBaseFmt = AST_FillTagArrays(AST_TF(lcol), 2, 0, "BaseFormat")
            If sBaseFmt = "num" And InStr(1, AST_TF(lcol), "Price") > 0 Then sBaseFmt = "cur"
            ' Set the format of the line
            If sBaseFmt = "dte" Then
                wks_Tmp.Cells(lRow, lcol + 1).NumberFormat = "ddd dd/mm/yy"
                wks_Tmp.Cells(lRow, lcol + 1).Value = CBA_PRa(0, lcol, lAryIdx)
            ElseIf sBaseFmt = "cur" And sFmt = "0.00" Then
                wks_Tmp.Cells(lRow, lcol + 1).Value = Format(CBA_PRa(0, lcol, lAryIdx), "$" & sFmt)
            ElseIf sFmt > "" Then
                wks_Tmp.Cells(lRow, lcol + 1).Value = Format(CBA_PRa(0, lcol, lAryIdx), sFmt)
            Else
                wks_Tmp.Cells(lRow, lcol + 1).Value = CBA_PRa(0, lcol, lAryIdx)
            End If
        Next
        Set rngA = Range(wks_Tmp.Cells(lRow, 1), wks_Tmp.Cells(lRow, AST_TF("GBDMApprovalDate") + 1))
''        Rows(lRow).Select
        rngA.WrapText = True
        'rngA.AutoFit
        rngA.BorderAround xlContinuous, xlThin, xlColorIndexNone
        rngA.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        rngA.Borders(xlInsideVertical).LineStyle = xlContinuous
        rngA.HorizontalAlignment = xlHAlignCenter
        rngA.VerticalAlignment = xlVAlignCenter

''        Rows(lRow).Select
''        Selection.Rows.WrapText = True
''        Selection.Rows.AutoFit
''        Selection.BorderAround xlContinuous, xlThin, xlColorIndexNone
''        Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
''        Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
''        Selection.HorizontalAlignment = xlHAlignCenter
''        Selection.VerticalAlignment = xlVAlignCenter
SkipID:
    Next
''    ' Save the xls file if rows have been added
''    If lRow > a_HdgRow Then
''        ''wks_Tmp.Protect Password:="SuperSavers"
''        ''wks_Tmp.Protect "SuperSavers", True, True, True
''''        wbs_PLF.
''        'sSavePath = g_getConfig("PromoGBDMPath", g_getDB("ASYST")) & "Promo" & sRevDate & ".xlsx"
'''        sSavePath = g_SaveFileTo("Promo Report " & g_RevDate(Date))
'''        wbk_Tmp.SaveAs sSavePath
''    End If
'''    wbk_Tmp.Close SaveChanges:=False
    
Exit_Routine:
    On Error Resume Next
    Application.ScreenUpdating = False
    If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
    wbk_Tmp.Application.Visible = True
    Exit Sub

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("ProductRows;s-cmdExcel_Click", 3)
    CBA_Error = Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "ASYST", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Sub

Private Sub cmdPageDown_Click()
    ' Page down (show the later ID numbers in the plPageIdx=0 index)
    plStartLine = plStartLine + plFORMLINES
    plPage = plPage + 1
    Call p_Display
    Call p_Pages
End Sub

Private Sub cmdPageUp_Click()
    ' Page up (show the ealier ID numbers in the plPageIdx=0 index)
    plStartLine = plStartLine - plFORMLINES
    plPage = plPage - 1
    If plStartLine < 0 Then
        plStartLine = 0: plPage = 0
    End If
    Call p_Display
    Call p_Pages
End Sub

Private Sub p_Pages()
    ' Decide if the page buttons should be shown or not
    If p_PageIdx(plPage + 1, plPageIdx, , True) < 0 Then
        Me.cmdPageDown.Visible = False
    Else
        Me.cmdPageDown.Visible = True
        Me.cmdPageDown.Caption = "Page " & plPage + 2
    End If
    
    If plStartLine < plFORMLINES Then
        Me.cmdPageUp.Visible = False
    Else
        Me.cmdPageUp.Visible = True
        Me.cmdPageUp.Caption = "Page " & plPage
    End If

End Sub

Private Sub cmdResetProduct_Click()
    Static bHasBeenRun As Boolean
    If g_SetupIP(FORMTAG) = True Then Exit Sub
    ' To stop prior updating of upd fields, set up is on
    Call g_SetupIP(FORMTAG, 2, True)
    plAryIdx = 0: plEffLines = 0: plMaxLines = 0: plActIdx = 0
    Call p_GetLiveData(CBA_lPromotion_ID)
    If bHasBeenRun = False Or pbPromoChg = True Then
        Call p_SetSearchParms("", "", "Init")
        plEffLines = plMaxLines                              ' Set the Effective Lines variable
        Call p_Sort(0, plNoCols, True)                       ' By default sort on the Product ID
        bHasBeenRun = True
        pbAsc2 = True
        pbPromoChg = False
    Else
        pbAsc1 = Not pbAsc1                                  ' Swap the sort order as the p_Sort routine will change it back when it sees the order is the same
        Call p_Sort(0, -1)                                   ' Sort by the prior field
    End If
    plPageIdx = 0: plPage = 0                                ' Set the default page idx at 0
    plStartLine = 0
    Call p_PageIdx(0, -2)                                    ' Initialise the start array lines for each page (for PageIdx=0)
    Call cmdPageUp_Click                                     ' Pull the 1st page
    ' Reset the cmd keys
    Call p_SetCmdKeys(False)
    
    Call g_SetupIP(FORMTAG, 2, False)

End Sub

Private Sub cmdSaveProduct_Click()
    ' Save the products that have been changed - Write to the table from the ASYST form concerned
    Dim aFrmFlds() As String, aTblFlds() As String, sTag As String, bfound As Boolean
    Dim sSQLFlds As String, sSQLVals As String, vVal, sPref As String
    Dim sAuditOld As String, sAuditNew As String, sAuditSQL As String, bAudit As Boolean, lID As Long, lRow As Long, lcol As Long
    Dim CN As ADODB.Connection, RS As ADODB.Recordset, bChgdFld As Boolean, sFmt As String, sAdt As String, lPGID As Long
    'Static bNot1st As Boolean
    On Error GoTo Err_Routine
    Set CN = New ADODB.Connection
    Set RS = New ADODB.Recordset
    CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("ASYST") & ";" ';INTEGRATED SECURITY=sspi;"
    ' Split the name of the Form fields into a private lcase array
    aFrmFlds = Split(Replace(Replace(sTABLE_FLDS_F, TABLE_PREF, ""), "_", ""), ",")
    ' Split the name of the Table fields into a private array
    aTblFlds = Split(sTABLE_FLDS_F, ",")
    CBA_ErrTag = ""
    Call p_Print("SaveProds--")
    ' For each array row
    For lRow = plStart To plMaxLines
        lID = NZ(CBA_PRa(1, AST_TF("ID"), lRow), 0) ' Get the ID of the table involved
        If lPGID = 0 Then
            lPGID = Me.cboPromotionID                 ''g_DLookup("PD_PG_ID", "L2_Products", "PD_ID=" & lID, "PD_ID", g_getDB("ASYST"), 0)
        End If
        If lRow = 2 Then
            lRow = lRow
        End If
        ' Decide if an add or an update
''        If lID = 0 Then
''            sSQLFlds = "INSERT INTO " & TABLE_name & " ( "
''            sSQLVals = vbCrLf & " VALUES("
''        Else
        sSQLFlds = "UPDATE " & TABLE_NAME & " SET " & TABLE_PREF & "UpdUser = '" & CBA_User & "'," & TABLE_PREF & "LastUpd=" & g_GetSQLDate(Now(), CBA_DMYHN) & "," & TABLE_PREF & "LockFlag='N',"
        sSQLVals = vbCrLf & " WHERE " & TABLE_PREF & "ID = " & lID
        sPref = ""
''        End If
        ' Cycle through the fields on the form to update the field names and the two Sqls
        For lcol = 0 To AST_TF("CGDesc") - 1
            CBA_ErrTag = "No"
            If InStr(1, aTblFlds(lcol), TABLE_PREF) = 0 Then
                If aTblFlds(lcol) <> "Actual_Marketing_Support" And aTblFlds(lcol) <> "Freq_Items" Then
                    GoTo NoTag
                Else
                    lcol = lcol
                End If
            End If
            sTag = AST_TF(lcol)
            If (sTag) > "" Then
                sFmt = AST_FillTagArrays(sTag, plFrmID, plAuth, "FullFormat")
                sAdt = AST_FillTagArrays(sTag, plFrmID, plAuth, "Audit")
                If aTblFlds(lcol) = "Actual_Marketing_Support" Or aTblFlds(lcol) = "Freq_Items" Then
                    sFmt = "text"
                End If
                CBA_ErrTag = "Yes"
                bfound = False
                sFmt = Replace(sFmt, "d2", "")          ' remove any numbers in the date to form the base date
                sFmt = Replace(sFmt, "d3", "")
                ''If (LCase(sTag) like "*Mrkting*" Or LCase(sTag) = "endcapmerch") Then
                If (LCase(sTag) Like "*marketing*") Then
                    sTag = sTag
''                    sField = frmf.Name
                End If
                'Debug.Print stag & ",";
               If aFrmFlds(lcol) = sTag Then
                    bfound = True
                    If IIf(IsNull(CBA_PRa(0, lcol, lRow)) = True, "", CBA_PRa(0, lcol, lRow)) = "" Then
                        vVal = "NULL"
                    ElseIf InStr(1, ",dtedmyy,", "," & sFmt & ",") > 0 Then
                        vVal = g_GetSQLDate(CBA_PRa(0, lcol, lRow), CBA_DMY)
                    ElseIf InStr(1, ",dtedmyyhn,", "," & sFmt & ",") > 0 Then
                        vVal = g_GetSQLDate(CBA_PRa(0, lcol, lRow), CBA_DMYHN)
                    ElseIf Left(sFmt, 3) = "num" Then
                        vVal = g_UnFmt(CBA_PRa(0, lcol, lRow), "num")
                    ElseIf (CBA_PRa(0, lcol, lRow) = True) And (sFmt = "opt" Or sFmt = "chk") Then
                        vVal = "'Y'"
                    ElseIf (CBA_PRa(0, lcol, lRow) = False) And (sFmt = "opt" Or sFmt = "chk") Then
                        vVal = "'N'"
                    Else
                        vVal = CBA_PRa(0, lcol, lRow)
                    End If
''                ElseIf (aFrmFlds(lcol) & CBA_PRa(0, lCol, lRow) = stag) And sFmt = "optyn" Then
''                    bFound = True
''                    vVal = True
                End If
                If bfound = True Then
                    sAuditOld = CStr(Replace(NZ(CBA_PRa(1, lcol, lRow), ""), "'", "`"))
                    sAuditNew = CStr(Replace(NZ(CBA_PRa(0, lcol, lRow), ""), "'", "`"))
''                    bChgdFld = ((Right(sAuditOld, 1) = "~") And Len(sAuditOld) > 1)
                    bChgdFld = sAuditOld <> sAuditNew                                          ' Has the field changed?????
                    If bChgdFld Then
                        CBA_PRa(1, lcol, lRow) = NZ(CBA_PRa(0, lcol, lRow), "")
                    End If
                    bAudit = Len(sAuditOld) > 0 And sAdt = "true"
                    ' Process any differences...
                    If sTag = "LastUpd" Then
                        vVal = g_GetSQLDate(Now(), CBA_DMYHN)
                        bChgdFld = True
                    ElseIf sTag = "UpdUser" Or sTag = "CrtUser" Then
                        vVal = "'" & CBA_User & "'"
                        bChgdFld = True
                    ElseIf sFmt = "txt1" Then
                        vVal = "'" & Left(vVal, 1) & "'"                     ' If the value is 1 char long
                    ElseIf vVal = "NULL" Or vVal = "'Y'" Or vVal = "'N'" Then
                        ' Do nothing to vVal on a null or if already formatted
                    ElseIf sFmt = "txt" Then
                        vVal = "'" & Replace(vVal, "'", "`") & "'"
                    End If
                    GoSub GSAdd_Fld_2_SQL
                    GoTo SkipFields
                End If
SkipFields:
    ''            If bFound = False Then
    ''                CBA_ErrTag = "Set"
    ''                vVal = ""
    ''            End If
            End If
NoTag:
        Next
        ' Save the line
        If sPref > "" Then
            CBA_ErrTag = "Write"
            GoSub GSEnd_Of_SQL
            RS.Open sSQLFlds & sSQLVals, CN
        End If
        ' Reset the flag, back to not updated
        CBA_PRa(0, AST_TF("Upd"), lRow) = "N"
    Next
    
Exit_Routine:
    CBA_ErrTag = "Write"
    sSQLFlds = "": sSQLVals = ""
    ' As the Promotion has been changed, it has to be resent to the regions, so blank the field that will eanble it
    sSQLFlds = "UPDATE L1_Promotions SET PG_Region_Date = NULL WHERE PG_ID = " & lPGID & ";"
    RS.Open sSQLFlds, CN
    If cmdSaveProduct = True Then
        MsgBox "Unexplained error occurred on save - please inform the ASYST maintanance staff", vbOKOnly
    Else
        MsgBox "Save was successfull", vbOKOnly
    End If
Skip_All:
    On Error Resume Next
    Set CN = Nothing
    Set RS = Nothing
        
    ' Reset the screen
    Call cmdResetProduct_Click
    Call p_Print("Aft-Reset--")
    Exit Sub
    
GSAdd_Fld_2_SQL:    ' Fill the SQL fields as per the Tag direction
    If lcol > 0 Then         ' The first in the array is the ProductCode
        If aTblFlds(lcol) = "Actual_Marketing_Support" Then
            If IsNumeric(vVal) = False Then
                vVal = Val(Right(Trim(" " & vVal), 1))
            ElseIf NZ(vVal, "") = "" Then
                vVal = Null
            End If
        ElseIf aTblFlds(lcol) = "Freq_Items" Then
            If vVal Like "Freq*" Then
                vVal = 1
            ElseIf vVal Like "Item*" Then
                vVal = 2
            ElseIf IsNumeric(vVal) = False Then
                vVal = Null
            End If
        End If

        If lID = 0 Then                                                             ' (INIT SQL=) "sSQLFlds = "INSERT INTO)   &   sSQLVals = " VALUES("
            sSQLFlds = sSQLFlds & sPref & IIf(Left(aTblFlds(lcol), 3) <> "PD_", "PD_", "") & aTblFlds(lcol)
            sSQLVals = sSQLVals & sPref & vVal
            sPref = ","      ' After the first field has been completed, fill in the comma
        Else                                                                        ' (INIT SQL=) "sSQLFlds = "UPDATE " & TABLE_name & " SET "
            If bChgdFld Then        ' If field has changed...
                sSQLFlds = sSQLFlds & sPref & IIf(Left(aTblFlds(lcol), 3) <> "PD_", "PD_", "") & aTblFlds(lcol) & "=" & vVal
                If bAudit Then      ' If field has changed, and is an Audit field...
                    CBA_ErrTag = "Audit"
                    sAuditSQL = "INSERT INTO L1_Audit (PA_Tbl_ID, PA_TF_ID, PA_Field, PA_OldValue, PA_NewValue, PA_CrtUser) " & _
                                " VALUES (" & CBA_lPromotion_ID & "," & plFrmID & ",'" & aTblFlds(lcol) & "','" & _
                                    Replace(sAuditOld, "~", "") & "','" & sAuditNew & "','" & CBA_User & "');"
                    RS.Open sAuditSQL, CN
                End If
                sPref = ","      ' After the first field has been completed, fill in the comma
            End If
        End If
    End If
    Return

GSEnd_Of_SQL:    ' Fill the last ')' and ';' when needed
    If lID = 0 Then
        sSQLFlds = sSQLFlds & ") "
        sSQLVals = sSQLVals & ");"
    Else
        sSQLFlds = sSQLFlds ''& ", " & TABLE_PREF & "LockFlag='N'"
        sSQLVals = sSQLVals & ";"
    End If
    Return

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("frmProdRows;s-cmdSaveProduct", 3)
    CBA_Error = CBA_ErrTag & " Error -" & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    If CBA_ErrTag = "No" Then
        Resume NoTag
    ElseIf CBA_ErrTag = "Set" Then
        Resume Next
    ElseIf CBA_ErrTag = "Write" Then
''        cmdSaveProduct = True
        CBA_Error = CBA_Error & vbCrLf & sSQLFlds & sSQLVals
        Debug.Print CBA_Error
        Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
        Call g_Write_Err_Table(Err, CBA_Error, "ASYST", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
        Resume Next
    ElseIf CBA_ErrTag = "Audit" Then
        CBA_Error = CBA_Error & vbCrLf & sAuditSQL
        Debug.Print CBA_Error
        Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
        Call g_Write_Err_Table(Err, CBA_Error, "ASYST", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
        Resume Next
    Else
        Debug.Print CBA_Error
        Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
        Call g_Write_Err_Table(Err, CBA_Error, "ASYST", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
''        cmdSaveProduct = True
        sPref = ""
        GoTo Exit_Routine
        Resume Next
    End If
    
End Sub

Private Sub lblProductCode_Click()
    '  Process the sort
    Call p_InitSort("ProductCode")
End Sub
Private Sub lblProductCode_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    '  Process the search
    Call p_InitSearch(Button, "ProductCode")
End Sub

Private Sub lblFutureProdCode_Click()
   '  Process the sort
    Call p_InitSort("FutureProdCode")
End Sub
Private Sub lblFutureProdCode_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    '  Process the search
    Call p_InitSearch(Button, "FutureProdCode")
End Sub


Private Sub lblProductDesc_Click()
    '  Process the sort
    Call p_InitSort("ProductDesc")
End Sub
Private Sub lblProductDesc_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    '  Process the search
    Call p_InitSearch(Button, "ProductDesc")
End Sub

Private Sub lblBD_Click()
    Call p_InitSort("BD")
End Sub
Private Sub lblBD_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    '  Process the search
    Call p_InitSearch(Button, "BD")
End Sub

Private Sub lblGBD_Click()
    Call p_InitSort("GBD")
End Sub
Private Sub lblGBD_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    '  Process the search
    Call p_InitSearch(Button, "GBD")
End Sub

Private Sub lblStatus_Click()
    '  Process the sort
    Call p_InitSort("Status")
End Sub
Private Sub lblStatus_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    '  Process the search
    Call p_InitSearch(Button, "Status")
End Sub

Private Sub lblSubBasket_Click()
    '  Process the sort
    Call p_InitSort("SubBasket")
End Sub

Private Sub lblSubBasket_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    '  Process the search
    Call p_InitSearch(Button, "SubBasket")
End Sub
Private Sub lblSuperSaverType_Click()
    '  Process the sort
    Call p_InitSort("SuperSaverType")
End Sub
Private Sub lblSuperSaverType_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    '  Process the search
    Call p_InitSearch(Button, "SuperSaverType")
End Sub

Private Sub lblACatCGSCG_Click()
    '  Process the sort
    Call p_InitSort("ACatCGSCG")
End Sub
Private Sub lblACatCGSCG_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    '  Process the search
    Call p_InitSearch(Button, "ACatCGSCG")
End Sub

Private Sub lblCGSCG_Click()
    '  Process the sort
    Call p_InitSort("CGSCG")
End Sub
Private Sub lblCGSCG_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    '  Process the search
    Call p_InitSearch(Button, "CGSCG")
End Sub

Private Sub lblTableMerch_Click()
    '  Process the sort
    Call p_InitSort("TableMerch")
End Sub

Private Sub lblTableMerch_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    '  Process the search
    Call p_InitSearch(Button, "TableMerch")
End Sub

Private Sub lblEndCapMerch_Click()
    '  Process the sort
    Call p_InitSort("EndCapMerch")
End Sub
Private Sub lblEndCapMerch_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    '  Process the search
    Call p_InitSearch(Button, "EndCapMerch")
End Sub

Private Sub lblOnSaleDate_Click()
    '  Process the sort
    Call p_InitSort("OnSaleDate")
End Sub
Private Sub lblOnSaleDate_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    '  Process the search
    Call p_InitSearch(Button, "OnSaleDate")
End Sub

Private Sub lblWeeksOfSale_Click()
    '  Process the sort
    Call p_InitSort("WeeksOfSale")
End Sub
Private Sub lblWeeksOfSale_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    '  Process the search
    Call p_InitSearch(Button, "WeeksOfSale")
End Sub

Private Sub lblEndCatDate_Click()
    Call p_InitSort(Me("lblEndCatDate").Tag)
End Sub
Private Sub lblEndCatDate_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    '  Process the search
    Call p_InitSearch(Button, Me("lblEndCatDate").Tag, "EndCatDate")
End Sub

Private Sub lblUCostExpMSupp_Click()
    Call p_InitSort(Me("lblUCostExpMSupp").Tag)
End Sub
Private Sub lblUCostExpMSupp_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    '  Process the search
    Call p_InitSearch(Button, Me("lblUCostExpMSupp").Tag, "UCostExpMSupp")
End Sub

Private Sub lblSuppCostActMSupp_Click()
    Call p_InitSort(Me("lblSuppCostActMSupp").Tag)
End Sub
Private Sub lblSuppCostActMSupp_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    '  Process the search
    Call p_InitSearch(Button, Me("lblSuppCostActMSupp").Tag, "SuppCostActMSupp")
End Sub

Private Sub lblCurRetFreq_Click()
    Call p_InitSort(Me("lblCurRetFreq").Tag)
End Sub
Private Sub lblCurRetFreq_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    '  Process the search
    Call p_InitSearch(Button, Me("lblCurRetFreq").Tag, "CurRetFreq")
End Sub

Private Sub lblNewRetOOH_Click()
    Call p_InitSort(Me("lblNewRetOOH").Tag)
End Sub
Private Sub lblNewRetOOH_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    '  Process the search
    Call p_InitSearch(Button, Me("lblNewRetOOH").Tag, "NewRetOOH")
End Sub

Private Sub lblDiscStandee_Click()
    Call p_InitSort(Me("lblDiscStandee").Tag)
End Sub
Private Sub lblDiscStandee_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    '  Process the search
    Call p_InitSearch(Button, Me("lblDiscStandee").Tag, "DiscStandee")
End Sub

Private Sub lblSalesMultMedRegs_Click()
    Call p_InitSort(Me("lblSalesMultMedRegs").Tag)
End Sub
Private Sub lblSalesMultMedRegs_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    '  Process the search
    Call p_InitSearch(Button, Me("lblSalesMultMedRegs").Tag, "SalesMultMedRegs")
End Sub

Private Sub lblUPSPWPress_Click()
    Call p_InitSort(Me("lblUPSPWPress").Tag)
End Sub
Private Sub lblUPSPWPress_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    '  Process the search
    Call p_InitSearch(Button, Me("lblUPSPWPress").Tag, "UPSPWPress")
End Sub

Private Sub lblPriorSalesTV_Click()
    Call p_InitSort(Me("lblPriorSalesTV").Tag)
End Sub
Private Sub lblPriorSalesTV_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    '  Process the search
    Call p_InitSearch(Button, Me("lblPriorSalesTV").Tag, "PriorSalesTV")
End Sub

Private Sub lblCalcRadio_Click()
    Call p_InitSort(Me("lblCalcRadio").Tag)
End Sub
Private Sub lblCalcRadio_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    '  Process the search
    Call p_InitSearch(Button, Me("lblCalcRadio").Tag, "CalcRadio")
End Sub

Private Sub lblProductApprovalDate_Click()
    Call p_InitSort("ProductApprovalDate")
End Sub
Private Sub lblProductApprovalDate_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    '  Process the search
    Call p_InitSearch(Button, "ProductApprovalDate")
End Sub

Private Sub lblGBDMApprovalDate_Click()
    Call p_InitSort("GBDMApprovalDate")
End Sub
Private Sub lblGBDMApprovalDate_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    '  Process the search
    Call p_InitSearch(Button, "GBDMApprovalDate")
End Sub

Private Sub p_InitSearch(ByVal Button As Integer, ByVal sTagName As String, Optional sFieldName As String)
    Dim sIn As String
    If plMaxLines = 0 Then Exit Sub
    ' This routine will get the search parms or null them - also set the Effective Lines variable
    If Button = 2 Then
        sIn = InputBox("Enter the value to search for (leave blank to cancel this search)")
''        If sIn > "" Then
        Call p_SetSearchParms(sTagName, sIn, "Enter", , sFieldName)
'        If sIn > "" Then
''            If plSrchCol2 > -1 Then Call p_SetSearchParms("lbl" & AST_TF(plSrchCol2), "", False)
''            plSrchVal2 = plSrchVal1
''            plSrchCol2 = plSrchCol1
''            plSrchVal1 = sIn
            plPageIdx = 1                                   ' Set the default page idx at 1
            plAryIdx = 0: plPageIdx = 0 '': plStart = 1
''            plSrchCol1 = AST_TF(sTagName)                 ' Set the column that the search belongs to
''            Call p_SetSearchParms("lbl" & sTagName, sIn, False)         ' Set the colour and append to the caption
            '''Call p_Sort(0, AST_TF(sTagName), True)     ' Sort the array
            Call p_InitPages                                ' Get the page order for the sort
'''        Else
'''''            plSrchVal1 = ""
'''''            plSrchVal2 = ""
'''''            plSrchCol1 = -1
'''''            plSrchCol2 = -1
'''''            plPageIdx = 0                                   ' Set the default page idx at 0 (the full ID idx)
'''''            Call p_SetSearchParms("", "", True)                         ' Take out any Append to the caption
'''            plPageIdx = 1                                   ' Set the default page idx at 1
'''            plEffLines = plMaxLines
'''        End If
        plStartLine = -1
        Call cmdPageUp_Click
    End If
End Sub
Private Sub p_InitSort(ByVal sFieldName As String)
    ' This routine will set the sort parameters for the form
    Call p_Sort(0, AST_TF(sFieldName))        ' Sort the array
    If plPageIdx = 1 Then Call p_InitPages  ' Get the page order for the sort
    plStartLine = -1
    Call cmdPageUp_Click                    ' Init the page
End Sub
Private Sub p_InitPages()
    Dim lIdx As Long, lFrom As Long, lTo As Long, lStep As Long, lActIdx As Long, sMsg As String, bInstr1 As Boolean, bInstr2 As Boolean
    ' This routine will init the sort order into the page arrays
    
    ' Set the sort order
    lFrom = plStart: lTo = plMaxLines: lStep = 1
    plEffLines = 0    ' Reset effective lines
    ' Find the start line of each page
    For lIdx = lFrom To lTo Step lStep
        lActIdx = p_Sort(lIdx)
''        bInstr1 = False
''        If InStr(1, LCase(CBA_PRa(0, plSrchCol1, lActIdx)), LCase(plSrchVal1)) > 0 Then
''            bInstr1 = True
''        End If
''        bInstr2 = Not (plSrchCol2 > -1)
''        If bInstr2 = False Then
''            If InStr(1, LCase(CBA_PRa(0, plSrchCol2, lActIdx)), LCase(plSrchVal2)) > 0 Then
''                bInstr2 = True
''            Else
''                bInstr2 = False
''            End If
''        End If
''        If bInstr1 = True And bInstr2 = True Then
        If p_SetSearchParms("", "", "Return", lActIdx) = True Then
            If (plEffLines / plFORMLINES) = (plEffLines \ plFORMLINES) Then
                Call p_PageIdx(lActIdx, plPageIdx, True)
'                sMsg = sMsg & lActIdx & ";(" & CBA_PRa(0, plNoCols, lActIdx) & ");"
            End If
            plEffLines = plEffLines + 1
        End If
    Next
End Sub

Private Sub txt00Status_AfterUpdate()
    Call p_ApplyValues(Me.txt00Status, , , True)
End Sub

Private Sub txt00SubBasket_AfterUpdate()
    Call p_ApplyValues(Me.txt00SubBasket, , , True)
End Sub

Private Sub txt00SuperSaverType_AfterUpdate()
    Call p_ApplyValues(Me.txt00SuperSaverType, , , True)
End Sub

Private Sub chk00EndCapMerch_Click()
    Static bChk As Boolean, bRunb4 As Boolean, lPromo As Long
    If bRunb4 Or lPromo <> Me.cboPromotionID Then
        bChk = Not bChk
        lPromo = Me.cboPromotionID
        bRunb4 = True
    Else
        bChk = True
    End If
    Call p_ApplyValues(Me.chk00EndCapMerch, , bChk, True)
End Sub

Private Sub chk00TableMerch_Click()
    Static bChk As Boolean, bRunb4 As Boolean, lPromo As Long
    If bRunb4 Or lPromo <> Me.cboPromotionID Then
        bChk = Not bChk
        lPromo = Me.cboPromotionID
        bRunb4 = True
    Else
        bChk = True
    End If
    Call p_ApplyValues(Me.chk00TableMerch, , bChk, True)
End Sub

Private Sub txt00OnSaleDate_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Static bRecurse As Boolean
    If bRecurse = True Then Exit Sub
    bRecurse = True
    If psMrk_Buy = "B" Then
        Call CBA_BT_CalendarShow(Me.Top, Me.Left, Me.txt00OnSaleDate, CBA_D2DMY, False)
        If varCal.bCalValReturned = False Then GoTo Exit_Routine
        Call p_ApplyValues(Me.txt00OnSaleDate, , , True)
    End If
Exit_Routine:
    bRecurse = False
End Sub

''Private Sub txt00WeeksOfSale_AfterUpdate()
''    Call p_ApplyValues(Me.txt00WeeksOfSale)
''End Sub
Private Sub txt00WeeksOfSale_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Static bRecurse As Boolean
    If bRecurse = True Then Exit Sub
    bRecurse = True
    If psMrk_Buy = "B" Then
        varFldVars.lFldWidth = 80
        varFldVars.lFldHeight = 0
        varFldVars.lFrmLeft = g_PosForm(0, (varFldVars.lFldWidth * 4), 0, "Left")
        ''varFldVars.lFrmwidth = g_PosForm(0, (varFldVars.lFldWidth * 4), 0, "Left")
        varFldVars.lFrmTop = g_PosForm(0, 0, 0, "Top")
        varFldVars.sHdg = "Enter Weeks Of Sales"
        varFldVars.sSQL = "AST_WeeksOfSale"
        varFldVars.sDB = "ASYST"
        varFldVars.bAllowNullOfField = False
        varFldVars.lCols = 1
        varFldVars.sType = "ComboBox"
        CBA_frmEntryField.Show vbModal
        If CBA_bDataChg Then
            txt00WeeksOfSale.Value = varFldVars.sField1
            Call p_ApplyValues(Me.txt00WeeksOfSale, , , True)
        End If
    End If
    bRecurse = False
End Sub

Private Sub txt00EndCatDate_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Static bRecurse As Boolean
    If bRecurse = True Then Exit Sub
    bRecurse = True

    If psMrk_Buy = "M" Then
        Call CBA_BT_CalendarShow(Me.Top, Me.Left, Me.txt00EndCatDate, CBA_D2DMY, False)
        If varCal.bCalValReturned = False Then Exit Sub
        Call p_ApplyValues(Me.txt00EndCatDate, , , True)
    End If
End Sub

Private Sub txt00UCostExpMSupp_AfterUpdate()
    ' Won't appear for Marketing, so just Unit Cost?
    Call p_ApplyValues(Me.txt00UCostExpMSupp, , , True)
End Sub

Private Sub txt00SuppCostActMSupp_AfterUpdate()
    If psMrk_Buy = "B" Then Call p_ApplyValues(Me.txt00SuppCostActMSupp, , , True)
End Sub
Private Sub txt00SuppCostActMSupp_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Static bRecurse As Boolean
    If bRecurse = True Then Exit Sub
    bRecurse = True
    If psMrk_Buy = "M" Then
        varFldVars.lFldWidth = 80
        varFldVars.lFldHeight = 0
        varFldVars.lFrmLeft = g_PosForm(0, (varFldVars.lFldWidth * 4), 0, "Left")
        ''varFldVars.lFrmwidth = g_PosForm(0, (varFldVars.lFldWidth * 4), 0, "Left")
        varFldVars.lFrmTop = g_PosForm(0, 0, 0, "Top")
        varFldVars.sHdg = "Enter Actual Marketing Support"
        varFldVars.sSQL = "SELECT * FROM [L0_Tiers] ORDER BY TR_ID"
        varFldVars.sDB = "ASYST"
        varFldVars.bAllowNullOfField = False
        varFldVars.lCols = 2
        varFldVars.sType = "ComboBox"
        CBA_frmEntryField.Show vbModal
        If CBA_bDataChg Then
            txt00SuppCostActMSupp.Value = varFldVars.sField2
            Call p_ApplyValues(Me.txt00SuppCostActMSupp, , , True)
        End If
    End If
    bRecurse = False
End Sub

Private Sub txt00CurRetFreq_AfterUpdate()
    If psMrk_Buy = "B" Then Call p_ApplyValues(Me.txt00CurRetFreq, , , True)
End Sub
Private Sub txt00CurRetFreq_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Static bRecurse As Boolean
    If bRecurse = True Then Exit Sub
    bRecurse = True
    If psMrk_Buy = "M" Then
        varFldVars.lFldWidth = 80
        varFldVars.lFldHeight = 0
        varFldVars.lFrmLeft = g_PosForm(0, (varFldVars.lFldWidth * 2), 0, "Left")
        varFldVars.lFrmTop = g_PosForm(0, 0, 0, "Top")
        varFldVars.sHdg = "Enter Frequency / Items Per Basket"
        varFldVars.sSQL = "SELECT * FROM [L0_Freq_Items] ORDER BY FI_ID"
        varFldVars.sDB = "ASYST"
        varFldVars.bAllowNullOfField = False
        varFldVars.lCols = 2
        varFldVars.sType = "ComboBox"
        CBA_frmEntryField.Show vbModal
        If CBA_bDataChg Then
            txt00CurRetFreq.Value = varFldVars.sField2
            Call p_ApplyValues(Me.txt00CurRetFreq, , , True)
        End If
    End If
    bRecurse = False
End Sub

Private Sub txt00NewRetOOH_AfterUpdate()
    If psMrk_Buy = "B" Then Call p_ApplyValues(Me.txt00NewRetOOH, , , True)
End Sub
Private Sub txt00NewRetOOH_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Static bRecurse As Boolean
    If bRecurse Then Exit Sub
    bRecurse = True
    If psMrk_Buy = "M" Then
        CBA_strAldiMsg = "Out_Of_Home_Dates~"
        varCal.sDate = g_FixDate(Date)
        CBA_AST_frm_AssocFld.Show vbModal
        If CBA_strAldiMsg > "" Then txt00NewRetOOH.Value = CBA_strAldiMsg
        Call p_ApplyValues(Me.txt00NewRetOOH, , , True)
    End If
    bRecurse = False
End Sub

Private Sub txt00DiscStandee_AfterUpdate()
    If psMrk_Buy = "B" Then Call p_ApplyValues(Me.txt00DiscStandee, , , True)
End Sub
Private Sub txt00DiscStandee_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Static bRecurse As Boolean
    If bRecurse = True Then Exit Sub
    bRecurse = True
    If psMrk_Buy = "M" Then
        CBA_strAldiMsg = "Standee_Dates~"
        varCal.sDate = g_FixDate(Date)
        CBA_AST_frm_AssocFld.Show vbModal
        If CBA_strAldiMsg > "" Then
            txt00DiscStandee.Value = CBA_strAldiMsg
            Call p_ApplyValues(Me.txt00DiscStandee, , , True)
        End If
    End If
    bRecurse = False
End Sub

Private Sub txt00SalesMultMedRegs_AfterUpdate()
    If psMrk_Buy = "B" Then Call p_ApplyValues(Me.txt00SalesMultMedRegs, , , True)
End Sub
Private Sub txt00SalesMultMedRegs_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Static bRecurse As Boolean
    ' If  already being processed, hop out
    If bRecurse Then Exit Sub
    bRecurse = True
    If psMrk_Buy = "M" Then
        CBA_strAldiMsg = "Media_Regions~"
        varCal.sDate = g_FixDate(Date)
        CBA_AST_frm_AssocFld.Show vbModal
        If CBA_strAldiMsg > "" Then txt00SalesMultMedRegs.Value = CBA_strAldiMsg
        Call p_ApplyValues(txt00SalesMultMedRegs, , , True)
    End If
    bRecurse = False
End Sub

Private Sub txt00UPSPWPress_AfterUpdate()
    If psMrk_Buy = "B" Then Call p_ApplyValues(Me.txt00UPSPWPress, , , True)
End Sub
Private Sub txt00UPSPWPress_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Static bRecurse As Boolean
    ' If already being processed, hop out
    If bRecurse Then Exit Sub
    bRecurse = True
    If psMrk_Buy = "M" Then
        CBA_strAldiMsg = "Press_Dates_(WA_only)~"
        varCal.sDate = g_FixDate(Date)
        CBA_AST_frm_AssocFld.Show vbModal
        If CBA_strAldiMsg > "" Then txt00UPSPWPress.Value = CBA_strAldiMsg
        Call p_ApplyValues(txt00UPSPWPress, , , True)
    End If
    bRecurse = False
End Sub

Private Sub txt00PriorSalesTV_AfterUpdate()
    If psMrk_Buy = "B" Then Call p_ApplyValues(Me.txt00PriorSalesTV, , , True)
End Sub
Private Sub txt00PriorSalesTV_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Static bRecurse As Boolean
    ' If already being processed, hop out
    If bRecurse Then Exit Sub
    bRecurse = True
    If psMrk_Buy = "M" Then
        CBA_strAldiMsg = "TV_Dates~"
        varCal.sDate = g_FixDate(Date)
        CBA_AST_frm_AssocFld.Show vbModal
        If CBA_strAldiMsg > "" Then txt00PriorSalesTV.Value = CBA_strAldiMsg
        Call p_ApplyValues(txt00PriorSalesTV, , , True)
    End If
    bRecurse = False
End Sub

Private Sub txt00CalcRadio_AfterUpdate()
    If psMrk_Buy = "B" Then Call p_ApplyValues(Me.txt00CalcRadio, , , True)
End Sub
Private Sub txt00CalcRadio_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Static bRecurse As Boolean
    ' If already being processed, hop out
    If bRecurse Then Exit Sub
    bRecurse = True
    If psMrk_Buy = "M" Then
        CBA_strAldiMsg = "Radio_Dates~"
        varCal.sDate = g_FixDate(Date)
        CBA_AST_frm_AssocFld.Show vbModal
        If CBA_strAldiMsg > "" Then txt00CalcRadio.Value = CBA_strAldiMsg
        Call p_ApplyValues(txt00CalcRadio, , , True)
    End If
    bRecurse = False
End Sub

Private Sub txt00ProductApprovalDate_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Static bRecurse As Boolean
    Dim sProds As String
    If bRecurse = True Then Exit Sub
    If psMrk_Buy = "M" Then Exit Sub
    bRecurse = True
    sProds = p_GetLineIDs
''    If AST_UpdGBDMDate(g_FixDate(Date), CBA_lPromotion_ID, True, sProds, "All", "Product Approval Date") <> "True" Then Exit Sub
    Call CBA_BT_CalendarShow(Me.Top, Me.Left, Me.txt00ProductApprovalDate, CBA_D2DMY, False)
    If varCal.bCalValReturned = False Then GoTo Exit_Routine
    Call p_ApplyValues(Me.txt00ProductApprovalDate, , , True)
Exit_Routine:
    bRecurse = False
End Sub

''Private Sub txt00GBDMApprovalDate_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
''    Static bRecurse As Boolean
''    Dim sProds As String
''    If bRecurse = True Then Exit Sub
''    If psMrk_Buy = "M" Then Exit Sub
''    bRecurse = True
''    sProds = p_GetLineIDs
''    If AST_UpdGBDMDate(g_FixDate(Date), CBA_lPromotion_ID, True, sProds, "Fatal") <> "True" Then GoTo Exit_Routine
''    Call CBA_BT_CalendarShow(Me.Top, Me.Left, Me.txt00GBDMApprovalDate, CBA_D2DMY, False)
''    If varCal.bCalValReturned = True Then
''        Call AST_UpdGBDMDate(g_FixDate(Me.txt00GBDMApprovalDate.Value), CBA_lPromotion_ID, , sProds, "All", , False)
''        Call p_ApplyValues(Me.txt00GBDMApprovalDate, , , True)
''        ''If g_IsDate(g_FixDate(Me.txt00GBDMApprovalDate.Value)) = True Then
''    End If
''Exit_Routine:
''    bRecurse = False
''End Sub

Private Sub UserForm_Initialize()
    Dim lIdx As Long, aTag() As String, lTagPos As Long, ctl As Control, sCtlTag As String, sLock As String, sBFields As String
    On Error GoTo Err_Routine
    
    ' To stop prior updating of upd fields, set up is on
    Call g_SetupIP(FORMTAG, 1, True, True)
    Me.StartUpPosition = 0
    Me.Height = 573
    Me.Width = 1260
    Me.Top = Application.Top + (Application.Height / 2) - (Me.Height / 2)
''    Me.Left = Application.Left + (Application.Width / 2) - (Me.Width / 2)
    ' Capture the Top, Width and Left positions
    Call g_PosForm(Me.Top, Me.Width, Me.Left, , True)
''    ' Get the authority level
''    CBA_lAuthority = AST_getUserASystAuthority(CBA_SetUser, True, True)
    plAuth = CBA_lAuthority
    plNoCols = 29 ' Number of elements in array
    plCGDesc = plNoCols - 3: plACatDesc = plNoCols - 2: plActMS = plNoCols - 2: plFreq = plNoCols - 3
    ' Decide which version of the form to show depending on the authority - (0 = Buyers type) (1 = Marketing type)
    If plAuth = 3 Then
        lTagPos = 1
        sTABLE_FLDS_F = TABLE_FLDS_M
        psMrk_Buy = "M"
    Else
        lTagPos = 0
        sTABLE_FLDS_F = TABLE_FLDS_B
        psMrk_Buy = "B"
    End If
    
    psDateFmt = CBA_D2DMY: CBA_lPromotion_ID = 0
    ' Setup the number of the field in an array of the input field names
    Call AST_TF("", sTABLE_FLDS_F, "PD_")
    ' This array will hold the actual column names which may be different to the Tag Names / Table Field Names.....
    psFldNames = Split(FORM_FLDS, ",")
    ' The Buying Fields
    sBFields = "UCostExpMSupp,SuppCostActMSupp,CurRetFreq,NewRetOOH,DiscStandee,SalesMultMedRegs,UPSPWPress,PriorSalesTV,CalcRadio"
    ' Set up the valid ApplyAll fields
    ReDim ps00Ctls(0 To 1, 0 To 0): pl00Ctls = -1
    ' Go through the form and set the tag and label to the form type found i.e. Name may = EndDate@CoverDate - B will = EndDate - M will = CoverDate
    CBA_ErrTag = "Tag"
    For Each ctl In Me.Controls
        sCtlTag = ""
        sCtlTag = ctl.Tag
        If sCtlTag = "" Then
            CBA_ErrTag = "Tag"
        ElseIf InStr(1, ctl.Tag, "@") > 0 Then
            ' Capture the Tag
            aTag = Split(ctl.Tag, "@")
            ctl.Tag = Trim(aTag(lTagPos))
            If Left(ctl.Name, 3) = "lbl" Then
                ' Capture the caption
                aTag = Split(ctl.Caption, "@")
                ctl.Caption = Trim(aTag(lTagPos))
            End If
        End If
        sCtlTag = ctl.Tag
        If sCtlTag > "" Then
            If Mid(ctl.Name, 4, 2) = "00" Then
                If psMrk_Buy = "B" And InStr(1, sBFields, g_Right(ctl.Name, 5)) > 0 Then
                    sLock = "lock"
                Else
                    sLock = AST_FillTagArrays(sCtlTag, plFrmID, plAuth, "Lock")
                End If
                pl00Ctls = pl00Ctls + 1
                ReDim Preserve ps00Ctls(0 To 1, 0 To pl00Ctls)
                ps00Ctls(0, pl00Ctls) = ctl.Name
                ps00Ctls(1, pl00Ctls) = sLock
            End If
        End If
    Next
    ' Get the latest version (Test to see if it is the latest)
    Me.lblVersion.Caption = CBA_getVersionStatus(g_GetDB("ASYST"), CBA_AST_Ver, "Super Saver Tool", "AST") & "-" & plAuth
    ' Set the Public Form No var
    CBA_lFrmID = plFrmID
    ' Create the 19 line classes
    For lIdx = 0 To 19
        Set ptProdRows(lIdx) = New CBA_AST_clsProdRows
        Call ptProdRows(lIdx).FormInit(Me.Controls, lIdx, Me.Top, Me.Left, FORMTAG, psMrk_Buy)
    Next
    
    ' Fill the promo cbo
    CBA_DBtoQuery = 1
    Call g_EraseAry(CBA_ABIarr)
    psSQL = "SELECT * FROM [qry_L1_Promotions] ORDER BY Sts_Seq,PG_ID"
    pbPassOk = CBA_DB_Connect.CBA_DB_CC_NonC("RETRIEVE", "qry_L1_Promotions", g_GetDB("ASYST"), CBA_MSAccess, psSQL, 120, , , False)
    If pbPassOk = True Then
        Call AST_FillDDBox(Me.cboPromotionID, 2)
    End If
    
    Call g_SetupIP(FORMTAG, 1, False)
    Call cboPromotionID_Change
    ' Call Reset to get the data and format it
''    Call cmdResetProduct_Click
    
Exit_Routine:
    On Error Resume Next
    Exit Sub
       
Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("ProdRows;s-UserFormInit", 3)
    CBA_Error = CBA_ErrTag & " Error - Field=" & sCtlTag & " - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    If CBA_ErrTag = "Tag" Then
        Resume Next
    Else
        Debug.Print CBA_Error
        Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
        Call g_Write_Err_Table(Err, CBA_Error, "ASYST", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
        GoTo Exit_Routine
        Resume Next
    End If

    
End Sub

Private Sub p_Display()
    Dim sRow As String, bVis As Boolean, bVis1 As Boolean, lcol As Long, sFldName As String, sFldNameTmp As String
    Dim bInstr1 As Boolean, bInstr2 As Boolean
''    Dim sOnSaleDate As String, lWeeksOfSale As Long
    Dim lRow As Long, sTested As String, lInc As Long
    Static bRunOnce As Boolean
    On Error GoTo Err_Routine
    ' To stop prior updating of upd fields, set up is on
    Call g_SetupIP(FORMTAG, 3, True)
    If pbAsc1 Then
        lInc = 1
    Else
        lInc = 1
'        lInc = -1
    End If
    Application.ScreenUpdating = False
    lRow = 1
    plAryIdx = p_PageIdx(plPage, plPageIdx) '- lInc      ' Set the aryidx at the first for the page
''    If plEffLines > 0 And plMaxLines <= plEffLines And plAryIdx > 0 Then plAryIdx = plAryIdx - 1
''    plMaxLines = 30
''    plAryIdx = p_PageIdx(0, -2): plPage = 0
''    plAryIdx = p_PageIdx(plPage, -1): plPage = 0
''    plAryIdx = p_PageIdx(plPage, 1)
''    plPage = 1
''    plAryIdx = p_PageIdx(plPage, 1)
''    plPage = 2
     
    ' For each row on the screen
    For lRow = 0 To plFORMLINES - 1
        plAryIdx = plAryIdx + lInc          ' Inc the ArrayIndex
        sRow = Format(lRow + 1, "00")       ' Set the Field row index
        sTested = "No"                      ' Not been tested
'''        If plAryIdx <= plMaxLines Then bVis = True Else bVis = False
        bVis = True
        If lRow = 4 Or lRow = 5 Then
            lRow = lRow
        End If
        
        For lcol = 0 To plNoCols
            If AST_TF(lcol) = "TableMerch" Or AST_TF(lcol) = "EndCapMerch" Then
                sFldName = "chk" & sRow & psFldNames(lcol)
            Else
                sFldName = "txt" & sRow & psFldNames(lcol)
            End If
            If sFldName Like "*NewRetOOH" Or sFldName Like "*Status" Then
                sFldName = sFldName
            End If
            GoSub GSVis
        Next
    Next
    bRunOnce = True
    Application.ScreenUpdating = True
    Call g_SetupIP(FORMTAG, 3, False)
Exit_Routine:
    On Error Resume Next
    Call p_Print("afp_Display")
    Exit Sub
GSVis:
    If sTested = "No" Then
    ''    bVis1 = p_SetSearchParms("", "", "Return", p_Sort(plAryIdx))
        If (plAryIdx > plMaxLines And lInc > 0) Or (plAryIdx < 0 And lInc < 0) Then
            bVis1 = False
    ''    ElseIf plSrchCol1 > -1 And plSrchVal1 > "" Then
        ElseIf p_SetSearchParms("", "", "Return", plAryIdx) = True Then
''            plActIdx = p_Sort(plAryIdx)
    ''        If sTested = "No" Then
    ''ReDoTest:
    ''''            plActIdx = p_Sort(plAryIdx)
    ''            bInstr1 = False: bInstr2 = Not (plSrchCol2 > -1)
    ''            If InStr(1, LCase(CBA_PRa(0, plSrchCol1, plActIdx)), LCase(plSrchVal1)) > 0 Then
    ''                bInstr1 = True
    ''            End If
    ''            If bInstr2 = False Then
    ''                If InStr(1, LCase(CBA_PRa(0, plSrchCol2, plActIdx)), LCase(plSrchVal2)) > 0 Then
    ''                    bInstr2 = True
    ''                Else
    ''                    bInstr2 = False
    ''                End If
    ''            End If
    ''            If bInstr1 = True And bInstr2 = True Then
    ''                bVis1 = True
    ''            Else
    ''                If (plAryIdx < plMaxLines And lInc > 0) Or (plAryIdx > 0 And lInc < 0) Then
    ''                    plAryIdx = plAryIdx + lInc
    ''                    GoTo ReDoTest
    ''                End If
    ''''                bVis1 = False
    ''''            End If
    ''            sTested = "Yes"
    ''''''        End If
    ''''''    ElseIf bVis = False Then
    ''''''        bVis1 = False
    ''''''        sTested = "Yes"
    ''''''    Else
            If sTested = "No" Then plActIdx = p_Sort(plAryIdx): sTested = "Yes"
            bVis1 = True
        Else
            bVis1 = False
            plAryIdx = plAryIdx + lInc
            GoTo GSVis
            ''sTested = "Yes"
        End If
    End If
    
    If bVis1 = True Then
        ' The last 4 columns are for the indexes, control function and tooltip descriptions
        If lcol >= plNoCols - 1 Then
            ''Debug.Print "row=" & lRow & ";col=" & lCol & ";";
            If lRow = 0 Then
                lRow = lRow
            End If
            ' Fill the value of the ID or the array index
            Me(sFldName).Value = CBA_PRa(0, lcol, plActIdx)
            If lcol = AST_TF("Upd") Then
                sFldNameTmp = Left(sFldName, 5) & "ProductCode"
                If CBA_PRa(0, lcol, plActIdx) = "E" Then            ' If being maint by s/o else
                    Me(sFldNameTmp).BackColor = CBA_Red
                ElseIf CBA_PRa(0, lcol, plActIdx) = "Y" Then        ' If locked by this form and is being maintained
                    Me(sFldNameTmp).BackColor = CBA_Green
                Else                                                ' Elae is not being maint
                    Me(sFldNameTmp).BackColor = CBA_Grey
                End If
                Me(sFldName).Tag = CBA_PRa(1, lcol, plActIdx)       ' Put the array index in the tag
            End If
            GoTo SkipID
        End If
        ' Fill the field
        Me(sFldName).Value = p_FormatFlds(Me(sFldName).Tag, plFrmID, plAuth, CBA_PRa(0, lcol, plActIdx))
        Call p_SetLockVis(Me(sFldName), False, plFrmID, plAuth)
        ' Set the CG Descriptions in the tool tip
        If InStr(1, sFldName, "ACatCGSCG") > 0 Then
            Me(sFldName).ControlTipText = CBA_PRa(0, plACatDesc, plActIdx)
        ElseIf InStr(1, sFldName, "CGSCG") > 0 Then
            Me(sFldName).ControlTipText = CBA_PRa(0, plCGDesc, plActIdx)
        End If
        If plAryIdx / 2 = plAryIdx \ 2 Then
            If Me(sFldName).BackColor <> CBA_Grey Then Me(sFldName).BackColor = plClr1
        Else
            If Me(sFldName).BackColor <> CBA_Grey Then Me(sFldName).BackColor = plClr2
        End If
    Else
        Me(sFldName).Visible = bVis1
    End If
SkipID:
    Return

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("frmProductRows;s-p_Display", 3)
    CBA_Error = " Error - Field=" & sFldName & " - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
Skip_Err_Routine:
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    GoTo Exit_Routine
    Resume Next
End Sub

Private Function p_GetLineIDs() As String
''    Dim sRow As String, bVis As Boolean, bVis1 As Boolean, lCol As Long, sFldName As String, sFldNameTmp As String
    Dim bInstr1 As Boolean, bInstr2 As Boolean, lAryIdx As Long, lActIdx As Long, sProds As String, sSep As String
    Dim lRow As Long, sTested As String, lInc As Long
    Static bRunOnce As Boolean
    On Error GoTo Err_Routine
    
    lAryIdx = plStart - 1: sProds = "": sSep = ""
    Do While lAryIdx < plMaxLines
        lAryIdx = lAryIdx + 1
        lActIdx = p_Sort(lAryIdx)
''        bInstr1 = Not (plSrchCol1 > -1)
''        bInstr2 = Not (plSrchCol2 > -1)
''        If bInstr1 = False Then
''            If InStr(1, LCase(CBA_PRa(0, plSrchCol1, lActIdx)), LCase(plSrchVal1)) > 0 Then bInstr1 = True
''        End If
''        If bInstr2 = False Then
''            If InStr(1, LCase(CBA_PRa(0, plSrchCol2, lActIdx)), LCase(plSrchVal2)) > 0 Then bInstr2 = True
''        End If
''        If bInstr1 = True And bInstr2 = True Then
        If p_SetSearchParms("", "", "Return", lActIdx) Then
            ' Fill the value of the ID or the array index
            sProds = sProds & sSep & CBA_PRa(0, AST_TF("ID"), lActIdx)
            sSep = ","
        End If
    Loop
    p_GetLineIDs = sProds
Exit_Routine:
    Exit Function
    
Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("frmProductRows;s-p_GetLineIDs", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
Skip_Err_Routine:
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    GoTo Exit_Routine
    Resume Next
End Function

Private Function p_SetSearchParms(ByVal sTagName As String, ByVal sValue As String, Init_Enter_Return As String, Optional ByVal lAryIdx As Long = 0, Optional ByVal sFieldName As String = "") As Boolean
    ' Set the back colour when the field is searched by
    Static sLastTag(plMaxSrch) As String, sLastFld(plMaxSrch) As String, sLastCap(plMaxSrch) As String, sLastVal(plMaxSrch) As String
    Static sLastOp(plMaxSrch) As String, lLastCol(plMaxSrch) As Long, sLastFmt(plMaxSrch) As String
    Dim sCaption As String, lIdx As Long, lcol As Long, sBaseFmt As String, sPlusEqMinus As String, sShowValue As String
    On Error Resume Next
    p_SetSearchParms = True
    sPlusEqMinus = ""
    If sFieldName = "" Then sFieldName = sTagName
    ' Set the Plus Minus Operator values
    If Left(sValue, 2) = "<=" Or Left(sValue, 2) = "=<" Then
        sPlusEqMinus = "<="
        sValue = LTrim(g_Right(sValue, 2))
    ElseIf Left(sValue, 2) = ">=" Or Left(sValue, 2) = "=>" Then
        sPlusEqMinus = ">="
        sValue = LTrim(g_Right(sValue, 2))
    ElseIf Left(sValue, 1) = ">" Or Left(sValue, 1) = "<" Or Left(sValue, 1) = "=" Then
        sPlusEqMinus = Left(sValue, 1)
        sValue = LTrim(g_Right(sValue, 1))
    Else
        sPlusEqMinus = "="
    End If
    ' Reset all on a change
    If Init_Enter_Return = "Init" Then
        For lIdx = plMaxSrch To 0 Step -1
            GoSub NullVals
        Next
        GoTo Exit_Routine
    ElseIf Init_Enter_Return = "Enter" Then
        ' First go through the index and see if the field has already been entered
        For lIdx = plMaxSrch To 0 Step -1
            If sLastTag(lIdx) = sTagName Then
                If sValue > "" Then
                    If sLastCap(lIdx) = "" Then sLastCap(lIdx) = Me("lbl" & sLastFld(lIdx)).Caption
                    sLastVal(lIdx) = sValue
                    sLastOp(lIdx) = sPlusEqMinus
                    Me("lbl" & sLastFld(lIdx)).Caption = sLastCap(lIdx) & vbCrLf & IIf(sLastFmt(lIdx) <> "txt", sLastOp(lIdx), "") & sValue
                Else
                    GoSub NullVals
                End If
                GoTo Exit_Routine
            End If
        Next
        ' Next go through and add it to an empty element
        For lIdx = plMaxSrch To 0 Step -1
            If sValue > "" Then
                If sLastTag(lIdx) = "" Then
                    sLastTag(lIdx) = sTagName                                         ' Capture the Tag name
                    sLastFld(lIdx) = sFieldName                                       ' Capture the field name
                    lLastCol(lIdx) = AST_TF(sTagName)                                 ' Capture the column no
                    sBaseFmt = AST_FillTagArrays(sTagName, plFrmID, 0, "BaseFormat")  ' Capture the field format
                    If sBaseFmt = "cur" Then
                        sBaseFmt = "num"
                    ElseIf sBaseFmt = "num" Then
                    ElseIf sBaseFmt = "chk" Or sBaseFmt = "opt" Then
                        sBaseFmt = "t/f"
                    Else
                        sBaseFmt = "txt"
                    End If
                    If sBaseFmt = "t/f" Then
                        If LCase(sValue) = "true" Or LCase(sValue) = "yes" Or sValue = "1" Then
                            sValue = "true"
                        ElseIf LCase(sValue) = "false" Or LCase(sValue) = "no" Or sValue = "0" Then
                            sValue = "false"
                        End If
                        sLastVal(lIdx) = sValue
                        sPlusEqMinus = "="
                        sBaseFmt = "txt"
                    Else
                        sLastVal(lIdx) = sValue
                    End If
                    sLastFmt(lIdx) = sBaseFmt
                    sLastOp(lIdx) = sPlusEqMinus
                    Me("lbl" & sLastFld(lIdx)).BackColor = CBA_LtGreen
                    Me("lbl" & sLastFld(lIdx)).ForeColor = vbBlack
                    sLastCap(lIdx) = Me("lbl" & sLastFld(lIdx)).Caption
                    Me("lbl" & sLastFld(lIdx)).Caption = sLastCap(lIdx) & vbCrLf & IIf(sLastFmt(lIdx) <> "txt", sLastOp(lIdx), "") & sValue
                    GoTo Exit_Routine
                End If
            Else
                If sLastTag(lIdx) > "" Then
                    GoSub NullVals
                    GoTo Exit_Routine
                End If
            End If
        Next
    ElseIf Init_Enter_Return = "Return" Then
        ' Go through and test the element against the ProductRows array - p_SetSearchParms will remain at True if all elements match
        For lIdx = plMaxSrch To 0 Step -1
            If sLastTag(lIdx) > "" Then
RedoFmt:
                If sLastFmt(lIdx) = "txt" Then
                    If InStr(1, LCase(CBA_PRa(0, lLastCol(lIdx), lAryIdx)), LCase(sLastVal(lIdx))) = 0 Then
                        p_SetSearchParms = False
                        GoTo Exit_Routine
                    End If
                ElseIf sLastFmt(lIdx) = "num" And g_IsNumeric(CBA_PRa(0, lLastCol(lIdx), lAryIdx)) = False Then
                    sLastFmt(lIdx) = "txt"
                    GoTo RedoFmt
                Else
                    If sLastOp(lIdx) = "<=" Then
                        If Val(CBA_PRa(0, lLastCol(lIdx), lAryIdx)) <= Val(sLastVal(lIdx)) Then
                        Else
                            p_SetSearchParms = False
                            GoTo Exit_Routine
                        End If
                    ElseIf sLastOp(lIdx) = ">" Then
                        If Val(CBA_PRa(0, lLastCol(lIdx), lAryIdx)) > Val(sLastVal(lIdx)) Then
                            ''Debug.Print ">y" & CBA_PRa(0, lLastCol(lIdx), lAryIdx);
                        Else
                            p_SetSearchParms = False
                            ''Debug.Print ">n" & CBA_PRa(0, lLastCol(lIdx), lAryIdx);
                            GoTo Exit_Routine
                        End If
                    ElseIf sLastOp(lIdx) = "<" Then
                        If Val(CBA_PRa(0, lLastCol(lIdx), lAryIdx)) < Val(sLastVal(lIdx)) Then
                        Else
                            p_SetSearchParms = False
                            GoTo Exit_Routine
                        End If
                    ElseIf sLastOp(lIdx) = ">=" Then
                        If Val(CBA_PRa(0, lLastCol(lIdx), lAryIdx)) >= Val(sLastVal(lIdx)) Then
                        Else
                            p_SetSearchParms = False
                            GoTo Exit_Routine
                        End If
                    Else
                        If Val(CBA_PRa(0, lLastCol(lIdx), lAryIdx)) <> Val(sLastVal(lIdx)) Then
                            p_SetSearchParms = False
                            GoTo Exit_Routine
                        End If
                    End If
                End If
            End If
        Next
    End If
Exit_Routine:
    If Init_Enter_Return = "Return" Then
        Init_Enter_Return = "Return"
    End If
    Exit Function

NullVals:
    Me("lbl" & sLastFld(lIdx)).BackColor = CBA_AldiBlue
    Me("lbl" & sLastFld(lIdx)).ForeColor = CBA_White
    If sLastCap(lIdx) > "" Then Me("lbl" & sLastFld(lIdx)).Caption = sLastCap(lIdx)
    sLastTag(lIdx) = ""
    sLastFld(lIdx) = ""
    sLastOp(lIdx) = ""
    sLastCap(lIdx) = ""
    lLastCol(lIdx) = -1
    sLastFmt(lIdx) = ""
    sLastVal(lIdx) = ""
    Return
End Function

Private Sub p_SetLockVis(frmf As Control, bLockOverRide As Boolean, ByVal lFrmID As Long, ByVal lAuth As Long)
    ' Set each field to Visible and or locked as per the value in the Tag
    
    Dim sLock As String, sVis As String, bHasClr As Boolean, sClr As String
    Dim bVis As Boolean, bLocked As Boolean, lColour As Long, aTag() As String, sFmt As String, sBaseFmt As String
    Dim sToolTip As String, bIsChgDate As Boolean
    Const A_FN = 0 '', A_NV = 1 ' A_Ft = 1, A_Clr = 2, A_LK = 3, A_Vs = 4, A_Ad = 5,

    On Error GoTo Err_Routine
    
    CBA_ErrTag = ""
''        If frmf.Visible = False Then GoTo NextTag          ' If the field is invisible from the off, then ignore it
    If (frmf.Tag & "") > "" Then
        CBA_ErrTag = "Yes"
        sFmt = AST_FillTagArrays(frmf.Tag, lFrmID, lAuth, "FullFormat")
        sBaseFmt = AST_FillTagArrays(frmf.Tag, lFrmID, lAuth, "BaseFormat")
        sClr = AST_FillTagArrays(frmf.Tag, lFrmID, lAuth, "Clr")
        aTag = Split((frmf.Tag & CBA_Sn), CBA_S)
        
        bVis = False: bLocked = False: bHasClr = True: bIsChgDate = True
        sToolTip = AST_FillTagArrays(frmf.Tag, lFrmID, lAuth, "ToolTip")
        ' Set the control's tooltip
        frmf.ControlTipText = sToolTip
        If sBaseFmt = "" Or sBaseFmt = "cmd" Or Left(frmf.Name, 3) = "lbl" Or Mid(frmf.Name, 4, 2) = "00" Then
            GoTo NextTag               ' Don't process CMD keys here as they are done later by themselves
        End If
        ' Set the default colour
        lColour = CBA_White
        ' Process Lock / Visible
        sLock = AST_FillTagArrays(frmf.Tag, lFrmID, lAuth, "Lock")
        sVis = AST_FillTagArrays(frmf.Tag, lFrmID, lAuth, "Vis")
        
        If aTag(A_FN) = "OutOfHome" Then
            aTag(A_FN) = aTag(A_FN)
        End If
        ' Handle the visibility
        If sVis = "invis" Then
            bVis = False
        ElseIf sVis = "vis" Or sVis = "" Then
            bVis = True
        ElseIf sProcNameSSTagAuth(sVis, lAuth) = (lAuth & "vis") Then
            bVis = True
        ElseIf sProcNameSSTagAuth(sVis, lAuth) = (lAuth & "invis") Then
            bVis = False
        ElseIf bLockOverRide = True Then
            bVis = False
        Else
           bVis = False
        End If
        ' Set the visable as per we have worked out
        frmf.Visible = bVis
        ' Don't do anything else if now invisible
        If bVis = False Then GoTo NextTag
        ' Process locked / unlocked / clrs
        If sLock = "unlock" Then
            bLocked = False
            lColour = CBA_White
        ElseIf sLock = "lock" Then
            bLocked = True
            lColour = CBA_Grey
        ElseIf bLockOverRide = True Then
            bLocked = True
            lColour = CBA_Grey
        ElseIf sProcNameSSTagAuth(sLock, lAuth) = (lAuth & "lock") Then
            bLocked = True
            lColour = CBA_Grey
        ElseIf sProcNameSSTagAuth(sLock, lAuth) = (lAuth & "unlock") Then
            bLocked = False
        ElseIf sProcNameSSTagAuth(sLock, lAuth) = (lAuth & "alock") Then        ' Add it as an alock if a date to be changed by double click
            bLocked = True
            lColour = CBA_White
        Else
            bLocked = False
            lColour = CBA_White
        End If
        ' Set the locked sts as per worked out
        frmf.Locked = bLocked
        Select Case sBaseFmt
            Case "opt"
                ''frmf.ForeColor = lColour
            Case "chk"
                ''frmf.ForeColor = lColour
                If bLocked = True Then
                    frmf.Enabled = False
                Else
                    frmf.Enabled = True
                End If
            Case Else
                If lColour = CBA_Grey Then
                    frmf.BackColor = lColour
                Else
                    Select Case sClr
                    Case Is = "y"
                        Call AST_FillYellow(frmf, "y")
                    Case Is = "n", "d"
                    Case Is = "o"
                        frmf.BackColor = CBA_OffYellow
                    Case Is = "g"
                        frmf.BackColor = CBA_Grey
                    Case Is = "p"
                        frmf.BackColor = CBA_Pink
                    Case Else
                        frmf.BackColor = lColour
                    End Select
                End If
        End Select
    End If
NextTag:
    Exit Sub
Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-frmProdRows-p_SetLockVis", 3)
    CBA_Error = frmf.Name & "-" & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    If CBA_ErrTag = "No" Then
        Resume NextTag
    Else
        Debug.Print CBA_Error
        Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
        Call g_Write_Err_Table(Err, CBA_Error, "ASYST", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
        Resume Next
    End If
End Sub

Private Function p_FormatFlds(FldTag As String, ByVal lFrmID As Long, ByVal lAuth As Long, ByVal vVal)
    Dim sFullFmt As String
    
    Call CBA_ProcI("f-p_FormatFlds", 3)
    ' Will format the field according to the value - Don't use as an AST as is specific
    sFullFmt = AST_FillTagArrays(FldTag, lFrmID, lAuth, "FullFormat")
    ' If the value comes from the field (as in if it is bought in from the form)
''    If bValFromField Then
''        vVal = NZ(frmf.Value, "")
''        If Left(sFullFmt, 3) = "dte" Then
''            If g_IsDate(vVal, True) Then
''                vVal = g_FixDate(vVal)
''            Else
''                vVal = ""
''            End If
''        ElseIf Left(sFullFmt, 3) = "num" Then
''            If g_IsNumeric(vVal) Then
''                vVal = g_UnFmt(vVal, "num")
''            Else
''                vVal = ""
''            End If
''        End If
''    End If
    If sFullFmt = "dted3dmyy" Then sFullFmt = "dted2dmyy"
    ' Format the field according to the below formulas
    If IIf(IsNull(vVal) = True, "", vVal) = "" Then
        p_FormatFlds = ""
    ElseIf sFullFmt = "dtedmyy" Then
        p_FormatFlds = g_FixDate(vVal, CBA_DMY)
    ElseIf sFullFmt = "dtedmyyhn" Then
        p_FormatFlds = g_FixDate(vVal, CBA_DMYHN)
    ElseIf sFullFmt = "dted3dmyy" Then
        p_FormatFlds = g_FixDate(vVal, CBA_D3DMY)
    ElseIf sFullFmt = "dted3dmyyhn" Then
        p_FormatFlds = g_FixDate(vVal, CBA_D3DMYHN)
    ElseIf sFullFmt = "dted2dmyy" Then
        p_FormatFlds = g_FixDate(vVal, CBA_D2DMY)
    ElseIf sFullFmt = "dted2dmyyhn" Then
        p_FormatFlds = g_FixDate(vVal, CBA_D2DMYHN)
    ElseIf sFullFmt = "num#,0" Then
        p_FormatFlds = Format(vVal, "#,0")
    ElseIf Left(sFullFmt, 5) = "num0." Then
        p_FormatFlds = Format(vVal, g_Right(sFullFmt, 3))
    ElseIf (vVal = "N") And (sFullFmt = "opt" Or sFullFmt = "chk") Then
        p_FormatFlds = False
    ElseIf (vVal = "Y") And (sFullFmt = "opt" Or sFullFmt = "chk") Then
        p_FormatFlds = True
    Else
        p_FormatFlds = vVal
    End If

End Function

''Private Sub p_setEndDate(cobj As control)
''    ' Will set the enddate for the ASyst form
''    Dim sOnSaleDate As String, sNamePrefix As String, sOnSaleDateName As String, sWeeksOfSaleName As String, sEndDateName As String, lWeeksOfSale As Long
''    sNamePrefix = Left(cobj.Name, 5)
''    sOnSaleDateName = sNamePrefix & "OnSaleDate"
''    sWeeksOfSaleName = sNamePrefix & "WeeksOfSale"
''    sEndDateName = sNamePrefix & "EndDate"
''
''    sOnSaleDate = g_FixDate(Me(sOnSaleDateName).Value)
''    lWeeksOfSale = Val(Me(sWeeksOfSaleName).Value)
''    If Not g_IsDate(sOnSaleDate) Then sOnSaleDate = ""
''    If g_IsDate(sOnSaleDate) Then
''        Me(sEndDateName).Value = g_FixDate(CDate(sOnSaleDate) + (lWeeksOfSale * 7), psDateFmt)
''    Else
''        Me(sEndDateName).Value = ""
''    End If
''End Sub

Private Function p_Sort(lIdx As Long, Optional lSortByCol1 As Long = -1, Optional bSetAsc As Boolean = False) As Long

    Static lColSortedBy As Long, lSortByCol2 As Long, bHasRun As Boolean
    Dim i As Long, j As Long, k As Long, vTemp As Variant, sSortByFmt
    Dim vTemp_i1 As Variant, vTemp_j1 As Variant, vTemp_i2 As Variant, vTemp_j2 As Variant
    Dim sSortByFmt1 As String, sSortByFmt2 As String, sSortByFld1 As String, sSortByFld2 As String
    On Error GoTo Err_Routine
    
    ' This routine will sort the index in the specified order required
    '   On the first sort it will capture the (assumedly ID) index in case that some values in the sorted index are equal
    
    ' 1st run setup
    If Not bHasRun Then
        lColSortedBy = -1
        lSortByCol2 = lSortByCol1
    End If
    

    ' If the sort has been changed...
    If lSortByCol1 > -1 Then
        ' Sort order has been changed
        If lSortByCol1 = lColSortedBy And bSetAsc = False Then
            pbAsc1 = Not pbAsc1
        Else
            lSortByCol2 = lSortByCol1           ' Save the 1st sort parameters
            pbAsc2 = pbAsc1
            pbAsc1 = True
        End If
        ' Get the field base names
        sSortByFld1 = AST_TF(lSortByCol1)
        sSortByFld2 = AST_TF(lSortByCol2)
        'Get the fields format
        sSortByFmt1 = AST_FillTagArrays(sSortByFld1, plFrmID, 1, "BaseFormat")
        sSortByFmt2 = AST_FillTagArrays(sSortByFld2, plFrmID, 1, "BaseFormat")
        ' Reset the statics....
        lColSortedBy = lSortByCol1
        ' Sort the array into it's ascending or decending order
        For i = plStart To plMaxLines - 1
            For j = i + 1 To plMaxLines
                vTemp = CBA_PRa(0, lSortByCol1, i): GoSub GSChange: vTemp_i1 = vTemp        ' Set temporary compare variables
                vTemp = CBA_PRa(0, lSortByCol1, j): GoSub GSChange: vTemp_j1 = vTemp
                vTemp = CBA_PRa(0, lSortByCol2, i): GoSub GSChange: vTemp_i2 = vTemp
                vTemp = CBA_PRa(0, lSortByCol2, j): GoSub GSChange: vTemp_j2 = vTemp
                ' If the sort is in the wrong place then swap it over
                If (vTemp_i1 > vTemp_j1 And pbAsc1 = True) Or (vTemp_i1 < vTemp_j1 And pbAsc1 = False) Then
                    GoSub GSChgArray
                ' Else if the sort place is equal by the 1st index and differs by the 2nd index then change it
                ElseIf (vTemp_i1 = vTemp_j1 And pbAsc1 = True) Or (vTemp_i1 = vTemp_j1 And pbAsc1 = False) Then
                    If (vTemp_i2 > vTemp_j2 And pbAsc2 = True) Or (vTemp_i2 < vTemp_j2 And pbAsc2 = False) Then
                        GoSub GSChgArray
                    End If
                End If
            Next j
        Next i
        
        ' Finally put the new array index into the Upd part of the (1) array
        j = AST_TF("Upd")
        For i = plStart To plMaxLines
            CBA_PRa(1, j, i) = i
        Next
    End If
    ' Deliver back the corresponding index to what was input
    p_Sort = lIdx
    bHasRun = True ' Tell the procedure that it has been run
Exit_Routine:
    On Error Resume Next
    Exit Function
    
GSChange:
    If sSortByFmt = "dte" Then
        If g_IsDate(vTemp, True) = True Then
            vTemp = g_FixDate(vTemp, "yyyymmdd")
        Else
            vTemp = g_FixDate("01/01/2000", "yyyymmdd")
        End If
    ElseIf sSortByFmt = "num" Then
        vTemp = NZ(vTemp, 0)
    Else
        vTemp = NZ(vTemp, ".")
    End If
Return

' This is the GoSub that will change the sort over
GSChgArray:
    ' Swap the sorted field across the whole array
    For k = 0 To plNoCols
        ' Swap the now values
        vTemp = CBA_PRa(0, k, j)
        CBA_PRa(0, k, j) = CBA_PRa(0, k, i)
        CBA_PRa(0, k, i) = vTemp
        ' Swap the old values
        vTemp = CBA_PRa(1, k, j)
        CBA_PRa(1, k, j) = CBA_PRa(1, k, i)
        CBA_PRa(1, k, i) = vTemp
    Next
    Return

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("ProductRows-f-p_Sort", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "ASYST", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function

Private Function p_PageIdx(ByVal lPageStrIdx_PageNo As Long, Optional ByVal lPageIdx As Long = -1, Optional ByVal bInitIdx2 As Boolean = False, Optional ByVal bRtnPageNo As Boolean = False) As Long
    
    Static aryPage() As Long, stcbHasBeenRun As Boolean, stclMaxPages As Long, stclPageIdx As Long, stclPageNo As Long, stcbInitIdx2 As Boolean
    Dim lIdx As Long, lPage As Long
    ' This routine will deliver back the page start line for the type of index being used
    ' I.e. 'lPageIdx=0' will bring back a first index of 0, the second of plFormLines, the third of plFormLies*2 etc
    ' 'lPageIdx=1' will bring back the same as the above but the input index has a different sequence
    ' When initing 'lPageIdx=1', lPageStrIdx_PageNo will = the index the page starts at; later it will = the pageno that the idx is req for
    
    ' If a reinit...
    If lPageIdx = -2 Then stcbHasBeenRun = False
    
    ' Ensure correct page index
    If lPageIdx < 0 Then
        lPageIdx = stclPageIdx
    ElseIf lPageIdx > 1 Then
        lPageIdx = 1
    End If
        
    ' On the first pass, generate and save the standard sequence (of page number start lines) - i.e. sequence of entry or ID
    If stcbHasBeenRun = False Or lPageIdx < -1 Then
        ' stclMaxPages should include the last index that will hold a value of -1, so add them
        stclMaxPages = IIf((plMaxLines / plFORMLINES) = (plMaxLines \ plFORMLINES), (plMaxLines / plFORMLINES) + 2, (plMaxLines / plFORMLINES) + 1)
        ReDim aryPage(0 To 1, 0 To stclMaxPages)
        lIdx = 0: lPage = 0
        Do While lPage <= stclMaxPages
            aryPage(0, lPage) = lIdx                 ' Init the start line for the 'lPageIdx=0)
            If lIdx = 0 Then
                aryPage(1, lPage) = 0                ' Init the 1st start line for the 'lPageIdx=1)
            Else
                aryPage(1, lPage) = -1               ' Init the other start lines for the 'lPageIdx=1)
            End If
            lIdx = lIdx + plFORMLINES                ' Increment by plFORMLINES
            lPage = lPage + 1
            If lIdx > plMaxLines Then lIdx = -1
        Loop
        stclPageIdx = lPageIdx
        stcbHasBeenRun = True
    End If
    
    ' If another sort order is selected, generate and save the new sequence (of page number start lines), with sucessive calls
    If bInitIdx2 Then
        If Not stcbInitIdx2 = bInitIdx2 Then
            stclPageNo = -1
        End If
        stclPageNo = stclPageNo + 1
        If stclPageNo > stclMaxPages Then
            stclPageNo = stclMaxPages
            lPageStrIdx_PageNo = -1
        End If
        aryPage(1, stclPageNo) = lPageStrIdx_PageNo
        ' Null the rest of the lines
        For lIdx = stclPageNo + 1 To stclMaxPages
            aryPage(1, lIdx) = -1
        Next
        lPageStrIdx_PageNo = 0                          ' If init the set the pageno to 0
    End If
    
    ' Sort out what page's index to return
    If lPageStrIdx_PageNo > stclMaxPages Then
        lPageStrIdx_PageNo = stclMaxPages
        p_PageIdx = -1
    ElseIf lPageStrIdx_PageNo < 0 Then
        lPageStrIdx_PageNo = 0
        p_PageIdx = -1
    Else
        If bRtnPageNo Then
            p_PageIdx = IIf((aryPage(lPageIdx, lPageStrIdx_PageNo) = -1), -1, lPageIdx)
        Else
            p_PageIdx = aryPage(lPageIdx, lPageStrIdx_PageNo)
        End If
    End If
    
    ' Save the last input
    stcbInitIdx2 = bInitIdx2
    stclPageIdx = lPageIdx
    
End Function

''Private Function p_PageIdx(ByVal lPageStrIdx_PageNo As Long, Optional ByVal lPageIdx As Long = -1, Optional ByVal bInitIdx2 As Boolean = False, Optional ByVal bRtnPageNo As Boolean = False) As Long
''
''    Static aryPage() As Long, stcbHasBeenRun As Boolean, stclMaxPages As Long, stclPageIdx As Long, stclPageNo As Long, stcbInitIdx2 As Boolean
''    Dim lIdx As Long, lPage As Long
''    ' This routine will deliver back the page start line for the type of index being used
''    ' I.e. 'lPageIdx=0' will bring back a first index of 0, the second of plFormLines, the third of plFormLies*2 etc
''    ' 'lPageIdx=1' will bring back the same as the above but the input index has a different sequence
''    ' When initing 'lPageIdx=1', lPageStrIdx_PageNo will = the index the page starts at; later it will = the pageno that the idx is req for
''
''    ' If a reinit...
''    If lPageIdx = -2 Then stcbHasBeenRun = False
''
''    ' Ensure correct page index
''    If lPageIdx < 0 Then
''        lPageIdx = stclPageIdx
''    ElseIf lPageIdx > 1 Then
''        lPageIdx = 1
''    End If
''
''    ' On the first pass, generate and save the standard sequence (of page number start lines) - i.e. sequence of entry or ID
''    If stcbHasBeenRun = False Or lPageIdx < -1 Then
''        ' stclMaxPages should include the last index that will hold a value of -1, so add them
''        stclMaxPages = IIf((plMaxLines / plFORMLINES) = (plMaxLines \ plFORMLINES), (plMaxLines / plFORMLINES) + 1, (plMaxLines / plFORMLINES))
''        ReDim aryPage(0 To 1, 0 To stclMaxPages)
''        lIdx = 1: lPage = 1
''        Do While lPage <= stclMaxPages
''            aryPage(0, lPage) = lIdx                 ' Init the start line for the 'lPageIdx=0)
''            If lIdx = 1 Then
''                aryPage(1, lPage) = 1                ' Init the 1st start line for the 'lPageIdx=1)
''            Else
''                aryPage(1, lPage) = 0               ' Init the other start lines for the 'lPageIdx=1)
''            End If
''            lIdx = lIdx + plFORMLINES                ' Increment by plFORMLINES
''            lPage = lPage + 1
''            If lIdx > plMaxLines Then lIdx = 0
''        Loop
''        stclPageIdx = lPageIdx
''        stcbHasBeenRun = True
''    End If
''
''    ' If another sort order is selected, generate and save the new sequence (of page number start lines), with sucessive calls
''    If bInitIdx2 Then
''        If Not stcbInitIdx2 = bInitIdx2 Then
''            stclPageNo = 0
''        End If
''        stclPageNo = stclPageNo + 1
''        If stclPageNo > stclMaxPages Then
''            stclPageNo = stclMaxPages
''            lPageStrIdx_PageNo = 0
''        End If
''        aryPage(1, stclPageNo) = lPageStrIdx_PageNo
''        ' Null the rest of the lines
''        For lIdx = stclPageNo + 1 To stclMaxPages
''            aryPage(1, lIdx) = 0
''        Next
''        lPageStrIdx_PageNo = 1                          ' If init the set the pageno to 0
''    End If
''
''    ' Sort out what page's index to return
''    If lPageStrIdx_PageNo > stclMaxPages Then
''        lPageStrIdx_PageNo = stclMaxPages
''        p_PageIdx = 0
''    ElseIf lPageStrIdx_PageNo < 1 Then
''        lPageStrIdx_PageNo = 1
''        p_PageIdx = 0
''    Else
''        If bRtnPageNo Then
''            p_PageIdx = IIf((aryPage(lPageIdx, lPageStrIdx_PageNo) = 0), 0, lPageIdx)
''        Else
''            p_PageIdx = aryPage(lPageIdx, lPageStrIdx_PageNo)
''        End If
''    End If
''
''    ' Save the last input
''    stcbInitIdx2 = bInitIdx2
''    stclPageIdx = lPageIdx
''
''End Function
''
Private Sub p_GetLiveData(lPromoID As Long)
    ' Will fill with the current Live data
    Dim sSQL As String, RS As ADODB.Recordset, CN As ADODB.Connection, lRow As Long, lcol As Long, lUpd As Long
    On Error GoTo Err_Routine
    
    CBA_ErrTag = "SQL"
    plMaxLines = plStart
    lRow = plStart - 1
    lUpd = AST_TF("Upd")
    sSQL = "SELECT " & sTABLE_FLDS_F & " FROM qry_L2_ProductRows WHERE PD_PG_ID=" & lPromoID & " ORDER BY PD_ID ;"
    Set CN = New ADODB.Connection
    Set RS = New ADODB.Recordset
    CN.Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & g_GetDB("ASYST") & ";"
    RS.Open sSQL, CN
    Do While Not RS.EOF
        lRow = lRow + 1
        If lRow = 2 Then
            lRow = lRow
        End If
        ReDim Preserve CBA_PRa(0 To 1, 0 To plNoCols, plStart To lRow)
        ' For each column / field in the query
        For lcol = 0 To plNoCols
            CBA_ErrTag = RS.Fields(lcol).Name
            CBA_ErrTag = "Lines"
            ' Fill the '0' array with the data as it is in the tables now
            CBA_PRa(0, lcol, lRow) = p_FormatFlds(AST_TF(lcol), plFrmID, plAuth, RS.Fields(lcol))
            If lcol = lUpd Then     ' For the Upd column, fill in the '1' array with the Rowno
                CBA_PRa(1, lcol, lRow) = lRow
            ElseIf lcol = plACatDesc Then  ' If is 'Act Supp', fill in the '1' array with a 'Y' if it is filled or an 'N', if not
                If psMrk_Buy = "B" Then
                    If NZ(RS.Fields(lUpd + 2), "") > "" Then
                        CBA_PRa(1, plActMS, lRow) = "Y"
                    Else
                        CBA_PRa(1, plActMS, lRow) = "N"
                    End If
                End If
            ElseIf lcol = plCGDesc Then  ' If is 'Freq Items', fill in the '1' array with a 'Y' if it is filled or an 'N', if not
                If psMrk_Buy = "B" Then
                    If NZ(RS.Fields(lUpd + 3), "") > "" Then
                        CBA_PRa(1, plFreq, lRow) = "Y"
                    Else
                        CBA_PRa(1, plFreq, lRow) = "N"
                    End If
                End If
            Else                    ' Else, fill in the '1' array with the same data that is in '0'. This will be updated with the new values, when the data is updated by the user
                CBA_PRa(1, lcol, lRow) = CBA_PRa(0, lcol, lRow)
            End If
        Next
        RS.MoveNext
    Loop
    plMaxLines = lRow
Exit_Routine:
    On Error Resume Next
    Set RS = Nothing
    Set CN = Nothing
    Exit Sub
Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("s-frm_ProductRows/p_GetLiveData", 3)
    CBA_Error = Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    If CBA_ErrTag = "SQL" Then CBA_Error = CBA_Error & vbCrLf & sSQL
    Debug.Print CBA_Error
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "ASYST", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Sub

Private Sub p_SetCmdKeys(bVis As Boolean)
    ' Will set the cmd keys after the field has been updated in cls_ProductRows
    If g_SetupIP(FORMTAG) = True And bVis = True Then Exit Sub
    If Not Me.cmdSaveProduct.Visible = bVis Then Me.cmdSaveProduct.Visible = bVis
    If Not Me.cmdResetProduct.Visible = bVis Then Me.cmdResetProduct.Visible = bVis
    Me.cmdExcel.Visible = Not bVis
End Sub

Private Sub p_ApplyValues(ctlCtl As Control, Optional bIsDate As Boolean = False, Optional bChk As Boolean = False, Optional bUpd As Boolean = False)
    Dim lIdx As Long, SNo As String, sFmt As String, sField As String, sField1 As String, sTag As String, sPrefix As String, bOSDDate As Boolean, lcol As Long, vVar
    Static bRecused As Boolean
    If NZ(ctlCtl.Value, "") = "" Then Exit Sub
    If bRecused Then Exit Sub
    Const MKT_FLDS = "UCostExpMSupp,SuppCostActMSupp,CurRetFreq,NewRetOOH,DiscStandee,SalesMultMedRegs,UPSPWPress,PriorSalesTV,CalcRadio"
    bRecused = True
    ' Get the format...
    sFmt = AST_FillTagArrays(ctlCtl.Tag, plFrmID, plAuth, "Format")
    sField = g_Right(ctlCtl.Name, 5)
    sPrefix = Left(ctlCtl.Name, 3)
    If sField = "OnSaleDate" Or sField = "WeeksOnSale" Then bOSDDate = True
'    If bUpd Then Debug.Print "aplyx;";
    ' Apply the values from the AllField to the Rows
    For lIdx = 1 To 19
        SNo = Format(lIdx, "00")
        If Me(sPrefix & SNo & sField).Visible = True Then
''            Debug.Print sPrefix & sNo & sField & " ";
            If sPrefix = "txt" Then
                If sFmt > "" And bIsDate = False Then
                    Me(sPrefix & SNo & sField) = Format(ctlCtl, sFmt)
                Else
                    Me(sPrefix & SNo & sField) = ctlCtl
                End If
            Else
                Me(sPrefix & SNo & sField).Value = bChk
            End If
            If bUpd And Me("txt" & SNo & "Upd").Value = "N" Then
                Me("txt" & SNo & "Upd").Value = "x"
''                Debug.Print sNo & ";";
            End If
            If sField = "OnSaleDate" Or sField = "WeeksOnSale" Then
                Call AST_setEndDate1(FORMTAG, Me(sPrefix & SNo & "OnSaleDate"), Me(sPrefix & SNo & "WeeksOfSale"), Me(sPrefix & SNo & "EndCatDate"), CBA_D2DMY)
            End If
            If InStr(1, MKT_FLDS, sField) > 0 And psMrk_Buy = "M" Then
                lcol = AST_TF(sField)
                If lcol < 0 Then
                    sField1 = Me("txt" & SNo & sField).Tag
                    lcol = AST_TF(sField1)
                Else
                    sField1 = sField
                End If
                vVar = ctlCtl.Value
''                CBA_PRa(0, lCol, Val(Me("txt" & SNo & "Upd").Tag)) = "x"
                CBA_PRa(0, lcol, Val(Me("txt" & SNo & "Upd").Tag)) = vVar
            End If
        End If
    Next
    Call p_Print("p_ApplyVals")
    ''Call p_SetCmdKeys(True)    ' Not needed - it should pull an update when the Upd field / event is  triggered, when setting it to 'x'
    If sPrefix = "txt" Then
        ctlCtl = ""
    Else
        ctlCtl.Value = Null
    End If
'    pbAllUpd = True:
    bRecused = False
End Sub

Private Sub p_Print(sPlace As String)
    Dim sMsg As String, SNo As String, lNo As Long
    Const lCNo As Long = 5
    If Me("txt01ProductCode").Visible = False Then Exit Sub
    For lNo = 1 To lCNo
''        SNo = Format(lNo, "00")
''        sMsg = "Form flds-" & sPlace & "-" & Me("txt" & SNo & "ProductCode").Value & "-" & Left(Me("chk" & SNo & "TableMerch").Value, 4) & "-" & Left(Me("chk" & SNo & "EndCapMerch").Value, 4) & _
''                          "-" & Me("txt" & SNo & "Upd").Value & "-" & Me("txt" & SNo & "Upd").Tag & "-" & Me("txt" & SNo & "ID").Value
''        Call g_FileWrite(Replace(g_getDB("ASYST", True), "Error", "Test"), sMsg, , False, False)
''        'Debug.Print sMsg
''        sMsg = "Arrayflds-" & sPlace & "-" & CBA_PRa(0, 0, lNo) & "-" & Left(CBA_PRa(0, AST_TF("TableMerch"), lNo), 4) & "-" & Left(CBA_PRa(0, AST_TF("EndCapMerch"), lNo), 4) & _
''               "-" & CBA_PRa(0, AST_TF("Upd"), lNo) & "-" & CBA_PRa(0, AST_TF("Upd"), lNo) & "-" & CBA_PRa(0, AST_TF("ID"), lNo) & "-" & CBA_PRa(0, AST_TF("ID"), lNo)
''        Call g_FileWrite(Replace(g_getDB("ASYST", True), "Error", "Test"), sMsg, , False, False)
''        sMsg = "Arrayflds-" & sPlace & "-" & CBA_PRa(1, 0, lNo) & "-" & Left(CBA_PRa(1, AST_TF("TableMerch"), lNo), 4) & "-" & Left(CBA_PRa(1, AST_TF("EndCapMerch"), lNo), 4) & _
''               "-" & CBA_PRa(1, AST_TF("Upd"), lNo) & "-" & CBA_PRa(1, AST_TF("Upd"), lNo) & "-" & CBA_PRa(1, AST_TF("ID"), lNo) & "-" & CBA_PRa(1, AST_TF("ID"), lNo)
''        If lNo = lCNo Then
''            sMsg = sMsg & vbCrLf & "**************" & vbCrLf
''        End If
        ''Call g_FileWrite(Replace(g_getDB("ASYST", True), "Error", "Test"), sMsg & vbCrLf, , False, False)
        'Debug.Print sMsg
    Next
End Sub

Private Sub txt01Upd_Change()
    Call p_SetCmdKeys(True)
End Sub
Private Sub txt02Upd_Change()
    Call p_SetCmdKeys(True)
End Sub
Private Sub txt03Upd_Change()
    Call p_SetCmdKeys(True)
End Sub
Private Sub txt04Upd_Change()
    Call p_SetCmdKeys(True)
End Sub
Private Sub txt05Upd_Change()
    Call p_SetCmdKeys(True)
End Sub
Private Sub txt06Upd_Change()
    Call p_SetCmdKeys(True)
End Sub
Private Sub txt07Upd_Change()
    Call p_SetCmdKeys(True)
End Sub
Private Sub txt08Upd_Change()
    Call p_SetCmdKeys(True)
End Sub
Private Sub txt09Upd_Change()
    Call p_SetCmdKeys(True)
End Sub
Private Sub txt10Upd_Change()
    Call p_SetCmdKeys(True)
End Sub
Private Sub txt11Upd_Change()
    Call p_SetCmdKeys(True)
End Sub
Private Sub txt12Upd_Change()
    Call p_SetCmdKeys(True)
End Sub
Private Sub txt13Upd_Change()
    Call p_SetCmdKeys(True)
End Sub
Private Sub txt14Upd_Change()
    Call p_SetCmdKeys(True)
End Sub
Private Sub txt15Upd_Change()
    Call p_SetCmdKeys(True)
End Sub
Private Sub txt16Upd_Change()
    Call p_SetCmdKeys(True)
End Sub
Private Sub txt17Upd_Change()
    Call p_SetCmdKeys(True)
End Sub
Private Sub txt18Upd_Change()
    Call p_SetCmdKeys(True)
End Sub
Private Sub txt19Upd_Change()
    Call p_SetCmdKeys(True)
End Sub

Private Sub UserForm_Terminate()
    Dim lIdx As Long
    For lIdx = 1 To 19
        Set ptProdRows(lIdx) = Nothing
    Next
End Sub
