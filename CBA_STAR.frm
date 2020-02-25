VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CBA_STAR 
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11730
   OleObjectBlob   =   "CBA_STAR.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "CBA_STAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private SCN As ADODB.Connection, SMCN As ADODB.Connection, CCN As ADODB.Connection, RCN(501 To 509) As ADODB.Connection
Private SRS As ADODB.Recordset, SMRS As ADODB.Recordset, CRS As ADODB.Recordset, RRS(501 To 509) As ADODB.Recordset

Private TrialList() As Variant, op() As Variant, SMoP() As Variant, CoP() As Variant, RoP() As Variant, PlanoList() As Variant
Private AffStores(501 To 509) As Scripting.Dictionary, StoDic As Scripting.Dictionary, StoNameDic As Scripting.Dictionary
Private SelTrial As Long, SelPlano As Long
Private PotentialStoresCol As Collection
Private SkipStoreChange As Boolean, SkipPlanoChange As Boolean, SkipTrialPlanoQuery As Boolean, CurLstBoxisallo As Boolean
Private shouldExist As Boolean, skipTrialPlanChange As Boolean
Private CurPlanoKey As String, CurPriorPlanoKey As String
Private PriorPlano As Variant
Private PlanoDic As Scripting.Dictionary, PriorPlanoDic As Scripting.Dictionary
Private Trial_DateFrom As Date, Trial_DateTo As Date
Private LastPlanoName As String, lastPriorPlanoName As String
Private AllocatedProducts(0 To 299) As Scripting.Dictionary
Private isAllocatedProdsAvailiable As Boolean
Private skiptrialnamechange As Boolean
Private skipregionchange As Boolean
Private skippriorplanochange As Boolean
''''PLANOLIST
'''' 0 = Region
'''' 1 = StoreName
'''' 2 = affectedFlag
'''' 3 = Current Planogram Key
'''' 4 = Current Prior Planogram Key
'''' 5 = ??
'''' 6 = ??
'''' 7 = ??
Private Sub cbo_Priorplan_Exit(ByVal Cancel As MSForms.ReturnBoolean)
'    If cbo_Priorplan.ListCount = 0 Then
'        MsgBox "Please enter a Prior Date From"
'    End If
End Sub

Private Sub lst_Allocatedplan_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim yn As Long
Dim temp As Variant
Dim a As Long
Dim b As Long
Dim cnt As Long
Dim prod As Variant
Dim TempAProd(0 To 299) As Scripting.Dictionary
    yn = MsgBox("Would you like to delete this allocated planogram?", vbYesNo)
    If yn = 6 Then
        If UBound(PlanoList, 2) = 0 Then
            ReDim PlanoList(0 To 0, 0 To 0)
            PlanoList(0, 0) = 0
            clearStoPlanoBoxes
            SelPlano = -1
            Me.fme_Product.Visible = False
        Else
        cnt = -1
        ReDim temp(0 To 7, 0 To UBound(PlanoList, 2) - 1)
            For a = LBound(PlanoList, 2) To UBound(PlanoList, 2)
                If a <> SelPlano Then
                    cnt = cnt + 1
                    For b = LBound(PlanoList, 1) To UBound(PlanoList, 1)
                        temp(b, cnt) = PlanoList(b, a)
                    Next
                    Set TempAProd(cnt) = New Scripting.Dictionary
                    For Each prod In AllocatedProducts(a)
                        TempAProd(cnt).Add prod, prod
                    Next
                End If
            Next
            PlanoList = temp
            Erase AllocatedProducts
            'ReDim AllocatedProducts(0 To 299)
            For a = LBound(PlanoList, 2) To UBound(PlanoList, 2)
                Set AllocatedProducts(a) = New Scripting.Dictionary
                Set AllocatedProducts(a) = TempAProd(a)
            Next
        End If
        SkipTrialPlanoQuery = True
        openPlanFrame
        SkipTrialPlanoQuery = False
    Else
        MsgBox "No changes made"
    End If

End Sub

Private Sub txt_Weeks_Afterupdate()
    If Me.txt_Datefrom <> "" And Trial_DateFrom > 0 And txt_Weeks > 0 And IsNumeric(txt_Weeks) Then
        Trial_DateTo = DateAdd("D", -1, DateAdd("WW", txt_Weeks, Trial_DateFrom))
        Me.txt_Dateto = Format(DateAdd("D", -1, DateAdd("WW", txt_Weeks, Trial_DateFrom)), "dddd, mmmm dd, yyyy")
    Else
        If Trial_DateFrom > 0 And Trial_DateTo > 0 Then
            Me.txt_Weeks = DateDiff("WW", Trial_DateFrom, Trial_DateTo)
        Else
            Me.txt_Weeks = ""
        End If
        MsgBox "Please enter numeric value greater than zero"
    End If
End Sub

Sub UserForm_Initialize()
Dim lTop As Long, lLeft As Long, lRow As Long, lcol As Long
Dim a As Long, div As Long
Dim strSQL As String
Dim RS As ADODB.Recordset

    
    If CBA_BasicFunctions.isRunningSheetDisplayed = False Then CBA_BasicFunctions.CBA_Running "Connecting to STAR..."
    DoEvents
    Me.StartUpPosition = 0
    Me.Top = Application.Top + (Application.Height / 2) - (Me.Height / 2)
    Me.Left = Application.Left + (Application.Width / 2) - (Me.Width / 2)
         
    Set SCN = New ADODB.Connection
    With SCN
        .ConnectionTimeout = 50
        .CommandTimeout = 50
        .Open "Provider=SQLNCLI10;DATA SOURCE=" & CBA_BasicFunctions.TranslateServerName("599DBL12", Date) & ";;INTEGRATED SECURITY=sspi;"
        Set RS = New ADODB.Recordset
        RS.Open "SELECT state_desc  FROM sys.databases where name = 'Tools'", SCN
        If LCase(RS(0)) <> "online" Then
            MsgBox "Due to dependant database states, STAR is currently unavailable. This issue should resolve itslef in 5 mins."
            GoTo Unavailable
        End If
    End With
    Set SMCN = New ADODB.Connection
    With SMCN
        .ConnectionTimeout = 50
        .CommandTimeout = 50
        .Open "Provider=SQLNCLI10;DATA SOURCE=" & CBA_BasicFunctions.TranslateServerName("599DBL11", Date) & ";;INTEGRATED SECURITY=sspi;"
        Set RS = New ADODB.Recordset
        RS.Open "SELECT state_desc  FROM sys.databases where name = 'spaceman'", SMCN
        If LCase(RS(0)) <> "online" Then
            MsgBox "Due to dependant database states, STAR is currently unavailable. This issue should resolve itslef in 5 mins."
            GoTo Unavailable
        End If
    End With
    Set CCN = New ADODB.Connection
    With CCN
        .ConnectionTimeout = 50
        .CommandTimeout = 50
        .Open "Provider=SQLNCLI10;DATA SOURCE=599DBL01;;INTEGRATED SECURITY=sspi;"
        Set RS = New ADODB.Recordset
        RS.Open "SELECT state_desc  FROM sys.databases where name = 'cbis599p'", CCN
        If LCase(RS(0)) <> "online" Then
            MsgBox "Due to dependant database states, STAR is currently unavailable. This issue should resolve itslef in 5 mins."
            GoTo Unavailable
        End If
    End With
    For div = 501 To 509
        If div = 508 Then div = 509
        Set RCN(div) = New ADODB.Connection
        With RCN(div)
            .ConnectionTimeout = 50
            .CommandTimeout = 50
            .Open "Provider=SQLNCLI10;DATA SOURCE=0" & div & "Z0IDBSRVL02;;INTEGRATED SECURITY=sspi;"
        End With
        Set RS = New ADODB.Recordset
        RS.Open "SELECT state_desc  FROM sys.databases where name = 'purchase'", RCN(div)
        If LCase(RS(0)) <> "online" Then
            MsgBox "Due to dependant database states, STAR is currently unavailable. This issue should resolve itslef in 5 mins."
            GoTo Unavailable
        End If
    Next
    getSTARData "Trial"
    Me.txt_Trialname.Visible = False
    cmd_AddTrial.Visible = True
    SelPlano = -1
    SelTrial = -1

    'Me.cmd_Newtrial.visible = False
    'Pull planogram product store mapping
    Set RS = New ADODB.Recordset
    strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
    strSQL = strSQL & "select pl.KEY_ID, pl.PLANOGRAM, p.PRODUCT_ID, p.name, convert(nvarchar(6),isnull(st.STORENUMBER,0)) as storenumber, pl.PLANOFOLDER03" & Chr(10)
    strSQL = strSQL & "into #DATA from spaceman.dbo.SC_PLANO_KEY pl" & Chr(10)
    strSQL = strSQL & "left join spaceman.dbo.sc_product_list pr on pr.KEY_ID = pl.KEY_ID" & Chr(10)
    strSQL = strSQL & "left join Spaceman.dbo.sc_product p on p.PRODUCT_ID = pr.PRODUCT_ID" & Chr(10)
    strSQL = strSQL & "left join Spaceman.dbo.SC_NSW s on pl.KEY_ID = s.KEY_ID" & Chr(10)
    strSQL = strSQL & "left join myspaceman.dbo.SYS_MYS_PLANOCONTAINER pc on pc.PLANOGRAMID = s.PLANOGRAMID" & Chr(10)
    strSQL = strSQL & "left join myspaceman.dbo.USR_CR_CONTENTRECVRPLANOGRAMS crp on crp.PLANOCONTAINERID = pc.CONTAINERID" & Chr(10)
    strSQL = strSQL & "left join myspaceman.dbo.USR_CR_CONTENTRECVR cr on cr.ID = crp.CONTENTRECEIVERID" & Chr(10)
    strSQL = strSQL & "left join myspaceman.dbo.USR_CR_STORELIST st on st.ID = cr.ID" & Chr(10)
    strSQL = strSQL & "where storenumber <> 0" & Chr(10)
    strSQL = strSQL & "" & Chr(10)
    strSQL = strSQL & "select  KEY_ID, PLANOGRAM,PRODUCT_ID as PCode, name as PDesc" & Chr(10)
    strSQL = strSQL & ", SUBSTRING(storenumber,1,3) as region, SUBSTRING(storenumber,4,3) as storeno ,PLANOFOLDER03" & Chr(10)
    strSQL = strSQL & "into #PSP from #DATA where PLANOFOLDER03 = 'Live'" & Chr(10)
    strSQL = strSQL & "" & Chr(10)
    strSQL = strSQL & "drop table #DATA" & Chr(10)
    RS.Open strSQL, SMCN
    Set RS = Nothing
    'Prepare Planogram Frame
    Me.cbo_Affected.AddItem "Yes"
    Me.cbo_Affected.AddItem "No"
    Me.fme_Plan.Visible = False
    Me.cbo_Region = ""
    For a = 501 To 509
        If a = 508 Then a = 509
        Me.cbo_Region.AddItem CBA_BasicFunctions.CBA_DivtoReg(a)
        Set AffStores(a) = New Scripting.Dictionary
    Next
    getSTARData "AffectedStores"
    
    'Prepare Product Frame
    Me.fme_Product.Visible = False
    
    
    If IsEmpty(TrialList) Then
        
    Else
        For a = LBound(TrialList, 2) To UBound(TrialList, 2)
            Me.cbo_Trialname.AddItem TrialList(1, a) & "-" & TrialList(0, a)
        Next
    End If
    lbl_PleaseClose.Visible = False
    If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
    shouldExist = True
Exit Sub

Unavailable:
    AffectAllObjects True, "Please close this userform"
    shouldExist = False
    If CBA_BasicFunctions.isRunningSheetDisplayed = True Then CBA_BasicFunctions.CBA_Close_Running
    
    
End Sub
Private Sub cmd_Addplan_Click()
Dim temp As Variant
Dim a As Long, b As Long
Dim alreadythere As Boolean
Dim strStore As String
Dim yn As Long
Dim c As Long
    If Me.cbo_Region <> "" And Me.cbo_Store <> "" And Me.cbo_Trialplan <> "" Then 'And Me.cbo_Priorplan <> "" Then
        For a = 0 To Me.lst_Allocatedplan.ListCount - 1
                If InStr(1, Me.cbo_Store, "- ALCOHOL") > 0 Then
                    strStore = Trim(Mid(Me.cbo_Store, 1, InStr(1, Me.cbo_Store, "- ALCOHOL") - 1))
                Else
                    strStore = Me.cbo_Store
                End If
                If lst_Allocatedplan.List(a, 0) = Me.cbo_Region And lst_Allocatedplan.List(a, 1) = strStore And _
                Trim(lst_Allocatedplan.List(a, 3)) = Trim(Me.cbo_Trialplan) Then 'And lst_Allocatedplan.List(a, 4) = Me.cbo_Priorplan Then
                    alreadythere = True
                    Exit For
            End If
        Next
        If alreadythere = False Then
            If PlanoList(0, 0) <> 0 Then
                ReDim temp(0 To 7, 0 To UBound(PlanoList, 2) + 1)
                For a = LBound(PlanoList, 2) To UBound(PlanoList, 2)
                    For b = LBound(PlanoList, 1) To UBound(PlanoList, 1)
                        temp(b, a) = PlanoList(b, a)
                    Next
                Next
            Else
                ReDim temp(0 To 7, 0 To 0)
            End If
            temp(0, UBound(temp, 2)) = CBA_BasicFunctions.CBA_DivtoReg(cbo_Region)
            temp(1, UBound(temp, 2)) = StoNameDic(CStr(Trim(Mid(cbo_Store, InStr(1, cbo_Store, "-") + 1, 999))))
            If AffStores(CBA_BasicFunctions.CBA_DivtoReg(cbo_Region))(StoNameDic(CStr(cbo_Store))) = "A" Then temp(2, UBound(temp, 2)) = 1 Else temp(2, UBound(temp, 2)) = 0
            temp(3, UBound(temp, 2)) = CurPlanoKey
            'getSPACEMANData "PlanogramNumber", CBA_BasicFunctions.CBA_DivtoReg(cbo_Region), "PriorTrialPlan"
            temp(4, UBound(temp, 2)) = CurPriorPlanoKey
            temp(7, UBound(temp, 2)) = 0
            PlanoList = temp
            SkipTrialPlanoQuery = True ': skippriorplanochange = True
            openPlanFrame
            SkipTrialPlanoQuery = False ': skippriorplanochange = False
            Me.lst_Allocatedplan.Selected(Me.lst_Allocatedplan.ListCount - 1) = True
            

            
            
        Else
            
            If Me.cbo_Priorplan <> "" And Me.cbo_Priorplan <> NZ(Me.lst_Allocatedplan.List(SelPlano, 4), "") Then
            'NEED TO UNDERSTAND IF PRIORPLAN ON ADD IS DIFFERENT TO PRIOR PLAN IN ALLOCATED LISTBOX. IF IT IS THEN USER MAY WANT TO CHANGE
                yn = MsgBox("Would you like to change the prior planogram assigment?", vbYesNo)
                If yn = 6 Then
                    Me.lst_Allocatedplan.List(SelPlano, 4) = Me.cbo_Priorplan
                    PlanoList(7, SelPlano) = 2
                    PlanoList(4, SelPlano) = CurPriorPlanoKey
                Else
                    MsgBox "Store/Planogram combination is already allocated"
                End If
            
            Else
                MsgBox "Store/Planogram combination is already allocated"
            End If
        End If
    ElseIf Me.cbo_Region <> "" And Me.cbo_Store = "" And Me.cbo_Trialplan <> "" Then 'And Me.cbo_Priorplan <> "" Then
        yn = MsgBox("Do you wish to add all stores that hold this planogram?", vbYesNo)
        If yn = 6 Then
            getSPACEMANData "ALLPLANONOSTORE", CBA_BasicFunctions.CBA_DivtoReg(cbo_Region)
            For c = LBound(SMoP, 2) To UBound(SMoP, 2)
                For a = 0 To Me.lst_Allocatedplan.ListCount - 1
                    If InStr(1, Me.cbo_Store, "- ALCOHOL") > 0 Then
                        strStore = Trim(Mid(Me.cbo_Store, 1, InStr(1, Me.cbo_Store, "- ALCOHOL") - 1))
                    Else
                        strStore = Me.cbo_Store
                    End If
                    If lst_Allocatedplan.List(a, 0) = Me.cbo_Region And lst_Allocatedplan.List(a, 1) = CLng(SMoP(0, c)) And _
                    Trim(lst_Allocatedplan.List(a, 3)) = Trim(Me.cbo_Trialplan) Then 'And lst_Allocatedplan.List(a, 4) = Me.cbo_Priorplan Then
                        alreadythere = True
                        Exit For
                    End If
                Next
                If alreadythere = False Then
                    If PlanoList(0, 0) <> 0 Then
                        ReDim temp(0 To 7, 0 To UBound(PlanoList, 2) + 1)
                        For a = LBound(PlanoList, 2) To UBound(PlanoList, 2)
                            For b = LBound(PlanoList, 1) To UBound(PlanoList, 1)
                                temp(b, a) = PlanoList(b, a)
                            Next
                        Next
                    Else
                        ReDim temp(0 To 7, 0 To 0)
                    End If
                    temp(0, UBound(temp, 2)) = CBA_BasicFunctions.CBA_DivtoReg(cbo_Region)
                    temp(1, UBound(temp, 2)) = CLng(SMoP(0, c))
                    If AffStores(CBA_BasicFunctions.CBA_DivtoReg(cbo_Region))(CLng(SMoP(0, c))) = "A" Then temp(2, UBound(temp, 2)) = 1 Else temp(2, UBound(temp, 2)) = 0
                    temp(3, UBound(temp, 2)) = CurPlanoKey
                    'getSPACEMANData "PlanogramNumber", CBA_BasicFunctions.CBA_DivtoReg(cbo_Region), "PriorTrialPlan"
                    temp(4, UBound(temp, 2)) = CurPriorPlanoKey
                    temp(7, UBound(temp, 2)) = 0
                    PlanoList = temp
                    SkipTrialPlanoQuery = True ': skippriorplanochange = True
                    openPlanFrame
                    SkipTrialPlanoQuery = False ': skippriorplanochange = False
                    Me.lst_Allocatedplan.Selected(Me.lst_Allocatedplan.ListCount - 1) = True
                Else
                
                
                End If
            Next
                
            
        Else
            
            
        End If
    
    Else
        MsgBox "Region, Store, Trial Planogram, and Prior Planogram must all be selected"
    End If
    
    
End Sub

Private Sub cmd_AddTrial_Click()
Dim temp As Variant
Dim a As Long
Dim b As Long
    If Me.cbo_Trialname <> "" And Me.txt_Datefrom <> "" And Me.txt_Dateto <> "" Then
        
        Me.cmd_AddTrial.Visible = False
        'Me.cmd_Newtrial.visible = True
        
        'Me.cbo_Trialname.visible = True
        
        ReDim temp(0 To UBound(TrialList, 1), 0 To UBound(TrialList, 2) + 1)
        For a = 0 To UBound(TrialList, 2)
            For b = LBound(TrialList, 1) To UBound(TrialList, 1)
                temp(b, a) = TrialList(b, a)
            Next
        Next
        temp(0, UBound(temp, 2)) = TrialList(0, UBound(TrialList, 2)) + 1
        temp(1, UBound(temp, 2)) = Me.cbo_Trialname
        temp(2, UBound(temp, 2)) = Trial_DateFrom
        temp(3, UBound(temp, 2)) = Trial_DateTo
        temp(4, UBound(temp, 2)) = 0
        TrialList = temp
        Me.cbo_Trialname.AddItem Me.cbo_Trialname & "-" & TrialList(0, UBound(TrialList, 2))
        Me.cbo_Trialname = Me.cbo_Trialname & "-" & TrialList(0, UBound(TrialList, 2))
        'Me.txt_Trialname = ""
        
        openPlanFrame
    Else
        MsgBox "Name, Datefrom and DateTo are all required to setup a new Merch Trial"
        Me.cmd_AddTrial.Visible = False
        'Me.cmd_Newtrial.visible = True
        Me.txt_Trialname.Visible = False
        Me.cbo_Trialname.Visible = True
    End If

End Sub
Private Sub lst_Productallo_Click()
    CurLstBoxisallo = True
End Sub

Private Sub lst_Productallo_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CurLstBoxisallo = True
End Sub

Private Sub lst_Productsel_Click()
    CurLstBoxisallo = False
End Sub

Private Sub lst_Productsel_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    CurLstBoxisallo = False
End Sub

Private Sub txt_Datefrom_AfterUpdate()
    If Me.txt_Datefrom <> Format(Trial_DateFrom, "dddd, mmmm dd, yyyy") Then
        If isDate(txt_Datefrom) = False Then
            MsgBox "Please enter present or future date with format (DD/MM/YYYY)"
            Me.txt_Datefrom = Format(Trial_DateFrom, "dddd, mmmm dd, yyyy")
        ElseIf txt_Datefrom <> "" Then
            Trial_DateFrom = CDate(txt_Datefrom)
            If txt_Weeks > 0 And txt_Weeks <> "" Then
                Trial_DateTo = CDate(DateAdd("WW", txt_Weeks, DateAdd("D", -1, txt_Datefrom)))
                Me.txt_Dateto = Format(Trial_DateTo, "dddd, mmmm dd, yyyy")
            End If
            Me.txt_Datefrom = Format(Trial_DateFrom, "dddd, mmmm dd, yyyy")
        End If
    End If
End Sub

Private Sub txt_Dateto_AfterUpdate()
    If txt_Dateto <> Format(Trial_DateTo, "dddd, mmmm dd, yyyy") Then
    If isDate(txt_Dateto) = False Then
        MsgBox "Please enter date 'greater than datefrom', with format (DD/MM/YYYY)"
        txt_Dateto = Format(Trial_DateTo, "dddd, mmmm dd, yyyy")
    ElseIf Trial_DateFrom <> 0 And txt_Weeks > 0 And txt_Weeks <> "" And txt_Dateto <> "" Then
        Trial_DateTo = CDate(txt_Dateto)
        txt_Weeks = DateDiff("WW", Trial_DateFrom, Trial_DateTo)
        Me.txt_Dateto = Format(Trial_DateTo, "dddd, mmmm dd, yyyy")
    ElseIf txt_Dateto <> "" Then
        Trial_DateTo = CDate(txt_Dateto)
        Me.txt_Dateto = Format(Trial_DateTo, "dddd, mmmm dd, yyyy")
    End If
    If txt_Dateto <> "" And txt_Datefrom <> "" Then
        If Trial_DateTo < Trial_DateFrom Then
            'Trial_DateTo = 0
            MsgBox "Please enter date 'greater than datefrom', with format (DD/MM/YYYY)"
            txt_Dateto = ""
        End If
    End If
    End If
End Sub
Private Sub cmd_Addprod_Click()
Dim a As Long
Dim newcol As Collection
Dim PCde As String
    'Set newcol = New Collection
    
    
    
    For a = 0 To Me.lst_Productsel.ListCount - 1
        If Me.lst_Productsel.Selected(a) = True Then
            Me.lst_Productallo.AddItem Me.lst_Productsel.List(a)
            PCde = Mid(Me.lst_Productsel.List(a), 1, InStr(1, Me.lst_Productsel.List(a), "-") - 1)
            If AllocatedProducts(SelPlano).Exists(PCde) = False Then
                AllocatedProducts(SelPlano).Add PCde, PCde
            End If
        End If
    Next
    For a = 0 To Me.lst_Productsel.ListCount - 1
        If Me.lst_Productsel.ListCount > a Then
            If Me.lst_Productsel.Selected(a) = True Then
                Me.lst_Productsel.RemoveItem (a)
                a = a - 1
            End If
        Else
            Exit For
        End If
    Next
End Sub

Private Sub cmd_Confirm_Click()
Dim yn As Long
    yn = MsgBox("Save new Merch Trial to STAR?", vbYesNo)
    If yn = 6 Then
        InterfaceDataToSTAR
    Else
        MsgBox "Trial Not Saved"
    End If
End Sub

Private Sub cmd_Removeprod_Click()
Dim a As Long
Dim newcol As Collection
Dim PCde As String
    'Set newcol = New Collection
    For a = 0 To Me.lst_Productallo.ListCount - 1
        If Me.lst_Productallo.Selected(a) = True Then
            Me.lst_Productsel.AddItem Me.lst_Productallo.List(a)
            PCde = Mid(Me.lst_Productallo.List(a), 1, InStr(1, Me.lst_Productallo.List(a), "-") - 1)
            If AllocatedProducts(SelPlano).Exists(PCde) = True Then
                AllocatedProducts(SelPlano).Remove PCde
            End If
        End If
    Next
'    For Each prod In AllocatedProducts(SelPlano)
'        Debug.Print prod
'    Next
    
    
    For a = 0 To Me.lst_Productallo.ListCount - 1
        If Me.lst_Productallo.ListCount > a Then
            If Me.lst_Productallo.Selected(a) = True Then
                Me.lst_Productallo.RemoveItem (a)
                a = a - 1
            End If
        Else
            Exit For
        End If
    Next
End Sub
Private Sub cmd_SelAll_Click()
Dim a As Long
    If CurLstBoxisallo = True Then
        For a = 0 To Me.lst_Productallo.ListCount - 1
            Me.lst_Productallo.Selected(a) = True
        Next
    Else
        For a = 0 To Me.lst_Productsel.ListCount - 1
            Me.lst_Productsel.Selected(a) = True
        Next
    End If
End Sub

Private Sub cbo_Priorplan_Change()
Dim a As Long
Dim bfound As Boolean
If skippriorplanochange = False Then
    If cbo_Priorplan.ListCount = 0 And CurPriorPlanoKey <> "" And SelPlano > -1 Then
        'POPLULATE THE PRIOR DATE FROM FILER FROM THE PLANOLIST
        Me.txt_Priordatefrom = CDate(Right(Me.lst_Allocatedplan.List(SelPlano, 4), 8))
        populatePriorPlanoBox
    End If
    For a = 0 To cbo_Priorplan.ListCount - 1
        If cbo_Priorplan.List(a) = cbo_Priorplan Then
            bfound = True
            Exit For
        End If
    Next
    If bfound = False Then cbo_Priorplan = "": CurPriorPlanoKey = "": Exit Sub
    CurPriorPlanoKey = PriorPlanoDic(CStr(cbo_Priorplan))
End If
End Sub
Private Sub txt_Priordatefrom_AfterUpdate()
    If isDate(txt_Priordatefrom) Or txt_Priordatefrom = "" Then
        Call populatePriorPlanoBox
    Else
        If isDate(txt_Priordatefrom) = False And txt_Priordatefrom <> "" Then
            MsgBox "Please enter date with format (DD/MM/YYYY)"
            txt_Priordatefrom = ""
        End If
    End If
End Sub
Private Sub txt_Priordateto_AfterUpdate()
    If isDate(txt_Priordateto) Or txt_Priordateto = "" Then
        Call populatePriorPlanoBox
    Else
        If isDate(txt_Priordateto) = False And txt_Priordateto <> "" Then
            MsgBox "Please enter date with format (DD/MM/YYYY)"
            txt_Priordateto = ""
        End If
    End If
End Sub
Function populatePriorPlanoBox()
Dim a As Long
    Me.cbo_Priorplan = ""
    Me.cbo_Priorplan.Clear
    If txt_Priordatefrom > 0 Or txt_Priordateto > 0 Then
        getSPACEMANData "ArchivePlano"
        Set PriorPlanoDic = New Scripting.Dictionary
        If PriorPlano(0, 0) <> 0 Then
            For a = LBound(PriorPlano, 2) To UBound(PriorPlano, 2)
                PriorPlanoDic.Add PriorPlano(1, a) & "-" & Format(PriorPlano(2, a), "DD-MM-YY"), PriorPlano(0, a)
                Me.cbo_Priorplan.AddItem PriorPlano(1, a) & "-" & Format(PriorPlano(2, a), "DD-MM-YY")
            Next
        End If
    
    End If
    
End Function
Private Sub cbo_Trialplan_Change()
Dim a As Long
Dim bfound As Boolean
'THIS IS WHERE THE CurPlanoKey should be allocated
    If skipTrialPlanChange = False Then
        For a = 0 To cbo_Trialplan.ListCount - 1
            If cbo_Trialplan.List(a) = cbo_Trialplan Then
                bfound = True
                Exit For
            End If
        Next
        If bfound = False Then cbo_Trialplan = "": CurPlanoKey = "": Exit Sub
       
        If SelPlano > -1 Then
            If Trim(cbo_Trialplan) <> Me.lst_Allocatedplan.List(SelPlano, 3) Then
                Me.lst_Allocatedplan.Selected(SelPlano) = False
                SelPlano = -1
            End If
        End If
        If cbo_Trialplan <> "" Then
            CurPlanoKey = PlanoDic(CStr(cbo_Trialplan))
        End If
        populatePriorPlanoBox
    End If
' how are we dealing with products
'    openProductFrame
End Sub
Private Function openProductFrame()
Dim a As Long, b As Long
Dim bfound As Boolean
Dim prod As Variant
Dim PCde As String
    If cbo_Trialplan = "" Then Me.fme_Product.Visible = False: Exit Function
    If Me.cbo_Store <> "" Then
        getSPACEMANData "PLANO", CBA_BasicFunctions.CBA_DivtoReg(Trim(CStr(cbo_Region.Value))), Mid(cbo_Store, 1, InStr(1, cbo_Store, "-") - 1)
        Me.fme_Product.Visible = True
        Me.lst_Productsel.Clear
        Me.lst_Productallo.Clear
        bfound = False: CurPlanoKey = ""
        If SMoP(0, 0) <> 0 Then
            For a = LBound(SMoP, 2) To UBound(SMoP, 2)
                If CStr(SMoP(1, a)) = CStr(Me.cbo_Trialplan) Then
                    bfound = True
                    CurPlanoKey = SMoP(0, a)
                    Me.lst_Productsel.AddItem SMoP(2, a) & "-" & SMoP(3, a)
                ElseIf bfound = True Then
                    Exit For
                End If
            Next
            For Each prod In AllocatedProducts(SelPlano)
                For b = 0 To Me.lst_Productsel.ListCount - 1
                    PCde = Mid(Me.lst_Productsel.List(b), 1, InStr(1, Me.lst_Productsel.List(b), "-") - 1)
                    If PCde = prod Then
                        lst_Productsel.Selected(b) = True
                        Exit For
                    End If
                Next
            Next
            Call cmd_Addprod_Click
        End If
    End If
'    getSTARData "ProdsInPlano"
'    If oP(0, 0) <> 0 Then
'        For a = LBound(oP, 2) To UBound(oP, 2)
'
'        Next a
'
'    End If
    
    
    
    
End Function
Private Sub cbo_Region_Change()
Dim a As Long, div As Long
Dim s As Variant
Dim desAff As String
Dim bfound As Boolean

'If skipregionchange = False Then
    For a = 0 To cbo_Region.ListCount - 1
        If cbo_Region.List(a) = cbo_Region Then
            bfound = True
            Exit For
        End If
    Next
    If bfound = False Then cbo_Region = "": Exit Sub

    'Application.EnableEvents = True
    'Me.fme_Plan.SetFocus
    If SkipStoreChange = False Then Me.cbo_Store = ""
    If SkipPlanoChange = False Then Me.cbo_Trialplan = "": CurPlanoKey = ""
    If skippriorplanochange = False Then Me.cbo_Priorplan = "": CurPriorPlanoKey = ""
    clearStoPlanoBoxes
    If cbo_Region.Value <> "" Then
        div = NZ(CBA_BasicFunctions.CBA_DivtoReg(Trim(CStr(cbo_Region.Value))))
        getREGIONData "AllStores", div
        If RoP(0, 0) = 0 Then Exit Sub
        Set StoDic = New Scripting.Dictionary
        Set StoNameDic = New Scripting.Dictionary
        For a = LBound(RoP, 2) To UBound(RoP, 2)
            StoDic.Add CStr(RoP(0, a)), CStr(RoP(1, a))
            If StoNameDic.Exists(CStr(RoP(1, a))) = False Then StoNameDic.Add CStr(RoP(1, a)), CStr(RoP(0, a))
            If Me.cbo_Affected.Value = "" Then Me.cbo_Store.AddItem RoP(0, a) & " - " & RoP(1, a)
        Next
        
        If Me.cbo_Affected.Value <> "" Then
            If Me.cbo_Affected.Value = "Yes" Then desAff = "a" Else desAff = "u"
            Set PotentialStoresCol = New Collection
            For Each s In AffStores(div).Keys
                If LCase(AffStores(div)(s)) = desAff Then
                    Me.cbo_Store.AddItem s & " - " & StoDic(CStr(s))
                    PotentialStoresCol.Add s
                End If
            Next
        End If
        
'        If Me.cbo_Region.Value <> "" And Me.cbo_Store.Value = "" And Me.cbo_Trialplan.Value = "" And Me.cbo_Priorplan.Value = "" Then
'            AvailablePlanograms
'        End If
        
        
        
        
        
        
        
        
        
    End If
    AvailablePlanograms
'End If
End Sub
Function clearStoPlanoBoxes()
    Me.cbo_Store.Clear
    Me.cbo_Trialplan.Clear
    Me.cbo_Priorplan.Clear
End Function

Function AvailablePlanograms()
Dim a As Long
Dim s As Variant
Dim strStores As String
Dim curPlanoName As String, curPriorPlanoName As String
    If Me.cbo_Region = "" Then
        Me.cbo_Trialplan = "": CurPlanoKey = "": Me.cbo_Trialplan.Clear
    Else
        If Me.cbo_Affected = "" And Me.cbo_Store = "" Then
            strStores = "ALL"
        Else
            If Me.cbo_Store = "" Then
                strStores = ""
                For Each s In PotentialStoresCol
                    If strStores = "" Then strStores = s Else strStores = strStores & "," & s
                Next
            Else
                For Each s In StoDic
                    If s & " - " & StoDic(s) = Me.cbo_Store Then
                        strStores = s
                        Exit For
                    End If
                Next
            
            End If
        End If
        getSPACEMANData "UNIQUEPLANO", CBA_BasicFunctions.CBA_DivtoReg(Trim(CStr(cbo_Region.Value))), strStores
        If Me.cbo_Trialplan <> "" Then LastPlanoName = Me.cbo_Trialplan
        If Me.cbo_Priorplan <> "" Then lastPriorPlanoName = Me.cbo_Priorplan
        Me.cbo_Trialplan.Clear
        Me.cbo_Trialplan = ""
        Set PlanoDic = New Scripting.Dictionary
        If SMoP(0, 0) <> 0 Then
        For a = LBound(SMoP, 2) To UBound(SMoP, 2)
            If Me.cbo_Store <> "" Then PlanoDic.Add CStr(SMoP(0, a)), CStr(SMoP(1, a))
            Me.cbo_Trialplan.AddItem SMoP(0, a)
            If SMoP(0, a) = LastPlanoName Then
                cbo_Trialplan = SMoP(0, a)
            End If
        Next
        If Me.cbo_Trialplan = "" Then Me.cbo_Priorplan = ""
        If txt_Priordatefrom <> "" And Me.cbo_Trialplan <> "" Then
            populatePriorPlanoBox
            For a = 0 To cbo_Priorplan.ListCount - 1
                If cbo_Priorplan.List(a) = lastPriorPlanoName Then
                    cbo_Priorplan = lastPriorPlanoName
                    Exit For
                End If
            Next
        End If
        
        
        End If
    End If
    
    'For a = 0 To Me.cbo_Store.ListCount - 1
        
        'Debug.Print StoDic.Key(Me.cbo_Store(a))
        
    'Next



End Function
Private Sub cbo_Store_Change()
Dim a As Long
Dim bfound As Boolean
    If SkipStoreChange = False Then
        For a = 0 To cbo_Store.ListCount - 1
            If cbo_Store.List(a) = cbo_Store Then
                bfound = True
                Exit For
            End If
        Next
        If bfound = False Then cbo_Store = "": Exit Sub

    
        'SkipPlanoChange = True
        AvailablePlanograms
        'SkipPlanoChange = False
    End If
End Sub
Private Sub cbo_Affected_Change()
    If cbo_Affected = "Yes" Or cbo_Affected = "No" Or cbo_Affected = "" Then
        If skipregionchange = False Then
            SkipStoreChange = True
            Call cbo_Region_Change
            SkipStoreChange = False
        End If
    Else
        cbo_Affected = ""
    End If
End Sub
Private Sub btn_OK_Click()
    
End Sub
Private Sub cbo_Trialname_Change()
Dim a As Long
Dim yn As Long

If skiptrialnamechange = False Then
    If SelTrial > -1 Then
        Me.fme_Plan.Visible = False
        Me.fme_Product.Visible = False
        yn = MsgBox("Any changes to the previous trial will be lost if not saved. Save now?", vbYesNoCancel)
        Me.fme_Plan.Visible = True
        Me.fme_Product.Visible = False
        If yn = 6 Then
            InterfaceDataToSTAR
        ElseIf yn = 2 Then
            skiptrialnamechange = True
            cbo_Trialname = cbo_Trialname.List(SelTrial)
            skiptrialnamechange = False
            Exit Sub
        End If
    End If

    SelPlano = -1
    populatePlanoBoxes
    SelTrial = -1
    For a = LBound(TrialList, 2) To UBound(TrialList, 2)
        If TrialList(1, a) & "-" & TrialList(0, a) = cbo_Trialname.Value Then
            SelTrial = a
            Exit For
        End If
    Next
    Call populateTrialDates
    Call openPlanFrame
    
    If SelTrial > -1 Then Me.cmd_AddTrial.Visible = False Else Me.cmd_AddTrial.Visible = True
    If SelPlano = -1 Then Me.fme_Product.Visible = False Else Me.fme_Product.Visible = True
End If
End Sub
Private Function populateTrialDates()
    If SelTrial = -1 Then
        Me.txt_Datefrom.Value = ""
        Me.txt_Dateto.Value = ""
        Me.txt_Weeks = ""
        Trial_DateFrom = 0
        Trial_DateTo = 0
    Else
        Trial_DateFrom = TrialList(2, SelTrial)
        Trial_DateTo = TrialList(3, SelTrial)
        Me.txt_Datefrom.Value = Format(Trial_DateFrom, "dddd, mmmm dd, yyyy")
        Me.txt_Dateto.Value = Format(Trial_DateTo, "dddd, mmmm dd, yyyy")
        Me.txt_Weeks = DateDiff("WW", Trial_DateFrom, Trial_DateTo)
    End If
End Function
Private Function openPlanFrame()
Dim a As Long
Dim b As Long
Dim c As Long
Dim d As Long
Dim isopening As Boolean
Dim wasplanokey As String
    If cbo_Trialname = "" Or SelTrial = -1 Then
        Me.fme_Plan.Visible = False: isAllocatedProdsAvailiable = False: Erase AllocatedProducts: Exit Function
    End If
    If Me.fme_Plan.Visible = False Or Me.lst_Allocatedplan.ListCount = 0 Then isopening = True
    Me.lst_Allocatedplan.Clear
    If SkipTrialPlanoQuery = False Then getSTARData "Planogram"
    If PlanoList(0, 0) = 0 Then
    Else
        If PlanoList(0, 0) <> 0 Then
            For a = LBound(PlanoList, 2) To UBound(PlanoList, 2)
                Me.lst_Allocatedplan.AddItem
                Me.lst_Allocatedplan.List(a, 0) = CBA_BasicFunctions.CBA_DivtoReg(PlanoList(0, a))
                getREGIONData "StoreName", PlanoList(0, a), PlanoList(1, a)
                If RoP(0, 0) = 0 Then
                    Me.lst_Allocatedplan.List(a, 1) = "-"
                Else
                    If InStr(1, RoP(0, 0), "- ALCOHOL") > 1 Then
                        Me.lst_Allocatedplan.List(a, 1) = Trim(Mid(RoP(0, 0), 1, InStr(1, RoP(0, 0), "- ALCOHOL") - 1))
                    Else
                        Me.lst_Allocatedplan.List(a, 1) = Trim(RoP(0, 0))
                    End If
                End If
                If PlanoList(2, a) = 0 Then Me.lst_Allocatedplan.List(a, 2) = "No" Else Me.lst_Allocatedplan.List(a, 2) = "Yes"
                'PlanoDic(Trim(PlanoList(3, a)))
                If PlanoList(3, a) <> "" Then
                    getSPACEMANData "PlanogramName", Trim(PlanoList(3, a))
                    If SMoP(0, 0) = 0 Then Me.lst_Allocatedplan.List(a, 3) = "-" Else Me.lst_Allocatedplan.List(a, 3) = SMoP(0, 0)
                End If
                If PlanoList(4, a) <> "" Then
                    getSPACEMANData "ArchivePlanogramName", Trim(PlanoList(4, a))
                    If SMoP(0, 0) = 0 Then Me.lst_Allocatedplan.List(a, 4) = "-" Else Me.lst_Allocatedplan.List(a, 4) = SMoP(0, 0)
                End If
    '            '5 and 6 need to be updated when date from and date to values are availbae in the database?
    '            Me.lst_Allocatedplan.List(a, 5) = NZ(oP(5, a), "-")
    '            Me.lst_Allocatedplan.List(a, 6) = NZ(oP(6, a), "-")
                                
                                
                'NOW THE Prods in plano and allocate to the respective point in the lstbox
                
                If isopening = True Then
                    Set AllocatedProducts(a) = New Scripting.Dictionary
                    SelPlano = a
                    wasplanokey = CurPlanoKey
                    CurPlanoKey = NZ(PlanoList(3, a), "")
                    getSTARData "ProdsInPlano"
                    If op(0, 0) <> 0 Then
                        For c = LBound(op, 2) To UBound(op, 2)
                            AllocatedProducts(a).Add CStr(op(0, c)), CStr(op(0, c))
                        Next c
                    End If
                    CurPlanoKey = wasplanokey
                    SelPlano = -1
                Else
                    If a = UBound(PlanoList, 2) Then
                        If AllocatedProducts(a) Is Nothing Then
                            Set AllocatedProducts(a) = New Scripting.Dictionary
                        End If
'                        getSTARData "ProdsInPlano"
'                        If oP(0, 0) <> 0 Then
'                            For c = LBound(oP, 2) To UBound(oP, 2)
'                                AllocatedProducts(a).Add oP(0, c)
'                            Next c
'                        End If
                    End If
                End If
                
    
            Next
        End If
    End If
    Me.txt_Priordatefrom = Format(DateAdd("YYYY", -1, Trial_DateFrom), "DD/MM/YYYY")
    Me.txt_Priordateto = Format(DateAdd("YYYY", -1, Trial_DateTo), "DD/MM/YYYY")
    Me.fme_Plan.Visible = True
End Function
Private Sub cmd_Newtrial_Click()
    Me.cbo_Trialname.Visible = False
    Me.txt_Trialname.Visible = True
    'Me.cmd_Newtrial.visible = False
    Me.cmd_AddTrial.Visible = True
    SelTrial = -1
    Me.txt_Datefrom = ""
    Me.txt_Dateto = ""
    Me.txt_Weeks = ""
    Trial_DateTo = 0
    Trial_DateFrom = 0
    openPlanFrame
End Sub


Private Sub lst_Allocatedplan_Click()
Dim a As Long
    AffectAllObjects True, "Please Wait..."
    DoEvents
    skipTrialPlanChange = True
    SelPlano = -1
    For a = 0 To Me.lst_Allocatedplan.ListCount - 1
        If Me.lst_Allocatedplan.Selected(a) = True Then
            SelPlano = a
            Exit For
        End If
    Next
    skipregionchange = True: SkipStoreChange = True: SkipPlanoChange = True: skippriorplanochange = True
    populatePlanoBoxes
    skipregionchange = False: SkipStoreChange = False: SkipPlanoChange = False: skippriorplanochange = False
    openProductFrame
    skipTrialPlanChange = False
    AffectAllObjects False
End Sub
Private Function populatePlanoBoxes()
    If SelPlano = -1 Then
        Me.cbo_Region.Value = ""
        Me.cbo_Affected.Value = ""
        Me.cbo_Store.Value = ""
        Me.cbo_Trialplan.Value = ""
        Me.cbo_Priorplan.Value = ""
        Me.lst_Allocatedplan.Clear
    Else
        Me.cbo_Region.Value = CBA_BasicFunctions.CBA_DivtoReg(PlanoList(0, SelPlano))
        If SkipStoreChange = False And PlanoList(2, SelPlano) = 0 Then Me.cbo_Affected.Value = "No" Else Me.cbo_Affected.Value = "Yes"
        getREGIONData "StoreName", PlanoList(0, SelPlano), PlanoList(1, SelPlano)
        If RoP(0, 0) = 0 Then Me.cbo_Store.Value = "-" Else Me.cbo_Store.Value = PlanoList(1, SelPlano) & " - " & Trim(RoP(0, 0))
        getSPACEMANData "PlanogramName", Trim(PlanoList(3, SelPlano))
        If SMoP(0, 0) = 0 Then Me.cbo_Trialplan.Value = "-" Else Me.cbo_Trialplan.Value = SMoP(0, 0)
        If PlanoList(4, SelPlano) <> "" Then
            getSPACEMANData "ArchivePlanogramName", PlanoList(4, SelPlano)
            If SMoP(0, 0) = 0 Then Me.cbo_Priorplan.Value = "-" Else Me.cbo_Priorplan.Value = SMoP(0, 0): CurPriorPlanoKey = PlanoList(4, SelPlano)
        End If
    End If

End Function
'Private Sub txt_Trialname_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'    If KeyCode = 12 Then
'
'
'
'    End If
'End Sub
Private Function getSTARData(ByVal DataRequest As String)
Dim strSQL As String
Dim a As Long
Dim strPlano As String

    Set SRS = New ADODB.Recordset
    If DataRequest = "Trial" Then
        SRS.Open "select * , 1 from tools.dbo.STARMerchandiseTrial", SCN
    ElseIf DataRequest = "AffectedStores" Then
        SRS.Open "select [sasRegion],[sasStore],[sasAffected]  from tools.dbo.STARAffectedstores order by [sasRegion],[sasStore]", SCN
    ElseIf DataRequest = "Planogram" Then
        strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
        strSQL = strSQL & "select mtsMTDRegionID, mtsStore, mtsAffected, mtsCurrentPlanoKey , mtsPriorPlanoKey, NULL , NULL" & Chr(10)
        strSQL = strSQL & "from  tools.dbo.STARMerchandiseTrialStore mts" & Chr(10)
        strSQL = strSQL & "where mtsMTID = " & SelTrial + 1 & Chr(10)
        SRS.Open strSQL, SCN
    ElseIf DataRequest = "ProdsInPlano" Then
        If SelPlano > -1 And CurPlanoKey <> "" Then
            strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
            strSQL = strSQL & "select mtpProductcode from tools.dbo.STARMerchandiseTrialProduct where mtpMTID = " & SelTrial + 1 & Chr(10)
            strSQL = strSQL & "and mtpPlanoKey_ID = " & PlanoList(3, SelPlano) & Chr(10)
            strSQL = strSQL & "and mtpregion = " & PlanoList(0, SelPlano) & Chr(10)
            strSQL = strSQL & "and mtpstoreNo = " & PlanoList(1, SelPlano) & Chr(10)
        Else
            strSQL = "select 0 from tools.dbo.STARMerchandiseTrialStore"
        End If
        SRS.Open strSQL, SCN
    ElseIf DataRequest = "ProdsInTrialPlanos" Then
        strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
        strSQL = strSQL & "select * from tools.dbo.STARMerchandiseTrialProduct where mtpMTID = " & SelTrial + 1 & Chr(10)
        SRS.Open strSQL, SCN
    End If
    
    If SRS.EOF Then
        If DataRequest = "Trial" Then
            ReDim TrialList(0, 0)
            TrialList(0, 0) = 0
        ElseIf DataRequest = "Planogram" Then
            ReDim PlanoList(0, 0)
            PlanoList(0, 0) = 0
        Else
            ReDim op(0, 0)
            op(0, 0) = 0
        End If
    Else
        If DataRequest = "Trial" Then
            TrialList = SRS.GetRows()
        ElseIf DataRequest = "AffectedStores" Then
            Do Until SRS.EOF
                AffStores(SRS(0)).Add CStr(SRS(1)), CStr(SRS(2))
                SRS.MoveNext
            Loop
        ElseIf DataRequest = "ProdsInTrialPlanos" Then

        ElseIf DataRequest = "Planogram" Then
            PlanoList = SRS.GetRows()
        Else
            op = SRS.GetRows()
        End If
    End If
End Function
Private Function InterfaceDataToSTAR()
Dim strSQL As String
Dim RS As ADODB.Recordset
Dim RSS As ADODB.Recordset
Dim S2CN As ADODB.Connection
Dim prod As Variant
Dim a As Long
Dim corSelTrial As Long
Dim r As Long
Dim bfound As Boolean
    
Set S2CN = New ADODB.Connection
With S2CN
    .ConnectionTimeout = 50
    .CommandTimeout = 50
    .Open "Provider=SQLNCLI10;DATA SOURCE=" & CBA_BasicFunctions.TranslateServerName("599DBL12", Date) & ";;INTEGRATED SECURITY=sspi;"
    Set RSS = New ADODB.Recordset
    RSS.Open "SELECT state_desc  FROM sys.databases where name = 'Tools'", SCN
    If LCase(RSS(0)) <> "online" Then
        MsgBox "Due to dependant database states, STAR is currently unavailable. This issue should resolve itslef in 5 mins."
        Exit Function
    End If
End With
    
    
    Set RS = New ADODB.Recordset
    RS.Open "select * from tools.dbo.STARMerchandiseTrial where mtName = '" & TrialList(1, SelTrial) & "'", SCN
    If RS.EOF Then
        'the record does not exist in the database with that name
        If PlanoList(7, SelPlano) = 0 Then
            'the list agrees that the record should not exist in the database
            Set RS = New ADODB.Recordset
            RS.Open "select max(mtID) from tools.dbo.STARMerchandiseTrial", SCN
            If RS.EOF Then
                corSelTrial = 1
            Else
                corSelTrial = RS(0) + 1
            End If
            Set RSS = New ADODB.Recordset
            strSQL = "Insert into tools.dbo.STARMerchandiseTrial(mtID,mtName,mtStartDate,mtEndDate)" & Chr(10)
            strSQL = strSQL & "VALUES(" & corSelTrial & ",'" & TrialList(1, SelTrial) & "', '" & Format(Trial_DateFrom, "YYYY-MM-DD") & "','" & Format(Trial_DateTo, "YYYY-MM-DD") & "')"
            'Debug.Print strSQL
            RSS.Open strSQL, S2CN
            Set RS = New ADODB.Recordset
            RS.Open "select * from tools.dbo.STARMerchandiseTrial where mtID = " & corSelTrial, SCN
            'TrialList(2, SelTrial) = Trial_DateFrom
        Else
            'the list does not agrees that the record should not exist in the database
            MsgBox "An error has been encountred"
            GoTo abortthis
        End If
    End If
    If RS(0) <> SelTrial + 1 Then
        'panic
        MsgBox "There has been an issue. Not interfaced to database"
        GoTo abortthis
    Else
        'UPDATE MAIN ELEMENTS
        If RS(1) <> TrialList(1, SelTrial) Then
            corSelTrial = RS(0)
        Else
            'name differs
            corSelTrial = RS(0)
        End If
        If RS(2) <> Trial_DateFrom Then
            Set RSS = New ADODB.Recordset
            RSS.Open "update tools.dbo.STARMerchandiseTrial set mtStartDate = '" & Format(Trial_DateFrom, "YYYY-MM-DD") & "' where mtID = " & RS(0), S2CN
            TrialList(2, SelTrial) = Trial_DateFrom
        End If
        If RS(3) <> Trial_DateTo Then
            Set RSS = New ADODB.Recordset
            RSS.Open "update tools.dbo.STARMerchandiseTrial set mtEndDate = '" & Format(Trial_DateTo, "YYYY-MM-DD") & "' where mtID = " & RS(0), S2CN
            TrialList(3, SelTrial) = Trial_DateTo
        End If
        ''''NOW LOOK AT THE PLANOGRAMS
        Set RS = New ADODB.Recordset
        RS.Open "select mtgTrialID, mtgPlanoKey from tools.[dbo].[STARMerchandiseTrialPlanogram] where mtgTrialID = " & corSelTrial, SCN, adOpenStatic
        
        'check if planogram is in the database
        If PlanoList(0, 0) <> 0 Then
            For a = LBound(PlanoList, 2) To UBound(PlanoList, 2)
                bfound = False
                If RS.RecordCount > 0 Then RS.MoveFirst
                Do Until RS.EOF
                    If Trim(RS(1)) = Trim(PlanoList(3, a)) Then
                        bfound = True
                        Exit Do
                    End If
                    RS.MoveNext
                Loop
                If bfound = False Then
                    Set RSS = New ADODB.Recordset
                    RSS.Open "insert into tools.[dbo].[STARMerchandiseTrialPlanogram](mtgTrialID,mtgPlanoKey) VALUES(" & corSelTrial & "," & Trim(PlanoList(3, a)) & ")", S2CN
                    Set RS = New ADODB.Recordset: RS.Open "select mtgTrialID, mtgPlanoKey from tools.[dbo].[STARMerchandiseTrialPlanogram] where mtgTrialID = " & corSelTrial, SCN, adOpenStatic
                End If
            Next
        End If
        'check if database has plano that is not in the Planolist
        If PlanoList(0, 0) <> 0 Then
            If RS.RecordCount > 0 Then RS.MoveFirst
            Do Until RS.EOF
                For a = LBound(PlanoList, 2) To UBound(PlanoList, 2)
                    bfound = False
                    If Trim(RS(1)) = Trim(PlanoList(3, a)) Then
                        bfound = True
                        Exit For
                    End If
                Next
                If bfound = False Then
                    Set RSS = New ADODB.Recordset
                    RSS.Open "delete from tools.[dbo].[STARMerchandiseTrialPlanogram] where mtgTrialID = " & corSelTrial & " and mtgPlanoKey = " & Trim(RS(1)), S2CN
                    Set RS = New ADODB.Recordset: RS.Open "select mtgTrialID, mtgPlanoKey from tools.[dbo].[STARMerchandiseTrialPlanogram] where mtgTrialID = " & corSelTrial, SCN, adOpenStatic
                End If
                RS.MoveNext
            Loop
        End If
        'NOW CHECK TO SEE IF THE Region is up to date
        Set RS = New ADODB.Recordset
        RS.Open "select * from tools.dbo.STARMerchandiseTrialregion where mtrMTID = " & corSelTrial, SCN, adOpenStatic
        
        'check if region vs TrialID is in the database
        For a = LBound(PlanoList, 2) To UBound(PlanoList, 2)
            bfound = False
            If RS.RecordCount > 0 Then RS.MoveFirst
            Do Until RS.EOF
                If Trim(CStr(RS(1))) = Trim(PlanoList(0, a)) Then
                    bfound = True
                    Exit Do
                End If
                RS.MoveNext
            Loop
            If bfound = False Then
                Set RSS = New ADODB.Recordset
                RSS.Open "insert into tools.[dbo].[STARMerchandiseTrialregion](mtrMTID,mtrRegionID) VALUES(" & corSelTrial & "," & Trim(PlanoList(0, a)) & ")", S2CN
                Set RS = New ADODB.Recordset: RS.Open "select * from tools.dbo.STARMerchandiseTrialregion where mtrMTID = " & corSelTrial, SCN, adOpenStatic
            End If
        Next
        
        'check if database has region vs TrialID that is not in the Planolist
        If RS.RecordCount > 0 Then RS.MoveFirst
        Do Until RS.EOF
            For a = LBound(PlanoList, 2) To UBound(PlanoList, 2)
                bfound = False
                If Trim(CStr(RS(1))) = Trim(PlanoList(0, a)) Then
                    bfound = True
                    Exit For
                End If
            Next
            If bfound = False Then
                Set RSS = New ADODB.Recordset
                RSS.Open "delete from tools.[dbo].[STARMerchandiseTrialregion] where mtrMTID = " & corSelTrial & " and mtrRegionID = " & Trim(CStr(RS(1))), S2CN
                Set RS = New ADODB.Recordset: RS.Open "select * from tools.dbo.STARMerchandiseTrialregion where mtrMTID = " & corSelTrial, SCN, adOpenStatic
            End If
            RS.MoveNext
        Loop
        
        'NOW CHECK TO SEE IF THE Stores are correct
        If PlanoList(0, 0) <> 0 Then
            Set RS = New ADODB.Recordset
            RS.Open "select * from tools.[dbo].[STARMerchandiseTrialStore] where mtsMTID = " & corSelTrial, SCN, adOpenStatic
            'check if store vs TrialID is in the database
            For a = LBound(PlanoList, 2) To UBound(PlanoList, 2)
                bfound = False
                'If RS.RecordCount > 0 Then RS.MoveFirst
                Do Until RS.EOF
                    If Trim(CStr(RS(1))) = Trim(PlanoList(0, a)) And Trim(CStr(RS(2))) = Trim(PlanoList(1, a)) Then
                        bfound = True
                        Exit Do
                    End If
                    RS.MoveNext
                Loop
                If bfound = False Then
                    Set RSS = New ADODB.Recordset
                    strSQL = "insert into tools.[dbo].[STARMerchandiseTrialStore] (mtsMTID,mtsMTDRegionID,mtsStore,mtsAffected,mtsInclude,mtsCurrentPlanoKey, mtsPriorPlanoKey)" & Chr(10)
                    strSQL = strSQL & "VALUES(" & corSelTrial & "," & Trim(NZ(PlanoList(0, a), "NULL")) & "," & Trim(NZ(PlanoList(1, a), "NULL")) & "," & IIf(AffStores(Trim(PlanoList(0, a)))(Trim(PlanoList(1, a))) = "A", 1, 0) & ", 1," & Trim(NZ(PlanoList(3, a), "NULL")) & "," & Trim(NZ(PlanoList(4, a), "NULL")) & ")"
                    'Debug.Print strSQL
                    RSS.Open strSQL, S2CN
                    Set RS = New ADODB.Recordset: RS.Open "select * from tools.[dbo].[STARMerchandiseTrialStore] where mtsMTID = " & corSelTrial, SCN, adOpenStatic
                Else
                    Set RSS = New ADODB.Recordset
                    strSQL = "Update tools.[dbo].[STARMerchandiseTrialStore]" & Chr(10)
                    strSQL = strSQL & "set mtsAffected = " & IIf(AffStores(Trim(PlanoList(0, a)))(Trim(PlanoList(1, a))) = "A", 1, 0) & ", mtsInclude = 1, mtsCurrentPlanoKey = " & Trim(PlanoList(3, a)) & ", mtsPriorPlanoKey = " & Trim(PlanoList(4, a)) & Chr(10)
                    strSQL = strSQL & "where mtsMTID = " & corSelTrial & " And mtsMTDRegionID = " & Trim(PlanoList(0, a)) & " And mtsStore = " & Trim(PlanoList(1, a))
                    RSS.Open strSQL, S2CN
                    Set RS = New ADODB.Recordset: RS.Open "select * from tools.[dbo].[STARMerchandiseTrialStore] where mtsMTID = " & corSelTrial, SCN, adOpenStatic
                End If
            Next
        End If
        If PlanoList(0, 0) <> 0 Then
            'check if database store vs TrialID inf is in the planomap
            For a = LBound(PlanoList, 2) To UBound(PlanoList, 2)
                If RS.RecordCount > 0 Then RS.MoveFirst
                bfound = False
                Do Until RS.EOF
                    If Trim(CStr(RS(1))) = Trim(PlanoList(0, a)) And Trim(CStr(RS(2))) = Trim(PlanoList(1, a)) Then
                        bfound = True
                        Exit Do
                    End If
                    RS.MoveNext
                Loop
                If bfound = False Then
                    Set RSS = New ADODB.Recordset
                    'Debug.Print "delete from tools.[dbo].[STARMerchandiseTrialStore] where mtsMTID = " & corSelTrial & " and mtsMTDRegionID = " & Trim(PlanoList(0, a)) & " and mtsStore = " & Trim(PlanoList(1, a))
                    RSS.Open "delete from tools.[dbo].[STARMerchandiseTrialStore] where mtsMTID = " & corSelTrial & " and mtsMTDRegionID = " & Trim(PlanoList(0, a)) & " and mtsStore = " & Trim(PlanoList(1, a)), S2CN
                    Set RS = New ADODB.Recordset: RS.Open "select * from tools.dbo.STARMerchandiseTrialregion where mtrMTID = " & corSelTrial, SCN, adOpenStatic
                End If
            Next
        End If
        
        If PlanoList(0, 0) <> 0 Then
        'NOW CHECK TO SEE IF THE Products are correct
            
            Set RS = New ADODB.Recordset
            RS.Open "select * from tools.[dbo].[STARMerchandiseTrialProduct] where mtpMTID = " & corSelTrial, SCN, adOpenStatic
            'check if store vs TrialID is in the database
            For a = LBound(PlanoList, 2) To UBound(PlanoList, 2)
                For Each prod In AllocatedProducts(a)
                    If RS.RecordCount > 0 Then RS.MoveFirst
                    bfound = False
                    Do Until RS.EOF
                        If Trim(CStr(RS(1))) = CStr(prod) And Trim(CStr(RS(2))) = Trim(PlanoList(3, a)) And CStr(RS(3)) = Trim(PlanoList(0, a)) And Trim(CStr(RS(4))) = Trim(PlanoList(1, a)) Then
                            bfound = True
                            Exit Do
                        End If
                        RS.MoveNext
                    Loop
                    If bfound = False Then
                        Set RSS = New ADODB.Recordset
                        strSQL = "insert into tools.[dbo].[STARMerchandiseTrialProduct] (mtpMTID,mtpRegion,mtpStoreNo,mtpPlanoKey_ID,mtpProductCode)" & Chr(10)
                        strSQL = strSQL & "VALUES(" & corSelTrial & "," & Trim(PlanoList(0, a)) & "," & Trim(PlanoList(1, a)) & "," & Trim(PlanoList(3, a)) & "," & prod & ")"
                        'Debug.Print strSQL
                        RSS.Open strSQL, S2CN
                        Set RS = New ADODB.Recordset: RS.Open "select * from tools.[dbo].[STARMerchandiseTrialProduct] where mtpMTID = " & corSelTrial, SCN, adOpenStatic
                    End If
                Next
            Next
        End If
        
        'check if database store vs TrialID inf is in the planomap
        
        If RS.RecordCount > 0 Then RS.MoveFirst
        Do Until RS.EOF
            bfound = False
            If PlanoList(0, 0) <> 0 Then
                For a = LBound(PlanoList, 2) To UBound(PlanoList, 2)
                    For Each prod In AllocatedProducts(a)
                        If Trim(CStr(RS(1))) = CStr(prod) And Trim(CStr(RS(2))) = Trim(PlanoList(3, a)) And CStr(RS(3)) = Trim(PlanoList(0, a)) And Trim(CStr(RS(4))) = Trim(PlanoList(1, a)) Then
                            bfound = True
                            Exit For
                        End If
                    Next
                    If bfound = True Then Exit For
                Next
            End If
            If bfound = False Then
                Set RSS = New ADODB.Recordset
                strSQL = "delete from tools.dbo.STARMerchandiseTrialProduct where mtpMTID = " & corSelTrial & Chr(10)
                strSQL = strSQL & "and mtpProductCode = " & RS(1) & " and mtpPlanoKey_ID = " & Trim(CStr(RS(2))) & " and mtpRegion = " & Trim(CStr(RS(3))) & " and mtpStoreNo = " & Trim(CStr(RS(4)))
                Debug.Print strSQL
                RSS.Open strSQL, S2CN
                Set RS = New ADODB.Recordset: RS.Open "select * from tools.[dbo].[STARMerchandiseTrialProduct] where mtpMTID = " & corSelTrial, SCN, adOpenStatic
            End If
            RS.MoveNext
        Loop
        MsgBox "Saved to Database"
    End If


abortthis:
    S2CN.Close
    Set S2CN = Nothing
End Function
Private Function getSPACEMANData(ByVal DataRequest As String, Optional ByVal LongVal As Long, Optional ByVal StringVal As String, Optional ByVal VariantVal As Variant)
Dim strSQL As String
    Set SMRS = New ADODB.Recordset
    If DataRequest = "PlanogramName" Then
        If LongVal <> 0 Then
            SMRS.Open "select Planogram from Spaceman.dbo.SC_PLANO_KEY pl where pl.[KEY_ID] =" & LongVal, SMCN
        End If
    ElseIf DataRequest = "ArchivePlanogramName" Then
        If LongVal <> 0 Then
            SMRS.Open "select Planogram + '-' + convert(nvarchar(8),ModifiedDate,5) from SpacemanArchive.dbo.SC_PLANO_KEY pl where pl.[KEY_ID] =" & LongVal, SMCN
        End If
    ElseIf DataRequest = "PlanogramNumber" Then
        If StringVal = "TrialPlan" Then
            SMRS.Open "select KEY_ID from #PSP where Planogram ='" & Me.cbo_Trialplan & "' and region = " & LongVal & " and storeno = " & StoNameDic(CStr(Me.cbo_Store)), SMCN
        ElseIf StringVal = "PriorTrialPlan" Then
            SMRS.Open "select KEY_ID from #PSP where Planogram ='" & Me.cbo_Priorplan & "' and region = " & LongVal & " and storeno = " & StoNameDic(CStr(Me.cbo_Store)), SMCN
        End If
    ElseIf DataRequest = "PLANO" Then
        If StringVal <> "" And LongVal <> 0 Then
            strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
            strSQL = strSQL & "select * from #PSP where region = " & LongVal & Chr(10)
            If StringVal <> "ALL" Then strSQL = strSQL & "and storeno in (" & StringVal & ")" & Chr(10)
            strSQL = strSQL & "order by key_id" & Chr(10)
        Else
            strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
            strSQL = strSQL & "select * from #PSP where PLANOFOLDER03 = 'DontReturnAnything' "
        End If
        SMRS.Open strSQL, SMCN
    ElseIf DataRequest = "ALLPLANONOSTORE" Then
        If LongVal <> 0 Then
            strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
            strSQL = strSQL & "select distinct storeno from #PSP where KEY_ID = " & CurPlanoKey & " and region = " & LongVal & Chr(10)
            strSQL = strSQL & "order by Storeno" & Chr(10)
        End If
        SMRS.Open strSQL, SMCN
    ElseIf DataRequest = "UNIQUEPLANO" Then
        If StringVal <> "" And LongVal <> 0 Then
            strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
            strSQL = strSQL & "select distinct PLANOGRAM, key_ID from #PSP where region = " & LongVal & Chr(10)
            If StringVal <> "ALL" Then strSQL = strSQL & "and storeno in (" & StringVal & ")"
        Else
            strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
            strSQL = strSQL & "select distinct PLANOGRAM from #PSP where PLANOFOLDER03 = 'DontReturnAnything' "
        End If
        SMRS.Open strSQL, SMCN
    ElseIf DataRequest = "ArchivePlano" Then
        If CurPlanoKey = "" Or (Me.txt_Priordatefrom = "" And Me.txt_Priordateto = "") Then
            strSQL = "select 0"
        Else
            strSQL = "SET NOCOUNT ON" & Chr(10) & "SET ANSI_WARNINGS OFF" & Chr(10)
            strSQL = strSQL & "select ak.KEY_ID, ak.PLANOGRAM, convert(date,ak.MODIFIEDDATE) as ModifiedDate" & Chr(10)
            strSQL = strSQL & "from Spaceman.dbo.SC_PLANO_KEY pk" & Chr(10)
            strSQL = strSQL & "left join SpacemanArchive.dbo.SC_PLANO_KEY ak on ak.PLANOFOLDER01 = pk.PLANOFOLDER01 and ak.PLANOFOLDER02 = pk.PLANOFOLDER02" & Chr(10)
            strSQL = strSQL & "and charindex(pk.PLANOFOLDER03,ak.PLANOFOLDER03,1) > 0 and ak.PLANOFOLDER04 = pk.PLANOFOLDER04 and ak.PLANOFOLDER05 = pk.PLANOFOLDER05" & Chr(10)
            strSQL = strSQL & "where PK.key_id = " & CurPlanoKey & Chr(10)
            If Me.txt_Priordatefrom <> "" Then strSQL = strSQL & "and ak.MODIFIEDDATE >= '" & Format(Me.txt_Priordatefrom, "YYYY-MM-DD") & "'" & Chr(10)
            If Me.txt_Priordateto <> "" Then strSQL = strSQL & "and ak.MODIFIEDDATE <= '" & Format(Me.txt_Priordateto, "YYYY-MM-DD") & "'" & Chr(10)
            strSQL = strSQL & "order by ak.PLANOGRAM, ak.MODIFIEDDATE" & Chr(10)
        End If
        SMRS.Open strSQL, SMCN
    End If
    If DataRequest = "ArchivePlano" Then
        If SMRS.State = 0 Then ReDim PriorPlano(0, 0): PriorPlano(0, 0) = 0: Exit Function
        If SMRS.EOF Then
            ReDim PriorPlano(0, 0)
            PriorPlano(0, 0) = 0
        Else
            PriorPlano = SMRS.GetRows()
        End If
    Else
        If SMRS.State = 0 Then ReDim SMoP(0, 0): SMoP(0, 0) = 0: Exit Function
        If SMRS.EOF Then
            ReDim SMoP(0, 0)
            SMoP(0, 0) = 0
        Else
            SMoP = SMRS.GetRows()
        End If
    End If
End Function
Private Function getCBISData(ByVal DataRequest As String, Optional ByVal LongVal As Long, Optional ByVal StringVal As String, Optional ByVal VariantVal As Variant)
    Set CRS = New ADODB.Recordset
    If DataRequest = "PlanogramName" Then
        If LongVal <> 0 Then
            CRS.Open "select Planogram from Spaceman.dbo.SC_PLANO_KEY pl where pl.[KEY_ID] =" & LongVal, CCN
        End If
    End If
    
    If CRS.State = 0 Then ReDim CoP(0, 0): CoP(0, 0) = 0: Exit Function
    If CRS.EOF Then
        ReDim CoP(0, 0)
        CoP(0, 0) = 0
    Else
        CoP = CRS.GetRows()
    End If

    
End Function
Private Function getREGIONData(ByVal DataRequest As String, ByVal div As Long, Optional ByVal LongVal As Long, Optional ByVal StringVal As String, Optional ByVal VariantVal As Variant)

    Set RRS(div) = New ADODB.Recordset
    If DataRequest = "StoreName" Then
        If LongVal <> 0 Then RRS(div).Open "select City  from purchase.dbo.Store where storeno = " & LongVal, RCN(div)
    ElseIf DataRequest = "AllStores" Then
        RRS(div).Open "select StoreNo, City  from purchase.dbo.Store", RCN(div)
    End If
    
    If RRS(div).State = 0 Then Erase RoP: Exit Function
    If RRS(div).EOF Then
        ReDim RoP(0, 0)
        RoP(0, 0) = 0
    Else
        RoP = RRS(div).GetRows()
    End If

End Function
Function AffectAllObjects(ByVal Hidden As Boolean, Optional ByVal Message As String)
Dim con As Object
    
    For Each con In Me.Controls
        If Hidden = True Then
            If con.Name <> "lbl_PleaseClose" Then
                con.Visible = False
            Else
                con.Caption = Message
                Me.lbl_PleaseClose.ForeColor = xlWhite
                con.Visible = True
            End If
        Else
            If con.Name <> "lbl_PleaseClose" Then
                con.Visible = True
            Else
                con.Caption = ""
                con.Visible = False
            End If
        End If
    Next

    
End Function
Private Sub UserForm_Terminate()
Dim div As Long
    Set SRS = Nothing
    Set SMRS = Nothing
    Set CRS = Nothing
    For div = 501 To 509
        Set RRS(div) = Nothing
    Next
    SCN.Close
    Set SCN = Nothing
    SMCN.Close
    Set SMCN = Nothing
    For div = 501 To 509
        If div = 508 Then div = 509
        On Error Resume Next
        If RCN(div).State = 1 Then RCN(div).Close
        Err.Clear
        On Error GoTo 0
        Set RCN(div) = Nothing
    Next
    Erase TrialList

End Sub
