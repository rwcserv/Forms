VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CBA_FX_PortfolioSelector 
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14685
   OleObjectBlob   =   "CBA_FX_PortfolioSelector.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CBA_FX_PortfolioSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private arr() As Variant
Private loadArr() As String
Private ArrFilter(1 To 4) As String

'#RW Added new mousewheel routines 190701
Private Sub ListBox1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call mw_SetBoxHook(ListBox1)
End Sub

Private Sub cb_OpenCalc_Click()
Dim a As Long
Dim VerName As String, PortName As String, Curr As String
Dim QTY As Long, ContNo As Long, PCode As Long
Dim OSD As Date
Dim Cost As Single, PortID As String, VerID As Byte







    
    
    For a = 0 To Me.ListBox1.ListCount - 1
        If Me.ListBox1.Selected(a) = True Then
            PortName = Me.ListBox1.List(a, 0)
            VerName = Me.ListBox1.List(a, 1)
            Exit For
        End If
    Next
    
    
    
    
    If PortName <> "" Then
        On Error Resume Next
        If UBound(arr, 2) < 0 Then
            Err.Clear
            On Error GoTo 0
            GoTo noArrAvailable
        Else
            On Error GoTo 0
            For a = LBound(arr, 2) To UBound(arr, 2)
                If arr(0, a) = PortName And arr(1, a) = VerName Then
                    If IsNull(arr(3, a)) = False Then PortID = CStr(arr(3, a))
                    If IsNull(arr(4, a)) = False Then VerID = CByte(arr(4, a))
                    If IsNull(arr(5, a)) = False Then ContNo = CLng(arr(5, a))
                    If IsNull(arr(6, a)) = False Then PCode = CLng(arr(6, a))
                    If IsNull(arr(7, a)) = False Then Curr = CStr(arr(7, a))
                    If IsNull(arr(11, a)) = True Then QTY = 0 Else QTY = CSng(arr(11, a))
                    If IsNull(arr(2, a)) = False Then OSD = CDate(arr(2, a))
                    If IsNull(arr(8, a)) = True Then Cost = 0 Else Cost = CSng(arr(8, a))
                    CBA_FX_FXCalc.setPortDetails PortID, VerID, ContNo, PCode, OSD, Cost, QTY, Curr, PortName, VerName
                    Exit For
                End If
            Next
        End If
        On Error GoTo 0
        
        

    End If
noArrAvailable:
    Unload Me
    CBA_FX_FXCalc.Show
    
End Sub
Private Function FilterList()
Dim a As Long, b As Long
Dim FilteredArray As Variant
Dim TempArr() As Variant

FilteredArray = arr
        
        If ArrFilter(1) <> "" Then
            TempArr = FilteredArray
            ReDim FilteredArray(0 To UBound(arr, 1), 0 To 0)
            For a = LBound(TempArr, 2) To UBound(TempArr, 2)
                If InStr(1, LCase(TempArr(0, a)), LCase(ArrFilter(1))) > 0 Then
                    ReDim Preserve FilteredArray(0 To UBound(arr, 1), UBound(FilteredArray, 2) + 1)
                    For b = LBound(TempArr, 1) To UBound(TempArr, 1)
                        FilteredArray(b, UBound(FilteredArray, 2)) = TempArr(b, a)
                    Next
                End If
            Next
        End If
        If ArrFilter(2) <> "" Then
            TempArr = FilteredArray
            ReDim FilteredArray(0 To UBound(arr, 1), 0 To 0)
            For a = LBound(TempArr, 2) To UBound(TempArr, 2)
                If InStr(1, LCase(TempArr(1, a)), LCase(ArrFilter(2))) > 0 Then
                    ReDim Preserve FilteredArray(0 To UBound(arr, 1), UBound(FilteredArray, 2) + 1)
                    For b = LBound(TempArr, 1) To UBound(TempArr, 1)
                        FilteredArray(b, UBound(FilteredArray, 2)) = TempArr(b, a)
                    Next
                End If
            Next
        End If
        
        
        If ArrFilter(3) <> "" Then
            TempArr = FilteredArray
            ReDim FilteredArray(0 To UBound(arr, 1), 0 To 0)
            For a = LBound(TempArr, 2) To UBound(TempArr, 2)
                If InStr(1, LCase(TempArr(5, a)), LCase(ArrFilter(3))) > 0 Then
                    ReDim Preserve FilteredArray(0 To UBound(arr, 1), UBound(FilteredArray, 2) + 1)
                    For b = LBound(TempArr, 1) To UBound(TempArr, 1)
                        FilteredArray(b, UBound(FilteredArray, 2)) = TempArr(b, a)
                    Next
                End If
            Next
        
        End If
        If ArrFilter(4) <> "" Then
            TempArr = FilteredArray
            ReDim FilteredArray(0 To UBound(arr, 1), 0 To 0)
            For a = LBound(TempArr, 2) To UBound(TempArr, 2)
                If InStr(1, LCase(TempArr(6, a)), LCase(ArrFilter(4))) > 0 Then
                    ReDim Preserve FilteredArray(0 To UBound(arr, 1), UBound(FilteredArray, 2) + 1)
                    For b = LBound(TempArr, 1) To UBound(TempArr, 1)
                        FilteredArray(b, UBound(FilteredArray, 2)) = TempArr(b, a)
                    Next
                End If
            Next
        End If




        ReDim loadArr(0 To UBound(FilteredArray, 2), 0 To 3)
        b = -1
        For a = LBound(FilteredArray, 2) To UBound(FilteredArray, 2)
            If IsNull(FilteredArray(0, a)) = False And IsEmpty(FilteredArray(0, a)) = False Then
                b = b + 1
                If IsNull(FilteredArray(0, a)) Then loadArr(b, 0) = "" Else loadArr(b, 0) = FilteredArray(0, a)
                If IsNull(FilteredArray(1, a)) Then loadArr(b, 1) = "" Else loadArr(b, 1) = FilteredArray(1, a)
                If IsNull(FilteredArray(5, a)) Then loadArr(b, 2) = "" Else loadArr(b, 2) = FilteredArray(5, a)
                If IsNull(FilteredArray(6, a)) Then loadArr(b, 3) = "" Else loadArr(b, 3) = FilteredArray(6, a)
            End If
        Next a
        
        Me.ListBox1.List = loadArr





End Function
Private Sub Search_Contract_Change()
    ArrFilter(3) = Search_Contract
    Call FilterList
End Sub

Private Sub Search_Portfolio_Change()
    ArrFilter(1) = Search_Portfolio
    Call FilterList
End Sub
Private Sub Search_Product_Change()
Dim a As Long
    ArrFilter(4) = Search_Product
    Call FilterList
End Sub
Private Sub Search_Version_Change()
    ArrFilter(2) = Search_Version
    Call FilterList
End Sub

Private Sub TextBox1_AfterUpdate()
Dim OSD As Date
Dim CN As ADODB.Connection
Dim RS As ADODB.Recordset
Dim a As Long, b As Long
Dim strSQL As String



Me.ListBox1.Clear



If g_IsDate(TextBox1) And TextBox1 <> "" Then
    If Year(TextBox1) < Year(Date) Then
        MsgBox "Date is in a previous year"
        Exit Sub
    End If
    
    OSD = TextBox1
    
    
    Set CN = New ADODB.Connection
    Set RS = New ADODB.Recordset
    
    
        With CN
            .ConnectionTimeout = 100
            .CommandTimeout = 100
            .Open "Provider=SQLNCLI10;DATA SOURCE=599DBL01;;INTEGRATED SECURITY=sspi;"
        End With
    
    strSQL = "select pfl.Description as PortName,pfvl.Description as VersionName,GroupAdvdate as OSD, pfvr.Portfolioid, pfvm.PfVersionID, pfvm.ContractNo, isnull(convert(nvarchar(10),pfvm.ProductCode), pfvr.Productcode) as productcode, pfv.CostCurrencyPickup , isnull(convert(nvarchar(10), pfv.CostPickup), pfvr.cost1),pfv.freight, pfvr.retail1, pfvr.quantity1, pfvl.supplier" & Chr(10)
    strSQL = strSQL & "from cbis599p.Portfolio.Rep_PfVersionReg pfvr" & Chr(10)
    strSQL = strSQL & "left join cbis599p.portfolio.PfVersionMapping pfvm on pfvm.PortfolioID = pfvr.Portfolioid and pfvm.PfVersionID = pfvr.PfVersionid" & Chr(10)
    strSQL = strSQL & "left join cbis599p.portfolio.PortfolioLng  pfl on pfl.PortfolioID = pfvr.Portfolioid and pfl.LanguageID = 0" & Chr(10)
    strSQL = strSQL & "left join cbis599p.portfolio.PfVersionLng pfvl on pfvl.PortfolioID = pfvr.Portfolioid and pfvl.PfVersionID = pfvr.PfVersionid" & Chr(10)
    strSQL = strSQL & "left join cbis599p.portfolio.PfVersionReg pfv on pfv.PortfolioID = pfvr.Portfolioid and pfv.PfVersionID = pfvr.PfVersionid" & Chr(10)
    strSQL = strSQL & "where pfvr.GroupAdvdate = '" & Format(OSD, "YYYY-MM-DD") & "'" & Chr(10)
    
    
    RS.Open strSQL, CN
    If RS.EOF Then
        MsgBox "No Data Returned", vbOKOnly
    Else
        arr = RS.GetRows()
        ReDim loadArr(0 To UBound(arr, 2), 0 To 3)
        b = -1
        For a = LBound(arr, 2) To UBound(arr, 2)
            If IsNull(arr(0, a)) = False Then
                b = b + 1
                If IsNull(arr(0, a)) Then loadArr(b, 0) = "" Else loadArr(b, 0) = arr(0, a)
                If IsNull(arr(1, a)) Then loadArr(b, 1) = "" Else loadArr(b, 1) = arr(1, a)
                If IsNull(arr(5, a)) Then loadArr(b, 2) = "" Else loadArr(b, 2) = arr(5, a)
                If IsNull(arr(6, a)) Then loadArr(b, 3) = "" Else loadArr(b, 3) = arr(6, a)
            End If
        Next a
        
        Me.ListBox1.List = loadArr
        
        
        
    
    End If
    
    
    Set CN = Nothing
    Set RS = Nothing
Else
    MsgBox "Invalid Entry"

End If


End Sub

Sub UserForm_Initialize()
    Dim lTop, lLeft, lRow, lcol As Long
    Me.StartUpPosition = 0
    Me.Top = Application.Top + (Application.Height / 2) - (Me.Height / 2)
    Me.Left = Application.Left + (Application.Width / 2) - (Me.Width / 2)
End Sub
