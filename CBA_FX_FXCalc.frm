VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CBA_FX_FXCalc 
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14745
   OleObjectBlob   =   "CBA_FX_FXCalc.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "CBA_FX_FXCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bx_pop_Cost As Boolean, bx_pop_Period As Boolean, bx_pop_SupCurr As Boolean, bx_pop_Advice As Boolean, bx_pop_SupRate As Boolean, bx_pop_Hedge As Boolean, bx_pop_Diff As Boolean, bx_pop_QTY As Boolean
Private bx_pop_Cost_Val As Single, bx_pop_SupRate_Val As Single, bx_pop_SupCurr_Val As String, bx_pop_Period_Val As Date, bx_pop_Hedge_Val As String, bx_pop_QTY_Val As Single
Private fx_PortID As String, fx_VerID As Byte, fx_ContractNo As Long, fx_PCode As Long, fx_OSD As Date, fx_Cost As Single, fx_Qty As Single, fx_Curr As String, fx_PortName As String, fx_VerName As String
Private fx_DataExisits As Boolean
Private bx_FXrate_Val As Single
Private fx_Orig_Curr As String
Private fx_Orig_Cost As Single

Private Sub UserForm_Activate()
    If fx_DataExisits = True Then
        bx_Period = Month(DateAdd("D", 14, fx_OSD)) & "-" & Year(DateAdd("D", 14, fx_OSD))
        bx_SupCurr = fx_Curr
        If fx_Curr = "EUR" Or fx_Curr = "GBP" Or fx_Curr = "USD" Then
            bx_Cost = ""
            Me.bx_SupRate = fx_Cost
        '    fx_Cost = bx_Cost
        Else
            bx_Cost = fx_Cost
        End If
        bx_QTY = fx_Qty
        
    End If
End Sub
Private Sub bx_Period_Change()
    Call checkData
End Sub
Private Sub bx_SupCurr_Change()
    Call checkData
End Sub
Private Sub bx_Cost_AfterUpdate()
    Call checkData
End Sub
Private Sub bx_SupRate_AfterUpdate()
    Call checkData
End Sub

Private Sub btn_OK_Click()
Unload Me
End Sub

Sub UserForm_Initialize()
    Dim lTop, lLeft, lRow, lcol As Long
    Dim DT As Date
    Dim bytCurMonth As Byte
    Me.StartUpPosition = 0
    Me.Top = Application.Top + (Application.Height / 2) - (Me.Height / 2)
    Me.Left = Application.Left + (Application.Width / 2) - (Me.Width / 2)
    Me.bx_SupCurr.AddItem "AUD"
    Me.bx_SupCurr.AddItem "USD"
    Me.bx_SupCurr.AddItem "EUR"
    Me.bx_SupCurr.AddItem "GBP"
    'Me.bx_SupCurr.AddItem "NZD"
    For DT = DateAdd("M", -1, Date) To DateAdd("M", 18, Date)
        If Month(DT) <> bytCurMonth Then
            bytCurMonth = Month(DT)
            bx_Period.AddItem Month(DT) & "-" & Year(DT)
        End If
    Next
'    Me.Hide
End Sub
Function setPortDetails(Optional ByVal PortID As String, Optional ByVal VerID As Byte, Optional ByVal ContractNo As Long, Optional ByVal PCode As Long, Optional ByVal OSD As Date, _
    Optional ByVal Cost As Single, Optional ByVal QTY As Single, Optional ByVal Curr As String, Optional ByVal PortName As String, Optional ByVal VerName As String)
Dim CN As ADODB.Connection
Dim RS As ADODB.Recordset
Dim strSQL
If PortID = "" And VerID = 0 And ContractNo = 0 And PCode = 0 Then
    fx_DataExisits = False
Else
    fx_PortID = PortID
    fx_VerID = VerID
    fx_PortName = PortName
    fx_VerName = VerName
    fx_ContractNo = ContractNo
    fx_PCode = PCode
    fx_OSD = OSD
    fx_Curr = Curr
    fx_Cost = Cost
    fx_Qty = QTY
    fx_Orig_Curr = Curr: fx_Orig_Cost = fx_Cost
    fx_DataExisits = True
End If


End Function

Function checkData() As Boolean
    If Me.bx_Cost <> "" And IsNumeric(Me.bx_Cost) Then bx_pop_Cost = True Else bx_pop_Cost = False
    If Me.bx_SupCurr <> "" Then bx_pop_SupCurr = True Else bx_pop_SupCurr = False
    If Me.bx_SupRate <> "" And IsNumeric(Me.bx_SupRate) Then bx_pop_SupRate = True Else bx_pop_SupRate = False
    If Me.bx_Hedge <> "" And IsNumeric(Me.bx_Hedge) Then bx_pop_Hedge = True Else bx_pop_Hedge = False
    If Me.bx_Diff <> "" And IsNumeric(Me.bx_Diff) Then bx_pop_Diff = True Else bx_pop_Diff = False
    If Me.bx_Advice <> "" Then bx_pop_Advice = True Else bx_pop_Advice = False
    If Me.bx_Period <> "" Then bx_pop_Period = True Else bx_pop_Period = False
    If Me.bx_QTY <> "" Then bx_pop_QTY = True Else bx_pop_QTY = False
    If bx_pop_QTY = True And bx_pop_Cost = True And bx_pop_Period = True And bx_pop_SupCurr = True And bx_pop_SupRate = True Then updateQuery
End Function
Function updateQuery()
Dim CN As ADODB.Connection
Dim RS As ADODB.Recordset
Dim strSQL As String
Dim a As Long
    
    If Me.bx_Hedge = bx_pop_Hedge_Val And bx_pop_Cost_Val = bx_Cost And bx_pop_SupRate_Val = bx_SupRate And bx_pop_SupCurr_Val = bx_SupCurr And bx_pop_Period_Val = bx_Period Then
            
            
            
    Else
        
    
        
    
        bx_pop_Cost_Val = bx_Cost
        bx_pop_SupRate_Val = bx_SupRate
        bx_pop_SupCurr_Val = bx_SupCurr
        bx_pop_Period_Val = bx_Period
        bx_pop_QTY_Val = bx_QTY
        bx_FXrate_Val = getFXRate
        bx_pop_Hedge_Val = bx_FXrate_Val
        If bx_pop_Hedge_Val <> 0 Then
            
            Me.bx_Hedge = Round(bx_pop_SupRate_Val / bx_pop_Cost_Val, 4)
            'Me.bx_Hedge = bx_pop_Hedge_Val
            'Me.bx_Diff = Format(((bx_pop_QTY_Val * bx_pop_Cost_Val) * bx_pop_SupRate_Val) - ((bx_pop_QTY_Val * bx_pop_Cost_Val) * bx_pop_Hedge_Val), "$#,0.00")
            bx_ALDIrate = bx_FXrate_Val
            Me.bx_Diff = Format(-Me.bx_QTY * ((bx_pop_SupRate_Val / bx_FXrate_Val) - bx_pop_Cost_Val), "$#,0")
            'Debug.Print ((bx_pop_SupRate_Val / bx_FXrate_Val) - bx_pop_Cost_Val)
            
            If Me.bx_Diff < 0 Then
                Me.bx_Advice = "AUD"
            ElseIf Me.bx_Diff > 0 Then
                Me.bx_Advice = bx_pop_SupCurr_Val
            ElseIf Me.bx_Diff = 0 Then
                Me.bx_Advice = ""
            End If
        End If

        

    End If


End Function
Function getFXRate() As Single
Dim CN As ADODB.Connection
Dim RS As ADODB.Recordset
Dim strSQL As String
        If bx_SupCurr = "" Or bx_Period = "" Then
            getFXRate = 0
        Else
            Set CN = New ADODB.Connection
            Set RS = New ADODB.Recordset
            With CN
                .ConnectionTimeout = 50
                .CommandTimeout = 50
                .Open "Provider=" & CBA_MSAccess & ";DATA SOURCE=" & CBA_BSA & "LIVE DATABASES\ABI.accdb;"
            End With
            
            strSQL = "Select Rate from FXCurrentData where Yearno = " & Year(bx_Period) & " and MonthNo = " & Month(bx_Period) & " and CurrencyTo = """ & bx_SupCurr & """"
            RS.Open strSQL, CN
            If RS.EOF Then
                MsgBox "No Hedge Rate Found" & Chr(10) & Chr(10) & "Please contact Buying Administration", vbOKOnly
                getFXRate = 0
            Else
                getFXRate = RS(0)
            End If
    
            Set RS = Nothing
            Set CN = Nothing
        End If
        
End Function
