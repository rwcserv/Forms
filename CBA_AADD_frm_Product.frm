VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CBA_AADD_frm_Product 
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9480
   OleObjectBlob   =   "CBA_AADD_frm_Product.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "CBA_AADD_frm_Product"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private DataCM As CBA_AADD_Product

'#RW Added new mousewheel routines 190701
Private Sub lbx_POSDATA_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call mw_SetBoxHook(lbx_POSDATA)
End Sub
Private Sub cbx_POSView_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call mw_SetBoxHook(cbx_POSView)
End Sub

Private Sub but_Export_Click()
    Dim RunTo As Byte
    Dim a As Long, b As Long
    If Me.cbx_POSView = "Store Level" Then RunTo = 4
    Workbooks.Add
    With ActiveSheet
        For a = 0 To Me.lbx_POSDATA.ListCount - 1
            For b = 0 To RunTo
                .Cells(a + 1, b + 1) = Me.lbx_POSDATA.Column(b, a)
            Next
        Next
    End With
End Sub

Private Sub UserForm_Terminate()
    Unload Me
End Sub
Private Sub btn_OK_Click()
    Unload Me
End Sub
Sub UserForm_Initialize()
    Me.StartUpPosition = 0
    Me.Top = Application.Top + (Application.Height / 2) - (Me.Height / 2)
    Me.Left = Application.Left + ((Application.Width / 8) * 5)
    Me.cbx_POSView.AddItem "National"
    Me.cbx_POSView.AddItem "Regional"
    Me.cbx_POSView.AddItem "Store Level"

End Sub
Property Set CBA_AADD_frm_ProductSetup(ByRef cls_Data As CBA_AADD_Product)
    Set DataCM = cls_Data
    If DataCM Is Nothing Then
    Else
        Me.lbl_CG = DataCM.CGno & "-" & DataCM.CGDescription
        Me.lbl_SCG = DataCM.SCGno & "-" & DataCM.SCGDescription
        Me.lbl_productcode = DataCM.Product
        Me.lbl_description = DataCM.ProductDesc
    End If
End Property
Property Get CGno()
    Set CGno = DataCM.CGno
End Property
Property Get SCGno()
    Set SCGno = DataCM.SCGno
End Property
Property Get Product()
    Product = DataCM.Product
End Property
Private Sub cbx_POSView_Change()
    Dim ListData() As Variant
    Dim Data As Variant
    Dim a As Long, b As Long, ListLine As Long
    Dim Pstart As Date, Pend As Date
    
    Pstart = DataCM.PromoStartDate
    Pend = DataCM.PromoEndDate

    Select Case cbx_POSView.Value
        Case "National"
            Me.lbx_POSDATA.Clear
            Data = DataCM.getdata
                    
        
        Case "Regional"
            Me.lbx_POSDATA.Clear
            Data = DataCM.getdata
        
        
        Case "Store Level"
            Data = DataCM.getdata
            For a = LBound(Data, 2) To UBound(Data, 2)
                If a = LBound(Data, 2) Then
                    ReDim ListData(1 To 6, 1 To 1)
                    ListLine = 1
                Else
                    If Data(1, a) <> ListData(1, ListLine) Then
                        ListLine = 0
                        For b = LBound(ListData, 2) To UBound(ListData, 2)
                            If ListData(1, b) = Data(1, a) & "-" & Data(2, a) Then
                                ListLine = b
                                Exit For
                            End If
                        Next
                        If ListLine = 0 Then
                            ReDim Preserve ListData(1 To 6, 1 To UBound(ListData, 2) + 1)
                            ListLine = UBound(ListData, 2)
                        End If
                    End If
                End If
                ListData(1, ListLine) = Data(1, a) & "-" & Data(2, a)
                If Data(3, a) >= Pstart Then
                    ListData(2, ListLine) = ListData(2, ListLine) + Data(4, a)
                    ListData(4, ListLine) = ListData(4, ListLine) + Data(5, a)
                Else
                    ListData(3, ListLine) = ListData(3, ListLine) + Data(4, a)
                    ListData(5, ListLine) = ListData(5, ListLine) + Data(5, a)
                End If
            Next
            If ListData(1, 1) <> "" Then
                Me.lbx_POSDATA.Clear
                Me.lbx_POSDATA.ColumnCount = 5
                Me.lbx_POSDATA.ColumnWidths = "80pt;100pt,55pt,100pt,55pt"
                Me.lbx_POSDATA.ColumnHeads = False
                
                'Me.lbx_POSDATA.Column(0).Name = "Name"
                Me.lbx_POSDATA.AddItem
                Me.lbx_POSDATA.Column(0, 0) = "StoreNo"
                Me.lbx_POSDATA.Column(1, 0) = "Retail"
                Me.lbx_POSDATA.Column(2, 0) = "%Dif"
                Me.lbx_POSDATA.Column(3, 0) = "Quantity"
                Me.lbx_POSDATA.Column(4, 0) = "%Dif"
                'Me.lbx_POSDATA.ListIndex = 0
                
                
                
                For a = LBound(ListData, 2) To UBound(ListData, 2)
                    If ListData(5, a) > 0 Then ListData(5, a) = ((ListData(4, a) - ListData(5, a)) / ListData(5, a))
                    If ListData(3, a) > 0 Then ListData(3, a) = ((ListData(2, a) - ListData(3, a)) / ListData(3, a))
                    Me.lbx_POSDATA.AddItem
                    For b = LBound(ListData, 1) To UBound(ListData, 1)
                        If b = 2 Or b = 4 Then
                            Me.lbx_POSDATA.Column(b - 1, a) = Format(IIf(IsEmpty(ListData(b, a)), 0, ListData(b, a)), "$#,0.00")
                        ElseIf b = 3 Or b = 5 Then
                            Me.lbx_POSDATA.Column(b - 1, a) = Format(IIf(IsEmpty(ListData(b, a)), 0, ListData(b, a)), "#,0.0%")
                        Else
                            Me.lbx_POSDATA.Column(b - 1, a) = IIf(IsEmpty(ListData(b, a)), 0, ListData(b, a))
                        End If
                    
                    Next
                    
                Next
            End If
           
        
        
    End Select

End Sub
