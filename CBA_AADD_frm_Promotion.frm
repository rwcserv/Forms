VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CBA_AADD_frm_Promotion 
   ClientHeight    =   14775
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   23760
   OleObjectBlob   =   "CBA_AADD_frm_Promotion.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "CBA_AADD_frm_Promotion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Data() As CBA_AADD_Product
Private StartDate As Date
Private EndDate As Date
Private PromoState As Boolean

'#RW Added new mousewheel routines 190701
Private Sub lbx_Prods_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call mw_SetBoxHook(lbx_Prods)
End Sub

Private Function checkforActiveMedium() As Boolean
    
    If Me.tbx_StartDate <> "" And Me.cbx_Medium <> "" And Me.cbx_Wkdur <> "" Then
        checkforActiveMedium = True
        Select Case cbx_Medium.Value
            Case "Television"
                setFrameToVisible Me.frame_TV
            Case "Radio"
                setFrameToVisible Me.frame_Radio
            Case "Press"
                setFrameToVisible Me.frame_Press
            Case "Digital"
                setFrameToVisible Me.frame_digital
            Case "Catalogue"
                setFrameToVisible Me.frame_Catalogue
            Case "Standee"
                setFrameToVisible Me.frame_Standee
            Case "POS"
                setFrameToVisible Me.frame_POS
        End Select
    Else
        checkforActiveMedium = False
        setFrameToVisible
    End If


End Function
Private Sub but_RemovePromo_Click()
    Dim a As Integer
    Dim b As Long
    Dim UF As Object
    For a = 0 To lbx_Prods.ListCount - 1
        If lbx_Prods.Selected(a) = True Then
            For b = LBound(Data) To UBound(Data)
                If CStr(Data(b).Product) = Split(Me.lbx_Prods.List(a), "-")(0) Then
                   Set Data(b) = Nothing
                    For Each UF In VBA.UserForms
                        If UF.Name = "CBA_AADD_frm_Product" Then
                            If UF.lbl_productcode = Split(Me.lbx_Prods.List(a), "-")(0) Then
                                Unload UF
                                Exit For
                            End If
                        End If
                    Next
                   Me.lbx_Prods.RemoveItem a
                   Exit For
                End If
            Next
        End If
    Next


End Sub

Private Sub cbx_Promo_Change()
    If cbx_Promo.Value = "" Then
        isPromoActive False
        Me.cbx_Medium = ""
        Me.cbx_Wkdur = ""
        Me.tbx_StartDate = ""
    Else
        isPromoActive True
    End If
End Sub
Function isPromoActive(Optional ByVal setToActive As Boolean) As Boolean
    If IsMissing(setToActive) = True Then isPromoActive = PromoState: Exit Function
    
    If setToActive = True Then
        Me.cbx_Medium.Visible = True
        Me.cbx_Wkdur.Visible = True
        Me.tbx_StartDate.Visible = True
        Me.lbx_Prods.Visible = True
        Me.Label2.Visible = True
        Me.Label4.Visible = True
        Me.Label6.Visible = True
        Me.Label7.Visible = True
        Me.Label8.Visible = True
        Me.Label9.Visible = True
        Me.Label10.Visible = True
        Me.Label11.Visible = True
        Me.Label12.Visible = True
        Me.Label13.Visible = True
        Me.Label14.Visible = True
        Me.Label15.Visible = True
        Me.tbx_Prod.Visible = True
        Me.but_RemovePromo = True
        checkforActiveMedium
        PromoState = True
    Else
        Me.cbx_Medium.Visible = False
        Me.cbx_Wkdur.Visible = False
        Me.tbx_StartDate.Visible = False
        Me.lbx_Prods.Visible = False
        Me.Label2.Visible = False
        Me.Label4.Visible = False
        Me.Label6.Visible = False
        Me.Label7.Visible = False
        Me.Label8.Visible = False
        Me.Label9.Visible = False
        Me.Label10.Visible = False
        Me.Label11.Visible = False
        Me.Label12.Visible = False
        Me.Label13.Visible = False
        Me.Label14.Visible = False
        Me.tbx_Prod.Visible = False
        Me.but_RemovePromo = False
        checkforActiveMedium
        PromoState = False
    End If

End Function

Sub UserForm_Initialize()
    Dim a As Long
    Dim this_Frame As Control
    Me.StartUpPosition = 0
    Me.Top = Application.Top + (Application.Height / 2) - (Me.Height / 2)
    Me.Left = Application.Left + (Application.Width / 30)
    For a = 1 To 52
        cbx_Wkdur.AddItem a
    Next
    cbx_Medium.AddItem "Television"
    cbx_Medium.AddItem "Radio"
    cbx_Medium.AddItem "Press"
    cbx_Medium.AddItem "Digital"
    cbx_Medium.AddItem "Catalogue"
    cbx_Medium.AddItem "Standee"
    cbx_Medium.AddItem "POS"
    For Each this_Frame In Me.Controls
        If TypeOf this_Frame Is Frame Then this_Frame.Visible = False
    Next
    ReDim Data(1 To 1)
    isPromoActive False
End Sub
Private Sub but_NewPromo_Click()
Dim newPromoName As String
    newPromoName = InputBox("Please enter the name of the new promotion", "New Promotion")
    If newPromoName <> "" Then
        Me.cbx_Promo.AddItem newPromoName
        Me.cbx_Promo.Value = newPromoName
    Else

         'MsgBox "No valid entry"
    End If
End Sub
Private Sub cbx_Medium_Change()
    checkforActiveMedium
End Sub
Private Sub cbx_Wkdur_Change()
    If cbx_Wkdur <> "" Then
        If DateAdd("WW", cbx_Wkdur, StartDate) > Date Then
            EndDate = Date
        Else
            EndDate = DateAdd("WW", cbx_Wkdur, StartDate)
        End If
    End If
    checkforActiveMedium
End Sub
Private Sub lbx_Prods_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim a As Integer
Dim b As Long
Dim frm As CBA_AADD_frm_Product
    For a = 0 To lbx_Prods.ListCount - 1
        If lbx_Prods.Selected(a) = True Then
            For b = LBound(Data) To UBound(Data)
                If CStr(Data(b).Product) = Mid(lbx_Prods.List(a), 1, InStr(1, lbx_Prods.List(a), "-") - 1) Then
                   Set frm = New CBA_AADD_frm_Product
                   Set frm.CBA_AADD_frm_ProductSetup = Data(b)
                   frm.Show vbModeless
                End If
            Next
        End If
    Next

End Sub

Private Sub tbx_Prod_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
Dim bOutput As Boolean
    If KeyCode = 13 Then
        If IsNumeric(tbx_Prod.Value) And tbx_Prod.Value <> "" And Len(tbx_Prod.Value) > 3 And Len(tbx_Prod.Value) < 6 Then
            bOutput = CCM_SQLQueries.CBA_COM_MATCHGenPullSQL("FINDPRODDESC", , , , , tbx_Prod.Value)
            If bOutput = True Then
                If Data(1) Is Nothing Then Else ReDim Preserve Data(1 To UBound(Data) + 1)
                Set Data(UBound(Data)) = New CBA_AADD_Product
                Me.lbx_Prods.AddItem tbx_Prod.Value & "-" & CBA_CBISarr(0, 0)
                Data(UBound(Data)).BuildIt tbx_Prod.Value, Mid(Year(StartDate), 3, 2), Month(StartDate), Day(StartDate), DateDiff("D", StartDate, EndDate), CBA_CBISarr(0, 0)
                tbx_Prod.Value = ""
            Else
                MsgBox "Invalid CBIS Product Code Entered", vbOKOnly
                tbx_Prod.Value = ""
            End If
        End If
    End If
End Sub
Private Sub tbx_StartDate_AfterUpdate()
    If isDate(tbx_StartDate.Value) And tbx_StartDate.Value <> "12:00:00 AM" Then
        StartDate = tbx_StartDate.Value
    Else
        MsgBox "Not a valid date"
        StartDate = "12:00:00 AM"
    End If
    checkforActiveMedium
End Sub


Private Sub UserForm_Terminate()
    Call but_Stop_Click
End Sub
Private Sub but_Stop_Click()
    Unload Me
End Sub
Private Sub setFrameToVisible(Optional ByRef med_frame As Object)

Dim this_Frame As Control
    For Each this_Frame In Me.Controls
            If TypeOf this_Frame Is Frame Then
                If med_frame Is Nothing Then
                    this_Frame.Visible = False
                Else
                   If this_Frame.Name = med_frame.Name Then this_Frame.Visible = True Else this_Frame.Visible = False
                End If
            End If
    Next
End Sub


