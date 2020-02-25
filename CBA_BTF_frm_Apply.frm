VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CBA_BTF_frm_Apply 
   Caption         =   "Forecast Entry"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   17520
   OleObjectBlob   =   "CBA_BTF_frm_Apply.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CBA_BTF_frm_Apply"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit       ' @CBA_BTF
 
Private Type FCData
    Year As Long
    Month As Long
    CG As Long
    scg As Long
    CR_PYRetail As Single
    CR_PYMarginP As Single
    CR_CYRetail As Single
    CR_OrigForcast As Single
    CR_Reforcast As Single
    CR_MarginOrigForcast As Single
    CR_MarginReforcast As Single
    FS_PYRetail As Single
    FS_PYMarginP As Single
    FS_CYRetail As Single
    FS_OrigForcast As Single
    FS_Reforcast As Single
    FS_MarginOrigForcast As Single
    FS_MarginReforcast As Single
    NFS_PYRetail As Single
    NFS_PYMarginP As Single
    NFS_CYRetail As Single
    NFS_OrigForcast As Single
    NFS_Reforcast As Single
    NFS_MarginOrigForcast As Single
    NFS_MarginReforcast As Single
    SEA_PYRetail As Single
    SEA_PYMarginP As Single
    SEA_CYRetail As Single
    SEA_OrigForcast As Single
    SEA_Reforcast As Single
    SEA_MarginOrigForcast As Single
    SEA_MarginReforcast As Single
End Type
Private FC As FCData
Private CGList As Variant
Private OriginalForecastCutoff As Date
Private DataSubmitted As Boolean

Private Sub but_NextMonth_Click()
    If cbx_Month = "" Or cbx_Year = "" Or cbx_CG = "" Or cbx_SCG = "" Then
        MsgBox "No current forecast period active"
        Exit Sub
    End If
    If cbx_Month = 12 Then
        cbx_Year = cbx_Year + 1
        cbx_Month = 1
    Else
        cbx_Month = cbx_Month + 1
    End If
End Sub

Private Sub but_PrevMonth_Click()
    If cbx_Month = "" Or cbx_Year = "" Or cbx_CG = "" Or cbx_SCG = "" Then
        MsgBox "No current forecast period active"
        Exit Sub
    End If
    If cbx_Month = 1 Then
        cbx_Year = cbx_Year - 1
        cbx_Month = 12
    Else
        cbx_Month = cbx_Month - 1
    End If
End Sub

Private Sub CR_Uplift_AfterUpdate()
    Uplifting CR_Uplift
End Sub

Private Sub FS_Uplift_AfterUpdate()
    Uplifting FS_Uplift
End Sub

Private Sub NFS_Uplift_AfterUpdate()
    Uplifting NFS_Uplift
End Sub
Private Sub SEA_Uplift_AfterUpdate()
    Uplifting SEA_Uplift
End Sub
Private Function Uplifting(ByRef tbx As Object)
    If InStr(1, tbx.Value, "%") > 0 Then
        If InStr(1, tbx.Value, "%") = 1 Then
            If IsNumeric(Mid(tbx.Value, 2, Len(tbx.Value) - 1)) = True Then tbx.Value = Mid(tbx.Value, 2, Len(tbx.Value) - 1) / 100
        ElseIf InStr(1, tbx.Value, "%") = Len(tbx.Value) Then
            If IsNumeric(Mid(tbx.Value, 1, InStr(1, tbx.Value, "%") - 1)) Then tbx.Value = Mid(tbx.Value, 1, InStr(1, tbx.Value, "%") - 1) / 100
        Else
            tbx.Value = ""
            MsgBox "You have made an invalid entry. Please try again."
            checkenteredvalues
            Exit Function
        End If
    End If
    If IsNumeric(tbx.Value) = False Then
        If tbx.Value <> "" Then
            tbx.Value = ""
            MsgBox "You have made an invalid entry. Please try again."
        End If
    Else
        If tbx.Value > 0 Then tbx.Value = tbx.Value / 100
        If Date <= OriginalForecastCutoff Then
            If tbx = CR_Uplift Then FC.CR_OrigForcast = FC.CR_PYRetail + (FC.CR_PYRetail * tbx.Value): FC.CR_MarginOrigForcast = FC.CR_PYMarginP: Me.CR_Sales.Value = Format(FC.CR_OrigForcast, "$#,0.00"): Me.CR_MarginP = Format(FC.CR_MarginOrigForcast, "#,0.00%")
            If tbx = FS_Uplift Then FC.FS_OrigForcast = FC.FS_PYRetail + (FC.FS_PYRetail * tbx.Value): FC.FS_MarginOrigForcast = FC.FS_PYMarginP: Me.FS_Sales.Value = Format(FC.FS_OrigForcast, "$#,0.00"): Me.FS_MarginP = Format(FC.FS_MarginOrigForcast, "#,0.00%")
            If tbx = NFS_Uplift Then FC.NFS_OrigForcast = FC.NFS_PYRetail + (FC.NFS_PYRetail * tbx.Value): FC.NFS_MarginOrigForcast = FC.NFS_PYMarginP: Me.NFS_Sales.Value = Format(FC.NFS_OrigForcast, "$#,0.00"): Me.NFS_MarginP = Format(FC.NFS_MarginOrigForcast, "#,0.00%")
            If tbx = SEA_Uplift Then FC.SEA_OrigForcast = FC.SEA_PYRetail + (FC.SEA_PYRetail * tbx.Value): FC.SEA_MarginOrigForcast = FC.SEA_PYMarginP: Me.SEA_Sales.Value = Format(FC.SEA_OrigForcast, "$#,0.00"): Me.SEA_MarginP = Format(FC.SEA_MarginOrigForcast, "#,0.00%")
        Else
            If tbx = CR_Uplift Then FC.CR_Reforcast = FC.CR_PYRetail + (FC.CR_PYRetail * tbx.Value): FC.CR_MarginReforcast = FC.CR_PYMarginP: Me.CR_Sales.Value = Format(FC.CR_Reforcast, "$#,0.00"): Me.CR_MarginP = Format(FC.CR_MarginReforcast, "#,0.00%")
            If tbx = FS_Uplift Then FC.FS_Reforcast = FC.FS_PYRetail + (FC.FS_PYRetail * tbx.Value): FC.FS_MarginReforcast = FC.FS_PYMarginP: Me.FS_Sales.Value = Format(FC.FS_Reforcast, "$#,0.00"): Me.FS_MarginP = Format(FC.FS_MarginReforcast, "#,0.00%")
            If tbx = NFS_Uplift Then FC.NFS_Reforcast = FC.NFS_PYRetail + (FC.NFS_PYRetail * tbx.Value): FC.NFS_MarginReforcast = FC.NFS_PYMarginP: Me.NFS_Sales.Value = Format(FC.NFS_Reforcast, "$#,0.00"): Me.NFS_MarginP = Format(FC.NFS_MarginReforcast, "#,0.00%")
            If tbx = SEA_Uplift Then FC.SEA_Reforcast = FC.SEA_PYRetail + (FC.SEA_PYRetail * tbx.Value): FC.SEA_MarginReforcast = FC.SEA_PYMarginP: Me.SEA_Sales.Value = Format(FC.SEA_Reforcast, "$#,0.00"): Me.SEA_MarginP = Format(FC.SEA_MarginReforcast, "#,0.00%")
        End If
        tbx.Value = ""
        checkenteredvalues
        DataSubmitted = False
    End If
End Function
Private Sub UserForm_Terminate()
    ClearFC
End Sub
Private Sub UserForm_Initialize()
Dim lY As Long, a As Long, b As Long
Dim bOutput As Boolean, bfound As Boolean

    Me.StartUpPosition = 0
    Me.Top = Application.Top + (Application.Height / 2) - (Me.Height / 2)
    Me.Left = Application.Left + (Application.Width / 2) - (Me.Width / 2)
    For a = 1 To 12
        cbx_Month.AddItem a
    Next a
    lY = Year(Date) - 1
    For a = lY To lY + 6
        cbx_Year.AddItem a
    Next a
    bOutput = CBA_SQL_Queries.CBA_GenPullSQL("CBIS_CGSCGList")
    If bOutput = True Then
        CGList = CBA_CBISarr
        Call g_EraseAry(CBA_CBISarr)
        For a = LBound(CGList, 2) To UBound(CGList, 2)
            bfound = False
            For b = 0 To cbx_CG.ListCount - 1
                If cbx_CG.List(b) = CGList(0, a) Then
                    bfound = True
                    Exit For
                End If
            Next
            If bfound = False Then cbx_CG.AddItem CGList(0, a)
        Next
    Else
        cbx_CG.Enabled = False
    End If
    cbx_SCG.Enabled = False
    DisableForecast
End Sub
Function DisableForecast()
    fam_CR.Enabled = False: Me.fam_FS.Enabled = False: fam_NFS.Enabled = False: fam_SEA.Enabled = False
End Function
Function UpdateForm()
    Dim bOutput As Boolean, a As Long
    Dim FCSQL As Variant

    fam_CR.Enabled = True: Me.fam_FS.Enabled = True: fam_NFS.Enabled = True: fam_SEA.Enabled = True
    ClearFormValues
    ClearFC True
    bOutput = CBA_SQL_Queries.CBA_GenPullSQL("CBIS_ForecastQuery", , , FC.CG, FC.scg, FC.Month, FC.Year)
    If bOutput = True Then
        FCSQL = CBA_CBISarr
        Erase CBA_CBISarr
    Else
        Exit Function
    End If
    For a = LBound(FCSQL, 2) To UBound(FCSQL, 2)
        If FCSQL(0, a) = 1 Then
            CR_PYRetail.Value = Format(FCSQL(1, a), "$#,0.00"): CR_CYRetail.Value = Format(FCSQL(2, a), "$#,0.00"): CR_PYMarginP.Value = Format(FCSQL(3, a), "#,0.00%")
            FC.CR_PYRetail = FCSQL(1, a): FC.CR_CYRetail = FCSQL(2, a): FC.CR_PYMarginP = FCSQL(3, a)
        ElseIf FCSQL(0, a) = 2 Then
            FS_PYRetail.Value = Format(FCSQL(1, a), "$#,0.00"): FS_CYRetail.Value = Format(FCSQL(2, a), "$#,0.00"): FS_PYMarginP.Value = Format(FCSQL(3, a), "#,0.00%")
            FC.FS_PYRetail = FCSQL(1, a): FC.FS_CYRetail = FCSQL(2, a): FC.FS_PYMarginP = FCSQL(3, a)
        ElseIf FCSQL(0, a) = 3 Then
            NFS_PYRetail.Value = Format(FCSQL(1, a), "$#,0.00"): NFS_CYRetail.Value = Format(FCSQL(2, a), "$#,0.00"): NFS_PYMarginP.Value = Format(FCSQL(3, a), "#,0.00%")
            FC.NFS_PYRetail = FCSQL(1, a): FC.NFS_CYRetail = FCSQL(2, a): FC.NFS_PYMarginP = FCSQL(3, a)
        ElseIf FCSQL(0, a) = 4 Then
            SEA_PYRetail.Value = Format(FCSQL(1, a), "$#,0.00"): SEA_CYRetail.Value = Format(FCSQL(2, a), "$#,0.00"): SEA_PYMarginP.Value = Format(FCSQL(3, a), "#,0.00%")
            FC.SEA_PYRetail = FCSQL(1, a): FC.SEA_CYRetail = FCSQL(2, a): FC.SEA_PYMarginP = FCSQL(3, a)
        End If
    Next
    bOutput = CBA_SQL_Queries.CBA_GenPullSQL("CBForecast_SalesReforecastQuery", , , 1, FC.Month, FC.Year, FC.CG, , , , , , , FC.scg)
    If bOutput = True Then If CBA_CBFCarr(0, 0) <> 0 Then Me.CR_Sales = Format(CBA_CBFCarr(0, 0), "$#,0.00")
    If Me.CR_Sales = "" Then
        bOutput = CBA_SQL_Queries.CBA_GenPullSQL("CBForecast_SalesOrigQuery", , , 1, FC.Month, FC.Year, FC.CG, , , , , , , FC.scg)
        If bOutput = True Then If CBA_CBFCarr(0, 0) <> 0 Then Me.CR_Sales = Format(CBA_CBFCarr(0, 0), "$#,0.00")
    End If
    bOutput = CBA_SQL_Queries.CBA_GenPullSQL("CBForecast_MarginReforecastQuery", , , 1, FC.Month, FC.Year, FC.CG, , , , , , , FC.scg)
    If bOutput = True Then If CBA_CBFCarr(0, 0) <> 0 Then Me.CR_MarginP = Format(CBA_CBFCarr(0, 0), "#,0.00%")
    If Me.CR_MarginP = "" Then
        bOutput = CBA_SQL_Queries.CBA_GenPullSQL("CBForecast_MarginOrigQuery", , , 1, FC.Month, FC.Year, FC.CG, , , , , , , FC.scg)
        If bOutput = True Then If CBA_CBFCarr(0, 0) <> 0 Then Me.CR_MarginP = Format(CBA_CBFCarr(0, 0), "#,0.00%")
    End If
    bOutput = CBA_SQL_Queries.CBA_GenPullSQL("CBForecast_SalesReforecastQuery", , , 2, FC.Month, FC.Year, FC.CG, , , , , , , FC.scg)
    If bOutput = True Then If CBA_CBFCarr(0, 0) <> 0 Then Me.FS_Sales = Format(CBA_CBFCarr(0, 0), "$#,0.00")
    If Me.FS_Sales = "" Then
        bOutput = CBA_SQL_Queries.CBA_GenPullSQL("CBForecast_SalesOrigQuery", , , 2, FC.Month, FC.Year, FC.CG, , , , , , , FC.scg)
        If bOutput = True Then If CBA_CBFCarr(0, 0) <> 0 Then Me.FS_Sales = Format(CBA_CBFCarr(0, 0), "$#,0.00")
    End If
    bOutput = CBA_SQL_Queries.CBA_GenPullSQL("CBForecast_MarginReforecastQuery", , , 2, FC.Month, FC.Year, FC.CG, , , , , , , FC.scg)
    If bOutput = True Then If CBA_CBFCarr(0, 0) <> 0 Then Me.FS_MarginP = Format(CBA_CBFCarr(0, 0), "#,0.00%")
    If Me.FS_MarginP = "" Then
        bOutput = CBA_SQL_Queries.CBA_GenPullSQL("CBForecast_MarginOrigQuery", , , 2, FC.Month, FC.Year, FC.CG, , , , , , , FC.scg)
        If bOutput = True Then If CBA_CBFCarr(0, 0) <> 0 Then Me.FS_MarginP = Format(CBA_CBFCarr(0, 0), "#,0.00%")
    End If
    bOutput = CBA_SQL_Queries.CBA_GenPullSQL("CBForecast_SalesReforecastQuery", , , 3, FC.Month, FC.Year, FC.CG, , , , , , , FC.scg)
    If bOutput = True Then If CBA_CBFCarr(0, 0) <> 0 Then Me.NFS_Sales = Format(CBA_CBFCarr(0, 0), "$#,0.00")
    If Me.NFS_Sales = "" Then
        bOutput = CBA_SQL_Queries.CBA_GenPullSQL("CBForecast_SalesOrigQuery", , , 3, FC.Month, FC.Year, FC.CG, , , , , , , FC.scg)
        If bOutput = True Then If CBA_CBFCarr(0, 0) <> 0 Then Me.NFS_Sales = Format(CBA_CBFCarr(0, 0), "$#,0.00")
    End If
    bOutput = CBA_SQL_Queries.CBA_GenPullSQL("CBForecast_MarginReforecastQuery", , , 3, FC.Month, FC.Year, FC.CG, , , , , , , FC.scg)
    If bOutput = True Then If CBA_CBFCarr(0, 0) <> 0 Then Me.NFS_MarginP = Format(CBA_CBFCarr(0, 0), "#,0.00%")
    If Me.NFS_MarginP = "" Then
        bOutput = CBA_SQL_Queries.CBA_GenPullSQL("CBForecast_MarginOrigQuery", , , 3, FC.Month, FC.Year, FC.CG, , , , , , , FC.scg)
        If bOutput = True Then If CBA_CBFCarr(0, 0) <> 0 Then Me.NFS_MarginP = Format(CBA_CBFCarr(0, 0), "#,0.00%")
    End If
    bOutput = CBA_SQL_Queries.CBA_GenPullSQL("CBForecast_SalesReforecastQuery", , , 4, FC.Month, FC.Year, FC.CG, , , , , , , FC.scg)
    If bOutput = True Then If CBA_CBFCarr(0, 0) <> 0 Then Me.SEA_Sales = Format(CBA_CBFCarr(0, 0), "$#,0.00")
    If Me.SEA_Sales = "" Then
        bOutput = CBA_SQL_Queries.CBA_GenPullSQL("CBForecast_SalesOrigQuery", , , 4, FC.Month, FC.Year, FC.CG, , , , , , , FC.scg)
        If bOutput = True Then If CBA_CBFCarr(0, 0) <> 0 Then Me.SEA_Sales = Format(CBA_CBFCarr(0, 0), "$#,0.00")
    End If
    bOutput = CBA_SQL_Queries.CBA_GenPullSQL("CBForecast_MarginReforecastQuery", , , 4, FC.Month, FC.Year, FC.CG, , , , , , , FC.scg)
    If bOutput = True Then If CBA_CBFCarr(0, 0) <> 0 Then Me.SEA_MarginP = Format(CBA_CBFCarr(0, 0), "#,0.00%")
    If Me.SEA_MarginP = "" Then
        bOutput = CBA_SQL_Queries.CBA_GenPullSQL("CBForecast_MarginOrigQuery", , , 4, FC.Month, FC.Year, FC.CG, , , , , , , FC.scg)
        If bOutput = True Then If CBA_CBFCarr(0, 0) <> 0 Then Me.SEA_MarginP = Format(CBA_CBFCarr(0, 0), "#,0.00%")
    End If
    DataSubmitted = True
    checkenteredvalues
    If Year(Date) > FC.Year Or (Year(Date) = FC.Year And Month(Date) >= FC.Month) Then
        Me.CR_Sales.Locked = True: Me.CR_MarginP.Locked = True: Me.CR_Uplift.Locked = True
        Me.FS_Sales.Locked = True: Me.FS_MarginP.Locked = True: Me.FS_Uplift.Locked = True
        Me.NFS_Sales.Locked = True: Me.NFS_MarginP.Locked = True: Me.NFS_Uplift.Locked = True
        Me.SEA_Sales.Locked = True: Me.SEA_MarginP.Locked = True: Me.SEA_Uplift.Locked = True
        MsgBox "You have selected a current or previous period, therefore the form is not editable"
    Else
        Me.CR_Sales.Locked = False: Me.CR_MarginP.Locked = False: Me.CR_Uplift.Locked = False
        Me.FS_Sales.Locked = False: Me.FS_MarginP.Locked = False: Me.FS_Uplift.Locked = False
        Me.NFS_Sales.Locked = False: Me.NFS_MarginP.Locked = False: Me.NFS_Uplift.Locked = False
        Me.SEA_Sales.Locked = False: Me.SEA_MarginP.Locked = False: Me.SEA_Uplift.Locked = False
    End If
End Function
Private Sub btn_Submit_Click()
    SubmitData
End Sub


Private Function SubmitData()
'submits forcast data to Access

If FC.CR_MarginOrigForcast <> 0 Or FC.CR_MarginReforcast <> 0 Or FC.CR_OrigForcast <> 0 Or FC.CR_Reforcast <> 0 Then
    CBA_SQL_Queries.CBA_GenPullSQL "CBForecast_Apply", Date, , 1, FC.Month, FC.Year, FC.CG, , , , , , , FC.scg, , FC.CR_OrigForcast, FC.CR_Reforcast, FC.CR_MarginOrigForcast, FC.CR_MarginReforcast
End If
If FC.FS_MarginOrigForcast <> 0 Or FC.FS_MarginReforcast <> 0 Or FC.FS_OrigForcast <> 0 Or FC.FS_Reforcast <> 0 Then
    CBA_SQL_Queries.CBA_GenPullSQL "CBForecast_Apply", Date, , 2, FC.Month, FC.Year, FC.CG, , , , , , , FC.scg, , FC.FS_OrigForcast, FC.FS_Reforcast, FC.FS_MarginOrigForcast, FC.FS_MarginReforcast
End If
If FC.NFS_MarginOrigForcast <> 0 Or FC.NFS_MarginReforcast <> 0 Or FC.NFS_OrigForcast <> 0 Or FC.NFS_Reforcast <> 0 Then
    CBA_SQL_Queries.CBA_GenPullSQL "CBForecast_Apply", Date, , 3, FC.Month, FC.Year, FC.CG, , , , , , , FC.scg, , FC.NFS_OrigForcast, FC.NFS_Reforcast, FC.NFS_MarginOrigForcast, FC.NFS_MarginReforcast
End If
If FC.SEA_MarginOrigForcast <> 0 Or FC.SEA_MarginReforcast <> 0 Or FC.SEA_OrigForcast <> 0 Or FC.SEA_Reforcast <> 0 Then
    CBA_SQL_Queries.CBA_GenPullSQL "CBForecast_Apply", Date, , 4, FC.Month, FC.Year, FC.CG, , , , , , , FC.scg, , FC.SEA_OrigForcast, FC.SEA_Reforcast, FC.SEA_MarginOrigForcast, FC.SEA_MarginReforcast
End If
DataSubmitted = True

    

End Function
'Private Sub cbx_CG_Change()
'    bFound = False
'    For a = LBound(CGList, 2) To UBound(CGList, 2)
'        If cbx_CG.Value = CGList(0, a) Then
'            bFound = True
'            Exit For
'        End If
'    Next
'    If bFound = True Then
'        cbx_SCG.Enabled = True
'        cbx_SCG.Clear
'        For a = LBound(CGList, 2) To UBound(CGList, 2)
'            If cbx_CG.Value = CGList(0, a) Then cbx_SCG.AddItem CGList(1, a)
'        Next
'    End If
'End Sub
Private Sub cbx_CG_Change()
    Dim bfound As Boolean, lRet As Long, a As Long

    bfound = False
    For a = LBound(CGList, 2) To UBound(CGList, 2)
        If cbx_CG.Value = CGList(0, a) Then
            bfound = True
            Exit For
        End If
    Next
    If bfound = False Then
        cbx_CG.Value = ""
        Exit Sub
    Else
        cbx_SCG.Enabled = True
        cbx_SCG.Clear
        For a = LBound(CGList, 2) To UBound(CGList, 2)
            If cbx_CG.Value = CGList(0, a) Then cbx_SCG.AddItem CGList(1, a)
        Next
    End If
    If checktorun = True Then
        If DataSubmitted = False Then
            lRet = MsgBox("Any entered and unsubmitted entries will be lost. Do you want to submit the month data you are moving away from?", vbYesNo)
            If lRet = 6 Then SubmitData
        End If
    End If

    FC.CG = Trim(Mid(cbx_CG.Value, 1, InStr(1, cbx_CG.Value, " - ")))
    If checktorun = True Then UpdateForm Else DisableForecast
    'DataSubmitted = False
End Sub
Private Sub cbx_Month_Change()
    Dim lRet As Long
    If IsNumeric(cbx_Month.Value) = False Or cbx_Month.Value = "" Then
        cbx_Month.Value = ""
        Exit Sub
    End If
    If checktorun = True Then
        If DataSubmitted = False Then
            lRet = MsgBox("Any entered and unsubmitted entries will be lost. Do you want to submit the month data you are moving away from?", vbYesNo)
            If lRet = 6 Then SubmitData
        End If
    End If
    FC.Month = cbx_Month.Value
    If checktorun = True Then UpdateForm Else DisableForecast
    'DataSubmitted = False
End Sub
Private Sub cbx_SCG_Change()
Dim bfound As Boolean, lRet As Long, a As Long
    bfound = False
    For a = LBound(CGList, 2) To UBound(CGList, 2)
        If cbx_SCG.Value = CGList(1, a) Then
        bfound = True
        Exit For
        End If
    Next
    If bfound = False Then
        cbx_SCG.Value = ""
        Exit Sub
    End If
    If checktorun = True Then
        If DataSubmitted = False Then
        lRet = MsgBox("Any entered and unsubmitted entries will be lost. Do you want to submit the month data you are moving away from?", vbYesNo)
        If lRet = 6 Then SubmitData
        End If
    End If
    FC.scg = Trim(Mid(cbx_SCG.Value, 1, InStr(1, cbx_SCG.Value, " - ")))
    If checktorun = True Then UpdateForm Else DisableForecast
    'DataSubmitted = False
End Sub
Private Sub cbx_Year_Change()
    Dim lRet As Long
    If IsNumeric(cbx_Year.Value) = False Then
        cbx_Year.Value = ""
        Exit Sub
    End If
    If checktorun = True Then
        If DataSubmitted = False Then
        lRet = MsgBox("Any entered and unsubmitted entries will be lost. Do you want to submit the month data you are moving away from?", vbYesNo)
        If lRet = 6 Then SubmitData
        End If
    End If
    FC.Year = cbx_Year.Value
    If checktorun = True Then UpdateForm Else DisableForecast
    'DataSubmitted = False
End Sub
Private Function checktorun() As Boolean
Dim arr As Variant
    If FC.Month > 0 And FC.Year > 0 And FC.CG > 0 And FC.scg > 0 Then
        arr = CBA_SQL_Queries.CBA_GenPullSQL("CBForecast_CutOff", , , FC.Year, FC.Month)
        OriginalForecastCutoff = DateSerial(CBA_CBFCarr(2, 0), CBA_CBFCarr(1, 0), CBA_CBFCarr(0, 0))
        checktorun = True
    End If
End Function
Private Function ClearFC(Optional keepSelection As Boolean)
'Clears FC
    If keepSelection = False Then FC.Year = 0: FC.Month = 0: FC.CG = 0: FC.scg = 0
    FC.CR_PYRetail = 0: FC.CR_CYRetail = 0: FC.CR_OrigForcast = 0: FC.CR_Reforcast = 0
    FC.FS_PYRetail = 0: FC.FS_CYRetail = 0: FC.FS_OrigForcast = 0: FC.FS_Reforcast = 0
    FC.NFS_PYRetail = 0: FC.NFS_CYRetail = 0: FC.NFS_OrigForcast = 0: FC.NFS_Reforcast = 0
    FC.SEA_PYRetail = 0: FC.SEA_CYRetail = 0: FC.SEA_OrigForcast = 0: FC.SEA_Reforcast = 0
    
    
    
    
End Function

Function checkenteredvalues()
    Dim LYel, LWhi
        
    LYel = RGB(255, 255, 168)
    LWhi = RGB(255, 255, 255)
    If Me.CR_Sales.Value = "" Then Me.CR_Sales.BackColor = LYel Else Me.CR_Sales.BackColor = LWhi
    If Me.CR_MarginP.Value = "" Then Me.CR_MarginP.BackColor = LYel Else Me.CR_MarginP.BackColor = LWhi
    If Me.FS_Sales.Value = "" Then Me.FS_Sales.BackColor = LYel Else Me.FS_Sales.BackColor = LWhi
    If Me.FS_MarginP.Value = "" Then Me.FS_MarginP.BackColor = LYel Else Me.FS_MarginP.BackColor = LWhi
    If Me.NFS_Sales.Value = "" Then Me.NFS_Sales.BackColor = LYel Else Me.NFS_Sales.BackColor = LWhi
    If Me.NFS_MarginP.Value = "" Then Me.NFS_MarginP.BackColor = LYel Else Me.NFS_MarginP.BackColor = LWhi
    If Me.SEA_Sales.Value = "" Then Me.SEA_Sales.BackColor = LYel Else Me.SEA_Sales.BackColor = LWhi
    If Me.SEA_MarginP.Value = "" Then Me.SEA_MarginP.BackColor = LYel Else Me.SEA_MarginP.BackColor = LWhi

    If Me.CR_Sales.Value <> "" And Me.CR_MarginP.Value <> "" Then Me.CR_MarginD.Value = Format(Mid(Me.CR_Sales.Value, 2, 99) * (Mid(Me.CR_MarginP.Value, 1, Len(Me.CR_MarginP.Value) - 1) / 100), "$#,0.00") Else Me.CR_MarginD.Value = ""
    If Me.FS_Sales.Value <> "" And Me.FS_MarginP.Value <> "" Then Me.FS_MarginD.Value = Format(Mid(Me.FS_Sales.Value, 2, 99) * (Mid(Me.FS_MarginP.Value, 1, Len(Me.FS_MarginP.Value) - 1) / 100), "$#,0.00") Else Me.FS_MarginD.Value = ""
    If Me.NFS_Sales.Value <> "" And Me.NFS_MarginP.Value <> "" Then Me.NFS_MarginD.Value = Format(Mid(Me.NFS_Sales.Value, 2, 99) * (Mid(Me.NFS_MarginP.Value, 1, Len(Me.NFS_MarginP.Value) - 1) / 100), "$#,0.00") Else Me.NFS_MarginD.Value = ""
    If Me.SEA_Sales.Value <> "" And Me.SEA_MarginP.Value <> "" Then Me.SEA_MarginD.Value = Format(Mid(Me.SEA_Sales.Value, 2, 99) * (Mid(Me.SEA_MarginP.Value, 1, Len(Me.SEA_MarginP.Value) - 1) / 100), "$#,0.00") Else Me.SEA_MarginD.Value = ""
End Function
Function ClearFormValues()
    Me.CR_CYRetail = "": Me.CR_PYMarginP = "": Me.CR_MarginD = "": Me.CR_MarginP = "": Me.CR_PYRetail = "": Me.CR_Sales = ""
    Me.FS_CYRetail = "": Me.FS_PYMarginP = "": Me.FS_MarginD = "": Me.FS_MarginP = "": Me.FS_PYRetail = "": Me.FS_Sales = ""
    Me.NFS_CYRetail = "": Me.NFS_PYMarginP = "": Me.NFS_MarginD = "": Me.NFS_MarginP = "": Me.NFS_PYRetail = "": Me.NFS_Sales = ""
    Me.SEA_CYRetail = "": Me.SEA_PYMarginP = "": Me.SEA_MarginD = "": Me.SEA_MarginP = "": Me.SEA_PYRetail = "": Me.SEA_Sales = ""
End Function
Private Sub CR_Sales_AfterUpdate()
    EnterForcastValtoFC CR_Sales
End Sub
Private Sub FS_Sales_AfterUpdate()
    EnterForcastValtoFC FS_Sales
End Sub
Private Sub NFS_Sales_AfterUpdate()
    EnterForcastValtoFC NFS_Sales
End Sub
Private Sub SEA_Sales_AfterUpdate()
    EnterForcastValtoFC SEA_Sales
End Sub
Private Function EnterForcastValtoFC(ByRef tbx As Object)
    If IsNumeric(tbx.Value) = False Then
        If tbx.Value <> "" Then
            tbx.Value = ""
            MsgBox "You have made an invalid entry. Please try again."
        End If
    Else
        If Date <= OriginalForecastCutoff Then
            If tbx = CR_Sales Then FC.CR_OrigForcast = tbx.Value
            If tbx = FS_Sales Then FC.FS_OrigForcast = tbx.Value
            If tbx = NFS_Sales Then FC.NFS_OrigForcast = tbx.Value
            If tbx = SEA_Sales Then FC.SEA_OrigForcast = tbx.Value
        Else
            If tbx = CR_Sales Then FC.CR_Reforcast = tbx.Value
            If tbx = FS_Sales Then FC.FS_Reforcast = tbx.Value
            If tbx = NFS_Sales Then FC.NFS_Reforcast = tbx.Value
            If tbx = SEA_Sales Then FC.SEA_Reforcast = tbx.Value
        End If
        tbx.Value = Format(tbx.Value, "$#,0.00")
        checkenteredvalues
        DataSubmitted = False
    End If

End Function
Private Sub CR_MarginP_AfterUpdate()
    EnterMarginValtoFC CR_MarginP
End Sub
Private Sub FS_MarginP_AfterUpdate()
    EnterMarginValtoFC FS_MarginP
End Sub
Private Sub NFS_MarginP_AfterUpdate()
    EnterMarginValtoFC NFS_MarginP
End Sub
Private Sub SEA_MarginP_AfterUpdate()
    EnterMarginValtoFC SEA_MarginP
End Sub
Private Function EnterMarginValtoFC(ByRef tbx As Object)
    If InStr(1, tbx.Value, "%") > 0 Then
        If InStr(1, tbx.Value, "%") = 1 Then
            If IsNumeric(Mid(tbx.Value, 2, Len(tbx.Value) - 1)) = True Then tbx.Value = Mid(tbx.Value, 2, Len(tbx.Value) - 1) / 100
        ElseIf InStr(1, tbx.Value, "%") = Len(tbx.Value) Then
            If IsNumeric(Mid(tbx.Value, 1, InStr(1, tbx.Value, "%") - 1)) Then tbx.Value = Mid(tbx.Value, 1, InStr(1, tbx.Value, "%") - 1) / 100
        Else
            tbx.Value = ""
            MsgBox "You have made an invalid entry. Please try again."
            checkenteredvalues
            Exit Function
        End If
    End If
    If IsNumeric(tbx.Value) = False Then
        If tbx.Value <> "" Then
            tbx.Value = ""
            MsgBox "You have made an invalid entry. Please try again."
        End If
    Else
        If tbx.Value > 0 Then tbx.Value = tbx.Value / 100
        If Date <= OriginalForecastCutoff Then
            If tbx = CR_MarginP Then FC.CR_MarginOrigForcast = tbx.Value
            If tbx = FS_MarginP Then FC.FS_MarginOrigForcast = tbx.Value
            If tbx = NFS_MarginP Then FC.NFS_MarginOrigForcast = tbx.Value
            If tbx = SEA_MarginP Then FC.SEA_MarginOrigForcast = tbx.Value
        Else
            If tbx = CR_MarginP Then FC.CR_MarginReforcast = tbx.Value
            If tbx = FS_MarginP Then FC.FS_MarginReforcast = tbx.Value
            If tbx = NFS_MarginP Then FC.NFS_MarginReforcast = tbx.Value
            If tbx = SEA_MarginP Then FC.SEA_MarginReforcast = tbx.Value
        End If
        tbx.Value = Format(tbx.Value, "#,0.00%")
        checkenteredvalues
        DataSubmitted = False
    End If

End Function
