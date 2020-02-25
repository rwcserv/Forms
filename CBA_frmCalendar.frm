VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CBA_frmCalendar 
   Caption         =   "Calendar"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   3360
   OleObjectBlob   =   "CBA_frmCalendar.frx":0000
End
Attribute VB_Name = "CBA_frmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit    ' CBA_frmCalendar   181231 - Add ability to return null dates on input flag

Private Sub CB_Close_Click()
    On Error Resume Next
''    'Debug.Print "Top=" & Me.Top & ";Left=" & Me.Left
''    Me.UndoAction
    Unload CBA_frmCalendar
End Sub

Sub addDate()
    varCal.sDate = Format(ActiveControl.Tag, CBA_DMY)
    varCal.bCalValReturned = True
    Call CB_Close_Click
    ''ActiveCell.Value = Parent
End Sub

Private Sub cmdReturn_Click()
    varCal.sDate = ""
    varCal.bCalValReturned = True
    Call CB_Close_Click
End Sub

Private Sub sp1_Change()
    Me.CB_Yr = Me.sp1.Value
End Sub

Private Sub UserForm_Initialize()

    Dim i As Long, lYearsAdd As Long, lYearStart As Long, dtDate As Date
''    bSetupIP = True
    Me.Top = varCal.lCalTop ''+ 100                 ' Top doesn't seem to make any difference
    Me.Left = varCal.lCalLeft
''    Me.Top = 247.5
''    Me.Left = 1728.75
'   'Debug.Print "Top=" & Me.Top & "=" & varCal.lCalTop & ";   Left=" & Me.Left & "=" & varCal.lCalLeft
    varCal.bCalValReturned = False
    Me.cmdReturn.Visible = varCal.bAllowNullOfDate
    If g_IsDate(varCal.sDate, True) Then
        dtDate = Format(g_FixDate(varCal.sDate), CBA_DMY)
    Else
        dtDate = Date
    End If

    lYearStart = Year(dtDate) - 10
    lYearsAdd = Year(Date) + 10
    With Me
        For i = 1 To 12
            .CB_Mth.AddItem Format(DateSerial(Year(dtDate), i, 1), "mmmm")
        Next
        .Tag = "Calendar"
        ' Set the spinner
        .sp1.Min = lYearStart
        .sp1.Max = lYearsAdd
        .sp1.Value = Year(dtDate)
        .CB_Yr = .sp1.Value

        .CB_Mth.ListIndex = Month(dtDate) - 1
        .Tag = ""
    End With
    Call Build_Calendar

End Sub

Private Sub CB_Mth_Change()
    If Not Me.Tag = "Calendar" Then Build_Calendar
End Sub

Private Sub CB_Yr_Change()
    If Not Me.Tag = "Calendar" Then Build_Calendar
End Sub

Sub Build_Calendar()
    ' This routine will format the buttons
    Dim i      As Integer, dTemp  As Date, dTemp2 As Date, iFirstDay As Integer, dtSetDate As Date
    If g_IsDate(varCal.sDate) Then
        dtSetDate = varCal.sDate
    Else
        dtSetDate = Date
    End If
    With Me
        .Caption = " " & .CB_Mth.Value & " " & .CB_Yr.Value

        dTemp = CDate("01/" & .CB_Mth.Value & "/" & .CB_Yr.Value)
        iFirstDay = WeekDay(dTemp, vbSunday)
        .Controls("D" & iFirstDay).SetFocus

        For i = 1 To 42
            With .Controls("D" & i)
                dTemp2 = DateAdd("d", (i - iFirstDay), dTemp)
                .Caption = Format(dTemp2, "d")
                ' Put the date into the tag and controltip
                .Tag = dTemp2
                .ControlTipText = Format(dTemp2, "dd/mm/yy")
                ' Add days to the buttons
                If Format(dTemp2, "mmmm") = CB_Mth.Value Then
                    If .BackColor <> &H80000016 Then .BackColor = &H80000018
                    If Format(dTemp2, "m/d/yy") = Format(dtSetDate, "m/d/yy") Then .SetFocus
                    .Font.Bold = True
                ElseIf dTemp2 = Date Then
                    If .BackColor <> &H80000016 Then .BackColor = vbYellow
                Else
                    If .BackColor <> &H80000016 Then .BackColor = &H8000000F
                    .Font.Bold = False
                End If
                'format the buttons
            End With
        Next
    End With

End Sub

' ROUTINES FOR DOUBLE CLICKS START HERE
Private Sub D1_Click()
    Call addDate
End Sub

Private Sub D2_Click()
    Call addDate
End Sub

Private Sub D3_Click()
    Call addDate
End Sub

Private Sub D4_Click()
    Call addDate
End Sub

Private Sub D5_Click()
    Call addDate
End Sub

Private Sub D6_Click()
    Call addDate
End Sub

Private Sub D7_Click()
    Call addDate
End Sub

Private Sub D8_Click()
    Call addDate
End Sub

Private Sub D9_Click()
    Call addDate
End Sub

Private Sub D10_Click()
    Call addDate
End Sub

Private Sub D11_Click()
    Call addDate
End Sub

Private Sub D12_Click()
    Call addDate
End Sub

Private Sub D13_Click()
    Call addDate
End Sub

Private Sub D14_Click()
    Call addDate
End Sub

Private Sub D15_Click()
    Call addDate
End Sub

Private Sub D16_Click()
    Call addDate
End Sub

Private Sub D17_Click()
    Call addDate
End Sub

Private Sub D18_Click()
    Call addDate
End Sub

Private Sub D19_Click()
    Call addDate
End Sub

Private Sub D20_Click()
    Call addDate
End Sub

Private Sub D21_Click()
    Call addDate
End Sub

Private Sub D22_Click()
    Call addDate
End Sub

Private Sub D23_Click()
    Call addDate
End Sub

Private Sub D24_Click()
    Call addDate
End Sub

Private Sub D25_Click()
    Call addDate
End Sub

Private Sub D26_Click()
    Call addDate
End Sub

Private Sub D27_Click()
    Call addDate
End Sub

Private Sub D28_Click()
    Call addDate
End Sub

Private Sub D29_Click()
    Call addDate
End Sub

Private Sub D30_Click()
    Call addDate
End Sub

Private Sub D31_Click()
    Call addDate
End Sub

Private Sub D32_Click()
    Call addDate
End Sub

Private Sub D33_Click()
    Call addDate
End Sub

Private Sub D34_Click()
    Call addDate
End Sub

Private Sub D35_Click()
    Call addDate
End Sub

Private Sub D36_Click()
    Call addDate
End Sub

Private Sub D37_Click()
    Call addDate
End Sub

Private Sub D38_Click()
    Call addDate
End Sub

Private Sub D39_Click()
    Call addDate
End Sub

Private Sub D40_Click()
    Call addDate
End Sub

Private Sub D41_Click()
    Call addDate
End Sub

Private Sub D42_Click()
    Call addDate
End Sub

