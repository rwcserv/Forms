VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CBA_COM_frm_Chart 
   ClientHeight    =   9555
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14145
   OleObjectBlob   =   "CBA_COM_frm_Chart.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "CBA_COM_frm_Chart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub but_Stop_Click()
Unload Me
End Sub

Sub formulatePPHChart(ByVal CompCode As String, ByVal compDesc As String, ByVal sStateLook As String, ByVal compet As String)
Dim bOutput As Boolean, a As Long, b As Long, c
Dim MyChart As Chart
Dim imagename As String
Dim wks_ChartData As Worksheet
Dim wbk As Workbook
'If Mid(UCase(Comp2Find), 1, 5) = "COLES" Or InStr(1, assPCode, "P") > 0 Then
'    bOutput = mod_SQLQueries.GenPullSQL("Chart_Coles", , , , , CStr(assPCode), CStr(sStateLook))
'ElseIf Mid(UCase(Comp2Find), 1, 2) = "WW" Or IsNumeric(assPCode) Then
'    bOutput = mod_SQLQueries.GenPullSQL("Chart_WW", , , , , CStr(assPCode), CStr(sStateLook))
'ElseIf Mid(UCase(Comp2Find), 1, 2) = "DM" Or Mid(assPCode, 1, 3) = "DM_" Then
'    bOutput = mod_SQLQueries.GenPullSQL("Chart_DM", , , , CStr(assPack), CStr(assPCode), CStr(sStateLook))
'End If

Application.ScreenUpdating = False
    
Set wbk = Application.Workbooks.Add
Set wks_ChartData = ActiveSheet


If CBA_COM_frm_MatchingTool.getCCM_UserDefinedState = True Then
    If compet = "WW" Then
        bOutput = CBA_COM_SQLQueries.CBA_COM_GenPullSQL("Chart_WW", , , , , CompCode, CStr(sStateLook))
    ElseIf compet = "Coles" Then
        bOutput = CBA_COM_SQLQueries.CBA_COM_GenPullSQL("Chart_Coles", , , , , CompCode, CStr(sStateLook))
    End If
Else
    If compet = "WW" Then
        bOutput = CBA_COM_SQLQueries.CBA_COM_GenPullSQL("Chart_WW", , , , , CompCode, CStr(sStateLook))
    ElseIf compet = "Coles" Then
        bOutput = CBA_COM_SQLQueries.CBA_COM_GenPullSQL("Chart_Coles", , , , , CompCode, CStr(sStateLook))
    End If
End If
    
    'Range(wks_ChartData.Cells(2, 1), wks_ChartData.Cells(19999, 199)).ClearContents
    
    
    
    If bOutput = True Then
        wks_ChartData.Cells(1, 4).Value = "Price"
        wks_ChartData.Cells(1, 5).Value = "Promotion"
        wks_ChartData.Cells(1, 6).Value = "Full Price"
        For a = 0 To UBound(CBA_COMarr, 2)
            For b = 0 To UBound(CBA_COMarr, 1)
                wks_ChartData.Cells(a + 2, b + 1) = CBA_COMarr(b, a)
            Next
        Next
    End If


    Call CBAR_PPHChartCreate.ChartCreate(200, 200, Range(wks_ChartData.Cells(2, 4), wks_ChartData.Cells(UBound(CBA_COMarr, 2), 6)), wks_ChartData, CStr(compDesc), Range(wks_ChartData.Cells(1, 4), wks_ChartData.Cells(1, 6)), , xlRight, xlXYScatterLines, , xlColumns, False, 670 / 28.34645669, 425 / 28.34645669)
    For Each c In wks_ChartData.ChartObjects
    Set MyChart = c.Chart
    Next
    imagename = "C:\TEMP\TempComChart.gif"
    MyChart.Export Filename:=imagename, FilterName:="GIF"
    CBA_COM_frm_Chart.img_Chart.Picture = LoadPicture(imagename)
    Kill imagename
    For Each c In wks_ChartData.ChartObjects
    c.Delete
    Next
    Application.DisplayAlerts = False
    wbk.Close
    Application.DisplayAlerts = True
    
    Me.StartUpPosition = 0
    Me.Top = Application.Top + (Application.Height / 2) - (Me.Height / 2)
    Me.Left = Application.Left + (Application.Width / 2) - (Me.Width / 2)
    
Application.ScreenUpdating = True
End Sub




