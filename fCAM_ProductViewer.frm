VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fCAM_ProductViewer 
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10830
   OleObjectBlob   =   "fCAM_ProductViewer.frx":0000
   StartUpPosition =   2  'CenterScreen
   Tag             =   "fCAM_ProductViewer"
End
Attribute VB_Name = "fCAM_ProductViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit                             ' fCam_ProductViewer
Private pdtDFrom As Date
Private pdtDTo As Date
Private pcCP As cCBA_Prod
Private pbIsActive As Boolean

Private Sub cbo_Period_Change()

    Select Case Me.cbo_Period.Value
        Case "MAT"
            PopulatePeriodDataBoxes dtDFrom, dtDTo
        Case "MAT PY"
            PopulatePeriodDataBoxes DateAdd("YYYY", -1, dtDFrom), DateAdd("YYYY", -1, dtDTo)
        Case CStr(Year(Date) - 1)
            PopulatePeriodDataBoxes DateSerial(Year(dtDTo) - 1, 1, 1), DateSerial(Year(dtDTo) - 1, 12, 31)
        Case CStr(Year(Date) - 2)
            PopulatePeriodDataBoxes DateSerial(Year(dtDTo) - 2, 1, 1), DateSerial(Year(dtDTo) - 2, 12, 31)
    End Select
    DoEvents
End Sub
Sub UserForm_Terminate()
    Unload Me
End Sub
Sub UserForm_Initialize()
    Dim lTop, lLeft, lRow, lcol As Long
    Me.StartUpPosition = 0
    Me.Top = Application.Top + (Application.Height / 2) - (Me.Height / 2)
    Me.Left = Application.Left + (Application.Width / 2) - (Me.Width / 2)
    
    'Me.lblVersion.Caption = CBA_getVersionStatus(g_getDB("Cam"), CBA_Cam_Ver, "Camera Tool", "Camera")  ' Get the latest version
    'Me.Hide
End Sub

Public Function SetData(ByVal PCode As String) As Boolean
    ' Set the private variables and the cbo_Period
    Dim PCde As String
    Dim DateFrom As Date, DateTo As Date
    PCde = Left(PCode, InStr(1, PCode, "-") - 1)
    If IsNumeric(PCde) Then Set cCP = mCAM_Runtime.getCBA_ProdEntity(PCde) Else SetData = False: bIsActive = False: Call UserForm_Terminate: Exit Function
    
    dtDTo = mCAM_Runtime.cActiveDataObject.dtDateTo
    dtDFrom = DateAdd("D", 1, DateAdd("YYYY", -1, dtDTo))
    With cCP
        Me.lbl_PCode.Caption = .lPcode
        Me.lbl_PDesc.Caption = .sProdDesc
        If dtDFrom = 0 Then
            cbo_Period.Enabled = False
            SetData = False
            bIsActive = False
        Else
            cbo_Period.Enabled = True
            Me.cbo_Period.AddItem "MAT"
            Me.cbo_Period.AddItem "MAT PY"
            Me.cbo_Period.AddItem Year(Date) - 1
            Me.cbo_Period.AddItem Year(Date) - 2
            Me.cbo_Period.Value = "MAT"
            SetData = True
            bIsActive = True
        End If
    End With
    
End Function
Private Function PopulatePeriodDataBoxes(ByVal DateFrom As Date, ByVal DateTo As Date) As Boolean
    ' Populate all Text Boxes dependant on the value in cbo_Period
    Dim s As Single
    On Error GoTo Err_Routine
    CBA_Error = ""
    PopulatePeriodDataBoxes = False
    With cCP
        Me.lbl_POSQTY.Caption = Format(.getPOSdata(DateFrom, DateTo, True), "#,0")
        s = .getPOSdata(DateAdd("YYYY", -1, DateFrom), DateAdd("YYYY", -1, DateTo), True)
        If s = 0 Then Me.lbl_POSQTYYOY.Caption = "-" Else Me.lbl_POSQTYYOY.Caption = Format(((Me.lbl_POSQTY.Caption - s) / s), "0%")
        Me.lbl_RCVQTY.Caption = Format(.getRCVdata(DateFrom, DateTo, "QTY"), "#,0")
        s = .getRCVdata(DateAdd("YYYY", -1, DateFrom), DateAdd("YYYY", -1, DateTo), "QTY")
        If s = 0 Then Me.lbl_RCVQTYYOY.Caption = "-" Else Me.lbl_RCVQTYYOY.Caption = Format(((Me.lbl_RCVQTY.Caption - s) / s), "0%")
        Me.lbl_RCVRet.Caption = Format(.getRCVdata(DateFrom, DateTo, "Retail"), "$#,0")
        s = .getRCVdata(DateAdd("YYYY", -1, DateFrom), DateAdd("YYYY", -1, DateTo), "Retail")
        If s = 0 Then Me.lbl_RCVRetYOY.Caption = "-" Else Me.lbl_RCVRetYOY.Caption = Format(((CCur(Me.lbl_RCVRet.Caption) - s) / s), "0%")
        Me.lbl_RCVCost.Caption = Format(.getRCVdata(DateFrom, DateTo, "Cost"), "$#,0")
        s = .getRCVdata(DateAdd("YYYY", -1, DateFrom), DateAdd("YYYY", -1, DateTo), "Cost")
        If s = 0 Then Me.lbl_RCVCostYOY.Caption = "-" Else Me.lbl_RCVCostYOY.Caption = Format((CCur(Me.lbl_RCVCost.Caption) - s) / s, "0%")
        Me.lbl_RCVCont.Caption = Format(.getRCVContribution(DateFrom, DateTo), "$#,0")
        s = .getRCVContribution(DateAdd("YYYY", -1, DateFrom), DateAdd("YYYY", -1, DateTo))
        If s = 0 Then Me.lbl_RCVContYOY.Caption = "-" Else Me.lbl_RCVContYOY.Caption = Format((CCur(Me.lbl_RCVCont.Caption) - s) / s, "0%")
        Me.lbl_RCVMargin.Caption = Format(.getRCVMargin(DateFrom, DateTo), "0.00%")
        s = .getRCVMargin(DateAdd("YYYY", -1, DateFrom), DateAdd("YYYY", -1, DateTo))
        If s = 0 Then Me.lbl_RCVMarginYOY.Caption = "-" Else Me.lbl_RCVMarginYOY.Caption = Format((.getRCVMargin(DateFrom, DateTo) - s) / s, "0%")
        Me.lbl_RCVShare.Caption = Format(.getRCVShare(DateFrom, DateTo, True) * 100, "0.00")
        s = .getRCVShare(DateAdd("YYYY", -1, DateFrom), DateAdd("YYYY", -1, DateTo), True)
        If s = 0 Then Me.lbl_RCVShareYOY.Caption = "-" Else Me.lbl_RCVShareYOY.Caption = Format((.getRCVShare(DateFrom, DateTo, True) - s) / s, "0%")
        Me.lbl_POSShare.Caption = Format(.getPOSShare(DateFrom, DateTo, True) * 100, "0.00")
        s = .getPOSShare(DateAdd("YYYY", -1, DateFrom), DateAdd("YYYY", -1, DateTo), True)
        If s = 0 Then Me.lbl_POSShareYOY.Caption = "-" Else Me.lbl_POSShareYOY.Caption = Format((.getPOSShare(DateFrom, DateTo, True) - s) / s, "0%")
''        Me.lbl_Wastage.Caption = Format(cCP.getWastageData(e_RetailorQTY.eRetail, e_InvMDALL.eBoth, DateFrom, DateTo), "$0.00")
        PopulateChartImage DateFrom, DateTo
        PopulatePeriodDataBoxes = True
    End With
    Me.Repaint
Exit_Routine:

    On Error Resume Next
    Exit Function
Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-fCam_ProductViewer.populatePeriodDataBoxes", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            '^RWCam Camera
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Cam", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next
End Function
Private Function PopulateChartImage(ByVal DateFrom As Date, ByVal DateTo As Date) As Boolean
    ' Creates the Chart object
    Dim bOutput As Boolean, a As Long, b As Long, c
    Dim MyChart As Chart
    Dim imagename As String
    Dim wks_ChartData As Worksheet
    Dim wbk As Workbook
    On Error GoTo Err_Routine
    CBA_Error = ""
    PopulateChartImage = False
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
    
    With cCP
    
        wks_ChartData.Cells(1, 1).Value = "Month"
        wks_ChartData.Cells(1, 2).Value = "POS Retail"
        wks_ChartData.Cells(1, 3).Value = "POS Retail (YOY)"
        wks_ChartData.Cells(1, 4).Value = "Margin%"
        wks_ChartData.Cells(1, 5).Value = "Contribution$"
        'wks_ChartData.Cells(1, 6).Value = "Wastage"
        For a = 0 To 11
            wks_ChartData.Cells(a + 2, 1).Value = a + 1
            wks_ChartData.Cells(a + 2, 2).Value = .getPOSdata(DateAdd("M", a, DateFrom), DateAdd("M", a + 1, DateFrom), True)
            wks_ChartData.Cells(a + 2, 3).Value = .getPOSdata(DateAdd("YYYY", -1, DateAdd("M", a, DateFrom)), DateAdd("YYYY", -1, DateAdd("M", a + 1, DateFrom)), True)
            wks_ChartData.Cells(a + 2, 4).Value = .getRCVMargin(DateAdd("M", a, DateFrom), DateAdd("M", a + 1, DateFrom))
            wks_ChartData.Cells(a + 2, 5).Value = .getRCVContribution(DateAdd("M", a, DateFrom), DateAdd("M", a + 1, DateFrom))
            'wks_ChartData.Cells(a + 2, 5).Value = .getWastage(Datefrom, DateTo)
        Next
    
    
        Call CBAR_PPHChartCreate.CBA_CR_LineProdChartCreate(200, 200, Range(wks_ChartData.Cells(1, 1), wks_ChartData.Cells(12, 5)), wks_ChartData, , _
            Range(wks_ChartData.Cells(1, 2), wks_ChartData.Cells(1, 5)), , xlRight, xlXYScatterLines, xlXYScatterLines, xlColumns, False, 520 / 28.34645669, 190 / 28.34645669)
    
        For Each c In wks_ChartData.ChartObjects
        Set MyChart = c.Chart
        Next
        imagename = "C:\TEMP\TempComChart.gif"
        MyChart.Export Filename:=imagename, FilterName:="GIF"
        Me.img_Graph.Picture = LoadPicture(imagename)
        Kill imagename
        For Each c In wks_ChartData.ChartObjects
        c.Delete
        Next
        Application.DisplayAlerts = False
        wbk.Close
        Application.DisplayAlerts = True
    
        Application.ScreenUpdating = True
        PopulateChartImage = True
    
    End With
Exit_Routine:

    On Error Resume Next
    Exit Function

Err_Routine:
    CBA_Erl = CLng(VBA.Erl): Call CBA_ProcI("f-fCam_ProductViewer.populateChartImage", 3)
    CBA_Error = " Error - " & Err.Number & "-" & Err.Description & "-" & CBA_ProcI(, 0)
    Debug.Print CBA_Error
    Stop            ' ^RW Camera + next line
    Resume Next
    Call g_FileWrite(g_GetDB("Gen", True), CBA_Error, , , True, True)
    Call g_Write_Err_Table(Err, CBA_Error, "Cam", CBA_ProcI(, 0, True), CBA_Erl, CBA_TestIP)
    GoTo Exit_Routine
    Resume Next

End Function

Private Property Get dtDFrom() As Date: dtDFrom = pdtDFrom: End Property
Private Property Let dtDFrom(ByVal dtNewValue As Date): pdtDFrom = dtNewValue: End Property

Private Property Get dtDTo() As Date: dtDTo = pdtDTo: End Property
Private Property Let dtDTo(ByVal dtNewValue As Date): pdtDTo = dtNewValue: End Property

Private Property Get cCP() As cCBA_Prod: Set cCP = pcCP: End Property
Private Property Set cCP(ByVal objNewValue As cCBA_Prod): Set pcCP = objNewValue: End Property

Private Property Get bIsActive() As Boolean: bIsActive = pbIsActive: End Property
Private Property Let bIsActive(ByVal bNewValue As Boolean): pbIsActive = bNewValue: End Property
