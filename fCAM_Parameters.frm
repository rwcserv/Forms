VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fCAM_Parameters 
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13335
   OleObjectBlob   =   "fCAM_Parameters.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "fCAM_Parameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit             'fCam_Parameters
Private pbIsActive As Boolean

Private Sub cbo_CGSel_Change()
    Call mw_RemoveBoxHook
End Sub

Private Sub cbo_CGSel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call mw_SetBoxHook(Me.cbo_CGSel) '%%RW
End Sub

Private Sub cbo_ReportSelect_Change()
    Call mw_RemoveBoxHook
    If cbo_ReportSelect <> "" Then
        ''If mCAM_Runtime.getCameraDocuTypes(cbo_ReportSelect.Value) = e_DocuType.eCoreRangeCategoryReview Then
        If mCAM_Runtime.cRibbonData.Get_eSysDocTypeCols(, "SH_ID", cbo_ReportSelect.Value) = e_DocuType.eCoreRangeCategoryReview Then
            Me.Width = 678.75
        Else
            Me.Width = 135.75
        End If
    End If
End Sub

Private Sub cbo_ReportSelect_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call mw_SetBoxHook(cbo_ReportSelect) '%%RW
End Sub

Sub UserForm_Initialize()
    Dim a As Long
    Dim Dte As Date, sFrom As String, sTo As String
    Me.StartUpPosition = 0
    Me.Top = Application.Top + (Application.Height / 2) - (Me.Height / 2)
    Me.Left = Application.Left + (Application.Width / 2) - (Me.Width / 2)
   
    Me.lblVersion.Caption = CBA_getVersionStatus(g_GetDB("Cam"), CBA_Cam_Ver, "Camera Tool", "Camera")  ' Get the latest version
    For a = 0 To 23
        Dte = DateSerial(Year(DateAdd("M", (a * -1), Date)), Month(DateAdd("M", (a * -1), Date)), 0)
        Me.cbo_MonthFrom.AddItem Format(Dte, "MMMM-YYYY")
        Me.cbo_MonthTo.AddItem Format(Dte, "MMMM-YYYY")
        If a = 0 Then sTo = Format(Dte, "MMMM-YYYY")
        If a = 11 Then sFrom = Format(Dte, "MMMM-YYYY")
    Next
    Me.cbo_MonthTo.Value = sTo
    Me.cbo_MonthFrom.Value = sFrom
    Me.cbo_ReportSelect.List = mCAM_Runtime.cRibbonData.Get_eSysDocTypeCols(, "List")
    Me.cbo_CGSel.List = mCAM_Runtime.cRibbonData.GetAllCategoryNames
    Me.Width = 678.75 'Full width
    Me.Width = 268.5 ' International comparison selection width
    Me.Width = 135.75 'Basic selection width
    cbo_ACG.AddItem "Legacy CG"
    cbo_ACG.AddItem "ACG"
    bIsActive = True
End Sub

Private Sub cmd_Confirm_Click()
    Dim DocuType As e_DocuType
    Dim cDocObject As cCAM_Category
    Dim SucCreate As Boolean, bfound As Boolean
    Dim DocID As Long
    Dim v As Variant
    Call mw_RemoveBoxHook
    If Me.cbo_CGSel.ListIndex = -1 Then MsgBox "No CG's have been selected", vbOKOnly: Exit Sub
    If Me.cbo_MonthFrom = "" Then MsgBox "Month From is empty", vbOKOnly: Exit Sub
    If Me.cbo_MonthTo = "" Then MsgBox "Month To is empty", vbOKOnly: Exit Sub
    If Me.cbo_ACG = "" Then MsgBox "Please choose an ACG Setting", vbOKOnly: Exit Sub
''    If Me.txt_CatRevName = "" Then MsgBox "No Category Review description has been entered": exit
    DocID = -1
    DocuType = mCAM_Runtime.cRibbonData.Get_eSysDocTypeCols(, "SH_ID", cbo_ReportSelect.Value) ''mCAM_Runtime.getCameraDocuTypes(cbo_ReportSelect.Value)
    mCAM_Runtime.CheckDocHolder
    
    Select Case DocuType
        Case e_DocuType.eCoreRangeCategoryReview
            If Me.txt_CatRevName = "" Then
                MsgBox "No Category Review description has been entered"
                Exit Sub
            Else
                SucCreate = True
            End If
            bfound = False
            If bfound = False Then
                Me.Hide
                SucCreate = GenerateDataObject(DocuType, DocID, cDocObject)
            Else
                SucCreate = bfound
            End If
''            If SucCreate = True Then SucCreate = cDocObject.Construct(Me.cbo_CGSel.Value, CDate(cbo_MonthFrom), CDate(cbo_MonthTo), IIf(cbo_ACG = "ACG", True, False), DocuType)
            If SucCreate = True Then SucCreate = mCAM_Runtime.clsDocHolder.Add_Document(DocuType, Me.txt_CatRevName, DocID, cDocObject)
            If SucCreate = True Then mCAM_Runtime.clsDocHolder.Render_Document DocID
        
        Case e_DocuType.eSpecSeasPerformance
            bfound = False
'            If Not mCAM_Runtime.ALLGeneratedDataObjects Is Nothing Then
'                For Each v In mCAM_Runtime.ALLGeneratedDataObjects
'                    If v.sCategoryName = Me.cbo_CGSel.Value And v.bACG = IIf(cbo_ACG = "ACG", True, False) _
'                        And v.dtDateFrom = CDate(Format(cbo_MonthFrom, CBA_DMY)) And v.dtDateTo = DateSerial(Year(cbo_MonthTo), Month(cbo_MonthTo) + 1, 0) Then ' And v.sBuildCode = "BCRPTLWMs" Then
'                            bFound = True
'                            Set cDocObject = v
'                            Exit For
'                    End If
'                Next
'            End If
            If bfound = False Then
                Me.Hide
                SucCreate = GenerateDataObject(DocuType, DocID, cDocObject)
            Else
                SucCreate = bfound
            End If
            If SucCreate = True Then SucCreate = mCAM_Runtime.clsDocHolder.Add_Document(DocuType, Me.txt_CatRevName, DocID, cDocObject)
            If SucCreate = True Then mCAM_Runtime.clsDocHolder.Render_Document DocID
        Case e_DocuType.eLineCountOverviewReport
            bfound = False
'            If Not mCAM_Runtime.ALLGeneratedDataObjects Is Nothing Then
'                For Each v In mCAM_Runtime.ALLGeneratedDataObjects
'                    If v.sCategoryName = Me.cbo_CGSel.Value And v.bACG = IIf(cbo_ACG = "ACG", True, False) _
'                        And v.dtDateTo >= CDate(Format(cbo_MonthFrom, CBA_DMY)) And v.dtDateFrom <= DateSerial(Year(cbo_MonthTo), Month(cbo_MonthTo) + 1, 0) _
'                            And v.cPGrp.CompareBuildCode(mCAM_Runtime.cRibbonData.Get_eSysDocTypeCols(DocuType, "Buildcode")) = True Then
'                            bFound = True
'                            Set cDocObject = v
'                            Exit For
'                    End If
'                Next
'            End If
            If bfound = False Then
                Me.Hide
                SucCreate = GenerateDataObject(DocuType, DocID, cDocObject)
            Else
                SucCreate = bfound
            End If
            If SucCreate = True Then SucCreate = mCAM_Runtime.clsDocHolder.Add_Document(DocuType, Me.txt_CatRevName, DocID, cDocObject)
            If SucCreate = True Then mCAM_Runtime.clsDocHolder.Render_Document DocID
        Case e_DocuType.eCoreRangePerformance
            bfound = False
'            If Not mCAM_Runtime.ALLGeneratedDataObjects Is Nothing Then
'                For Each v In mCAM_Runtime.ALLGeneratedDataObjects
'                    If v.sCategoryName = Me.cbo_CGSel.Value And v.bACG = IIf(cbo_ACG = "ACG", True, False) _
'                        And v.dtDateTo >= CDate(Format(cbo_MonthFrom, CBA_DMY)) And v.dtDateFrom <= DateSerial(Year(cbo_MonthTo), Month(cbo_MonthTo) + 1, 0) _
'                            And v.cPGrp.CompareBuildCode(mCAM_Runtime.cRibbonData.Get_eSysDocTypeCols(DocuType, "Buildcode")) = True Then
'                            bFound = True
'                            Set cDocObject = v
'                            Exit For
'                    End If
'                Next
'            End If
            If bfound = False Then
                Me.Hide
                SucCreate = GenerateDataObject(DocuType, DocID, cDocObject)
            Else
                SucCreate = bfound
            End If
            If SucCreate = True Then SucCreate = mCAM_Runtime.clsDocHolder.Add_Document(DocuType, Me.txt_CatRevName, DocID, cDocObject)
            If SucCreate = True Then mCAM_Runtime.clsDocHolder.Render_Document DocID
        Case e_DocuType.eToplineCategoryPerformance
            bfound = False
'            If Not mCAM_Runtime.ALLGeneratedDataObjects Is Nothing Then
'                For Each v In mCAM_Runtime.ALLGeneratedDataObjects
'                    If v.sCategoryName = Me.cbo_CGSel.Value And v.bACG = IIf(cbo_ACG = "ACG", True, False) _
'                        And v.dtDateTo >= CDate(Format(cbo_MonthFrom, CBA_DMY)) And v.dtDateFrom <= DateSerial(Year(cbo_MonthTo), Month(cbo_MonthTo) + 1, 0) _
'                            And v.cPGrp.CompareBuildCode(mCAM_Runtime.cRibbonData.Get_eSysDocTypeCols(DocuType, "Buildcode")) = True Then
'                            bFound = True
'                            Set cDocObject = v
'                            Exit For
'                    End If
'                Next
'            End If
            If bfound = False Then
                Me.Hide
                SucCreate = GenerateDataObject(DocuType, DocID, cDocObject)
            Else
                SucCreate = bfound
            End If
            If SucCreate = True Then SucCreate = mCAM_Runtime.clsDocHolder.Add_Document(DocuType, Me.txt_CatRevName, DocID, cDocObject)
            If SucCreate = True Then mCAM_Runtime.clsDocHolder.Render_Document DocID
        Case e_DocuType.eMarketOverview
            bfound = False
'            If Not mCAM_Runtime.ALLGeneratedDataObjects Is Nothing Then
'                For Each v In mCAM_Runtime.ALLGeneratedDataObjects
'                    If v.sCategoryName = Me.cbo_CGSel.Value And v.bACG = IIf(cbo_ACG = "ACG", True, False) _
'                        And v.dtDateTo >= CDate(Format(cbo_MonthFrom, CBA_DMY)) And v.dtDateFrom <= DateSerial(Year(cbo_MonthTo), Month(cbo_MonthTo) + 1, 0) _
'                            And v.cPGrp.CompareBuildCode(mCAM_Runtime.cRibbonData.Get_eSysDocTypeCols(DocuType, "Buildcode")) = True Then
'                            bFound = True
'                            Set cDocObject = v
'                            Exit For
'                    End If
'                Next
'            End If
            If bfound = False Then
                Me.Hide
                SucCreate = GenerateDataObject(DocuType, DocID, cDocObject)
            Else
                SucCreate = bfound
            End If
            If SucCreate = True Then SucCreate = mCAM_Runtime.clsDocHolder.Add_Document(DocuType, Me.txt_CatRevName, DocID, cDocObject)
            If SucCreate = True Then mCAM_Runtime.clsDocHolder.Render_Document DocID
        Case e_DocuType.eCoreRangeProductListing
            bfound = False
'            If Not mCAM_Runtime.ALLGeneratedDataObjects Is Nothing Then
'                For Each v In mCAM_Runtime.ALLGeneratedDataObjects
'                    If v.sCategoryName = Me.cbo_CGSel.Value And v.bACG = IIf(cbo_ACG = "ACG", True, False) _
'                        And v.dtDateTo >= CDate(Format(cbo_MonthFrom, CBA_DMY)) And v.dtDateFrom <= DateSerial(Year(cbo_MonthTo), Month(cbo_MonthTo) + 1, 0) _
'                            And v.cPGrp.CompareBuildCode(mCAM_Runtime.cRibbonData.Get_eSysDocTypeCols(DocuType, "Buildcode")) = True Then
'                            bFound = True
'                            Set cDocObject = v
'                            Exit For
'                    End If
'                Next
'            End If
            If bfound = False Then
                Me.Hide
                SucCreate = GenerateDataObject(DocuType, DocID, cDocObject)
            Else
                SucCreate = bfound
            End If
            If SucCreate = True Then SucCreate = mCAM_Runtime.clsDocHolder.Add_Document(DocuType, Me.txt_CatRevName, DocID, cDocObject)
            If SucCreate = True Then mCAM_Runtime.clsDocHolder.Render_Document DocID
        Case e_DocuType.eForecast
            bfound = False
'            If Not mCAM_Runtime.ALLGeneratedDataObjects Is Nothing Then
'                For Each v In mCAM_Runtime.ALLGeneratedDataObjects
'                    If v.sCategoryName = Me.cbo_CGSel.Value And v.bACG = IIf(cbo_ACG = "ACG", True, False) _
'                        And v.dtDateTo >= CDate(Format(cbo_MonthFrom, CBA_DMY)) And v.dtDateFrom <= DateSerial(Year(cbo_MonthTo), Month(cbo_MonthTo) + 1, 0) _
'                            And v.cPGrp.CompareBuildCode(mCAM_Runtime.cRibbonData.Get_eSysDocTypeCols(DocuType, "Buildcode")) = True Then
'                            bFound = True
'                            Set cDocObject = v
'                            Exit For
'                    End If
'                Next
'            End If
            If bfound = False Then
                Me.Hide
                SucCreate = GenerateDataObject(DocuType, DocID, cDocObject)
            Else
                SucCreate = bfound
            End If
            If SucCreate = True Then SucCreate = mCAM_Runtime.clsDocHolder.Add_Document(DocuType, Me.txt_CatRevName, DocID, cDocObject)
            If SucCreate = True Then mCAM_Runtime.clsDocHolder.Render_Document DocID
    End Select

    CBA_BasicFunctions.CBA_Close_Running
    Unload Me
End Sub
Private Function GenerateDataObject(ByVal DocuType As e_DocuType, ByRef DocID As Long, ByRef cDocObject As cCAM_Category, Optional ByVal BuildCode As String = "") As Boolean
    Set cDocObject = New cCAM_Category
    CBA_BasicFunctions.CBA_Running "Generating Category Data Object"
    GenerateDataObject = cDocObject.Construct(Me.cbo_CGSel.Value, CDate(Format(cbo_MonthFrom, CBA_DMY)), DateSerial(Year(cbo_MonthTo), Month(cbo_MonthTo) + 1, 0), _
                                                        IIf(cbo_ACG = "ACG", True, False), DocuType, mCAM_Runtime.cRibbonData.Get_eSysDocTypeCols(DocuType, "Buildcode"))
    mCAM_Runtime.AddToALLGeneratedDataObjects cDocObject
End Function
Sub UserForm_Terminate()
    ''Unload Me
    Call mw_SetBoxHook(cbo_ReportSelect) '%%RW
End Sub

Public Property Get bIsActive() As Boolean: bIsActive = pbIsActive: End Property
Private Property Let bIsActive(ByVal bNewValue As Boolean): pbIsActive = bNewValue: End Property

