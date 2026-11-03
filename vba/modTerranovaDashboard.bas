Attribute VB_Name = "modTerranovaDashboard"
'==============================================================================
' Module  : modTerranovaDashboard
' Project : Terranova 4G 100 - Excel Executive Dashboard
' Purpose : Automates the recurring reporting cycle of the dashboard:
'             1. Refresh data connections + full recalculation
'             2. Validate the 6 model-integrity checks (Control!D20:D25)
'             3. Log a KPI snapshot (history for trend reporting)
'             4. Export one PDF executive report per province filter
'             5. Export the analytical report pack (one PDF per sheet)
'             6. Protect / unprotect the dashboard sheets (inputs stay editable)
' Author  : Fernando Munoz - FxMr Consulting
' Version : 1.0
' Setup   : Save the workbook as .xlsm, then import this module
'           (VBA editor > File > Import File...).
' Names   : Uses the workbook names StatusDate, SelProvince, SelSite,
'           SelScenario, GreenThr, AmberThr, FX_Base, OnTimeTol.
'==============================================================================
Option Explicit

Private Const SH_CONTROL As String = "Control"
Private Const SH_EXEC    As String = "Exec_Dashboard"
Private Const SH_COST    As String = "Cost_Analysis"
Private Const SH_LOG     As String = "KPI_Log"
Private Const OUT_FOLDER As String = "Reports"
Private Const APP_TITLE  As String = "Terranova Dashboard"

'------------------------------------------------------------------------------
' MASTER ROUTINE - assign this one to a button on the Control sheet
'------------------------------------------------------------------------------
Public Sub RunMonthlyReportCycle()
    If Not RefreshAndValidate() Then Exit Sub
    LogKpiSnapshot
    ExportProvinceReports
    ExportReportPack
    MsgBox "Reporting cycle completed." & vbCrLf & vbCrLf & _
           "Output folder:" & vbCrLf & ReportFolder(), vbInformation, APP_TITLE
End Sub

'------------------------------------------------------------------------------
' 1. REFRESH + VALIDATE
'------------------------------------------------------------------------------
Public Function RefreshAndValidate() As Boolean
    Dim t0 As Double, ok As Boolean
    t0 = Timer
    On Error GoTo EH
    Application.ScreenUpdating = False
    Application.StatusBar = "Refreshing data connections and recalculating..."
    ThisWorkbook.RefreshAll                      'Power Query + Data Model + PivotTables
    Application.CalculateUntilAsyncQueriesDone   'wait for background queries
    Application.CalculateFull
    ok = ModelIsHealthy(True)
    RefreshAndValidate = ok
    Application.StatusBar = "Refresh finished in " & Format(Timer - t0, "0.0") & _
        " s | model integrity: " & IIf(ok, "OK", "CHECK")
CleanUp:
    Application.ScreenUpdating = True
    Exit Function
EH:
    MsgBox "Refresh failed: " & Err.Description, vbCritical, APP_TITLE
    Resume CleanUp
End Function

' Reads the six reconciliation checks on the Control sheet (rows 20-25).
Public Function ModelIsHealthy(Optional ByVal showMessage As Boolean = True) As Boolean
    Dim ws As Worksheet, r As Long, failed As String
    Set ws = ThisWorkbook.Worksheets(SH_CONTROL)
    For r = 20 To 25
        If Trim$(CStr(ws.Cells(r, "D").Value)) <> "OK" Then
            failed = failed & vbCrLf & " - " & CStr(ws.Cells(r, "B").Value)
        End If
    Next r
    ModelIsHealthy = (Len(failed) = 0)
    If Not ModelIsHealthy And showMessage Then
        MsgBox "Model integrity checks failed:" & failed & vbCrLf & vbCrLf & _
               "Reports were not generated. Review the Control sheet.", _
               vbExclamation, APP_TITLE
    End If
End Function

'------------------------------------------------------------------------------
' 2. KPI SNAPSHOT LOG  (builds a time series of the executive KPIs)
'------------------------------------------------------------------------------
Public Sub LogKpiSnapshot()
    Dim wsX As Worksheet, wsLog As Worksheet, r As Long, rag As String
    Set wsX = ThisWorkbook.Worksheets(SH_EXEC)
    Set wsLog = GetOrCreateLog()
    r = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row + 1
    rag = Split(CStr(wsX.Range("L8").Value), " ")(0)      'GREEN / AMBER / RED

    wsLog.Cells(r, 1).Value = Now
    wsLog.Cells(r, 2).Value = NamedCell("StatusDate").Value
    wsLog.Cells(r, 3).Value = NamedCell("SelProvince").Value
    wsLog.Cells(r, 4).Value = wsX.Range("B7").Value        'sites
    wsLog.Cells(r, 5).Value = wsX.Range("D7").Value        'investment USD
    wsLog.Cells(r, 6).Value = wsX.Range("H7").Value        'milestones completed %
    wsLog.Cells(r, 7).Value = wsX.Range("J7").Value        'on-time % of completed
    wsLog.Cells(r, 8).Value = wsX.Range("L7").Value        'SPI
    wsLog.Cells(r, 9).Value = rag
    wsLog.Cells(r, 10).Value = CStr(ThisWorkbook.Worksheets(SH_CONTROL).Range("D26").Value)

    Select Case rag
        Case "GREEN": wsLog.Cells(r, 9).Interior.Color = RGB(198, 239, 206)
        Case "AMBER": wsLog.Cells(r, 9).Interior.Color = RGB(255, 235, 156)
        Case "RED":   wsLog.Cells(r, 9).Interior.Color = RGB(255, 199, 206)
    End Select
End Sub

Private Function GetOrCreateLog() As Worksheet
    Dim ws As Worksheet, prev As Object
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_LOG)
    On Error GoTo 0
    If ws Is Nothing Then
        Set prev = ActiveSheet
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SH_LOG
        ws.Range("A1:J1").Value = Array("Logged at", "Status date", "Province filter", "Sites", _
            "Investment USD", "Milestones completed %", "On-time % of completed", "SPI", "RAG", "Model integrity")
        With ws.Range("A1:J1")
            .Font.Bold = True
            .Font.Color = vbWhite
            .Interior.Color = RGB(31, 56, 100)
            .WrapText = True
        End With
        ws.Columns("A").NumberFormat = "yyyy-mm-dd hh:mm"
        ws.Columns("B").NumberFormat = "dd-mmm-yyyy"
        ws.Columns("E").NumberFormat = "$#,##0"
        ws.Columns("F:G").NumberFormat = "0.0%"
        ws.Columns("H").NumberFormat = "0.00"
        ws.Columns("A:J").ColumnWidth = 18
        prev.Activate
    End If
    Set GetOrCreateLog = ws
End Function

'------------------------------------------------------------------------------
' 3. PROVINCE-BY-PROVINCE EXECUTIVE PDF REPORTS
'------------------------------------------------------------------------------
Public Sub ExportProvinceReports()
    Dim folder As String, stamp As String, original As Variant, msg As String
    Dim rngSel As Range, wsX As Worksheet, p As Variant, n As Long

    If Not ModelIsHealthy(True) Then Exit Sub
    folder = ReportFolder()
    If Len(folder) = 0 Then
        MsgBox "Save the workbook before exporting reports.", vbExclamation, APP_TITLE
        Exit Sub
    End If

    Set wsX = ThisWorkbook.Worksheets(SH_EXEC)
    Set rngSel = NamedCell("SelProvince")
    original = rngSel.Value
    stamp = Format(Now, "yyyymmdd_hhnn")

    On Error GoTo EH
    Application.ScreenUpdating = False
    For Each p In ProvinceList()
        rngSel.Value = p                    'one filter drives every KPI and chart
        Application.Calculate
        wsX.ExportAsFixedFormat Type:=xlTypePDF, _
            Filename:=folder & "Terranova_Exec_" & SafeName(CStr(p)) & "_" & stamp & ".pdf", _
            Quality:=xlQualityStandard, IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, OpenAfterPublish:=False
        n = n + 1
    Next p
    rngSel.Value = original
    Application.Calculate
    Application.ScreenUpdating = True
    MsgBox n & " province reports exported to:" & vbCrLf & folder, vbInformation, APP_TITLE
    Exit Sub
EH:
    msg = Err.Description
    On Error Resume Next
    rngSel.Value = original
    Application.Calculate
    Application.ScreenUpdating = True
    MsgBox "Province export failed: " & msg, vbCritical, APP_TITLE
End Sub

' Builds ["All", province 1 ... province n] from the Cost_Analysis matrix (rows 7..).
Private Function ProvinceList() As Collection
    Dim c As New Collection, ws As Worksheet, r As Long, v As String
    Set ws = ThisWorkbook.Worksheets(SH_COST)
    c.Add "All"
    r = 7
    Do
        v = Trim$(CStr(ws.Cells(r, "B").Value))
        If Len(v) = 0 Or v = "Programme total" Then Exit Do
        c.Add v
        r = r + 1
    Loop
    Set ProvinceList = c
End Function

'------------------------------------------------------------------------------
' 4. ANALYTICAL REPORT PACK (one PDF per analysis sheet, current filters)
'------------------------------------------------------------------------------
Public Sub ExportReportPack()
    Dim folder As String, stamp As String, nm As Variant, msg As String, n As Long

    If Not ModelIsHealthy(True) Then Exit Sub
    folder = ReportFolder()
    If Len(folder) = 0 Then
        MsgBox "Save the workbook before exporting reports.", vbExclamation, APP_TITLE
        Exit Sub
    End If
    stamp = Format(Now, "yyyymmdd_hhnn")

    On Error GoTo EH
    Application.ScreenUpdating = False
    For Each nm In ReportSheets()
        ThisWorkbook.Worksheets(CStr(nm)).ExportAsFixedFormat Type:=xlTypePDF, _
            Filename:=folder & "Terranova_" & CStr(nm) & "_" & stamp & ".pdf", _
            Quality:=xlQualityStandard, IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, OpenAfterPublish:=False
        n = n + 1
    Next nm
    Application.ScreenUpdating = True
    MsgBox n & " report-pack PDFs exported to:" & vbCrLf & folder, vbInformation, APP_TITLE
    Exit Sub
EH:
    msg = Err.Description
    Application.ScreenUpdating = True
    MsgBox "Report pack export failed: " & msg, vbCritical, APP_TITLE
End Sub

' One-off: landscape, fit-to-width, footer with page numbers on every report sheet.
Public Sub ConfigurePrintLayout()
    Dim nm As Variant
    On Error GoTo EH
    For Each nm In ReportSheets()
        With ThisWorkbook.Worksheets(CStr(nm)).PageSetup
            .Orientation = xlLandscape
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False
            .LeftFooter = "Terranova 4G 100 | &A"
            .RightFooter = "Page &P of &N"
        End With
    Next nm
    MsgBox "Print layout applied to " & (UBound(ReportSheets()) + 1) & " sheets.", vbInformation, APP_TITLE
    Exit Sub
EH:
    MsgBox "Could not set the print layout (a printer driver must be installed): " & _
           Err.Description, vbExclamation, APP_TITLE
End Sub

Private Function ReportSheets() As Variant
    ReportSheets = Array(SH_EXEC, SH_COST, "Schedule_Analysis", "EVM_SCurve", "Scenario")
End Function

'------------------------------------------------------------------------------
' 5. SHEET PROTECTION - formulas locked, input cells stay editable
'------------------------------------------------------------------------------
Public Sub ProtectDashboardSheets()
    Dim sh As Variant, nm As Variant
    On Error GoTo EH
    For Each sh In DashboardSheets()
        ThisWorkbook.Worksheets(CStr(sh)).Unprotect
    Next sh
    For Each nm In Array("StatusDate", "OnTimeTol", "FX_Base", "GreenThr", "AmberThr", _
                         "SelProvince", "SelSite", "SelScenario")
        NamedCell(CStr(nm)).Locked = False
    Next nm
    ThisWorkbook.Worksheets("Scenario").Range("C10:H12").Locked = False   'scenario assumptions
    For Each sh In DashboardSheets()
        ThisWorkbook.Worksheets(CStr(sh)).Protect DrawingObjects:=True, Contents:=True, _
            Scenarios:=True, UserInterfaceOnly:=True
    Next sh
    MsgBox "Dashboard sheets protected. Inputs remain editable.", vbInformation, APP_TITLE
    Exit Sub
EH:
    MsgBox "Protection failed: " & Err.Description, vbCritical, APP_TITLE
End Sub

Public Sub UnprotectDashboardSheets()
    Dim sh As Variant
    On Error GoTo EH
    For Each sh In DashboardSheets()
        ThisWorkbook.Worksheets(CStr(sh)).Unprotect
    Next sh
    MsgBox "Dashboard sheets unprotected.", vbInformation, APP_TITLE
    Exit Sub
EH:
    MsgBox "Unprotect failed (password set?): " & Err.Description, vbCritical, APP_TITLE
End Sub

Private Function DashboardSheets() As Variant
    DashboardSheets = Array(SH_CONTROL, SH_EXEC, SH_COST, "Schedule_Analysis", _
                            "EVM_SCurve", "Site_Profile", "Scenario")
End Function

'------------------------------------------------------------------------------
' HELPERS
'------------------------------------------------------------------------------
Private Function NamedCell(ByVal nm As String) As Range
    Set NamedCell = ThisWorkbook.Names(nm).RefersToRange
End Function

Private Function ReportFolder() As String
    Dim p As String
    p = ThisWorkbook.Path
    If Len(p) = 0 Then Exit Function                       'workbook never saved
    p = p & Application.PathSeparator & OUT_FOLDER & Application.PathSeparator
    If Len(Dir(p, vbDirectory)) = 0 Then MkDir p
    ReportFolder = p
End Function

Private Function SafeName(ByVal s As String) As String
    Dim ch As Variant
    For Each ch In Array("\", "/", ":", "*", "?", """", "<", ">", "|", " ")
        s = Replace(s, CStr(ch), "_")
    Next ch
    SafeName = s
End Function
