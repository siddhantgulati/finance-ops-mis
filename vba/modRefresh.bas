Attribute VB_Name = "modRefresh"
Option Explicit

' ---------------------------------------------------------------------------
'  Finance Operations MIS - refresh routine
'
'  RefreshMIS          points every Power Query at the \data folder next to
'                      this workbook, refreshes the tables in order, stamps the
'                      refresh time and shows the reconciliation result.
'  ExportDashboardPDF  saves the Dashboard sheet as a PDF in \output.
'
'  Import: Alt+F11 > File > Import File > modRefresh.bas, then save as .xlsm
' ---------------------------------------------------------------------------

Private Const FOLDER_TAG As String = "Folder = """

Public Sub RefreshMIS()
    Dim started As Double
    Dim dataFolder As String
    Dim q As WorkbookQuery
    Dim oldCalc As XlCalculation
    Dim missing As String

    started = Timer
    dataFolder = ThisWorkbook.Path & Application.PathSeparator & "data" & Application.PathSeparator

    missing = MissingFiles(dataFolder, Array("invoices.csv", "payments.csv", "clients.csv"))
    If Len(missing) > 0 Then
        MsgBox "These files are missing from " & dataFolder & vbCrLf & vbCrLf & missing, _
               vbExclamation, "Refresh MIS"
        Exit Sub
    End If

    oldCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.StatusBar = "Refreshing MIS..."
    Application.Calculation = xlCalculationManual
    On Error GoTo Fail

    ' Queries keep an absolute path, so re-point them in case the file was moved
    For Each q In ThisWorkbook.Queries
        q.Formula = SetFolder(q.Formula, dataFolder)
    Next q

    ' Lookups depend on the client master, and invoices depend on payments,
    ' so load the masters first and calculate once at the end
    RefreshTable "ClientMaster", "tblClientMaster"
    RefreshTable "Invoices", "tblInvoices"
    RefreshTable "Payments", "tblPayments"

    Application.Calculation = oldCalc
    Application.CalculateFull

    ThisWorkbook.Names("LastRefresh").RefersToRange.Value = Now
    ThisWorkbook.Names("RefreshSecs").RefersToRange.Value = Round(Timer - started, 1)

    Application.StatusBar = False
    Application.ScreenUpdating = True
    ThisWorkbook.Worksheets("Dashboard").Activate

    MsgBox "Refreshed in " & Format(Timer - started, "0.0") & " sec." & vbCrLf & _
           "Reconciliation: " & ThisWorkbook.Names("ChecksOverall").RefersToRange.Value, _
           vbInformation, "Refresh MIS"
    Exit Sub

Fail:
    Application.Calculation = oldCalc
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Refresh stopped: " & Err.Description, vbCritical, "Refresh MIS"
End Sub

Public Sub ExportDashboardPDF()
    Dim outFolder As String
    Dim outFile As String
    Dim asOf As Date

    outFolder = ThisWorkbook.Path & Application.PathSeparator & "output"
    If Dir(outFolder, vbDirectory) = "" Then MkDir outFolder

    asOf = ThisWorkbook.Names("AsOfDate").RefersToRange.Value
    outFile = outFolder & Application.PathSeparator & "MIS_Dashboard_" & Format(asOf, "mmm-yy") & ".pdf"

    ThisWorkbook.Worksheets("Dashboard").ExportAsFixedFormat _
        Type:=xlTypePDF, Filename:=outFile, Quality:=xlQualityStandard, _
        IncludeDocProperties:=False, IgnorePrintAreas:=False

    MsgBox "Saved " & outFile, vbInformation, "Export"
End Sub

Private Sub RefreshTable(sheetName As String, tableName As String)
    Application.StatusBar = "Refreshing " & sheetName & "..."
    ThisWorkbook.Worksheets(sheetName).ListObjects(tableName).QueryTable.Refresh BackgroundQuery:=False
End Sub

Private Function SetFolder(ByVal m As String, ByVal folder As String) As String
    ' Replaces the text between  Folder = "  and the next quote in the M code
    Dim p1 As Long, p2 As Long
    p1 = InStr(1, m, FOLDER_TAG, vbTextCompare)
    If p1 = 0 Then
        SetFolder = m
        Exit Function
    End If
    p1 = p1 + Len(FOLDER_TAG)
    p2 = InStr(p1, m, """")
    SetFolder = Left$(m, p1 - 1) & folder & Mid$(m, p2)
End Function

Private Function MissingFiles(ByVal folder As String, files As Variant) As String
    Dim f As Variant
    For Each f In files
        If Dir(folder & f) = "" Then MissingFiles = MissingFiles & "  - " & f & vbCrLf
    Next f
End Function
