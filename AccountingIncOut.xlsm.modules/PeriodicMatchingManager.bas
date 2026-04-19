Attribute VB_Name = "PeriodicMatchingManager"
Option Explicit

Public Sub RunPeriodicMatchingWithFileSelection()
    Dim filePath As String
    Dim resultText As String

    filePath = Application.GetOpenFilename( _
        "Excel Files (*.xlsx),*.xlsx,CSV Files (*.csv),*.csv,All Files (*.*),*.*", _
        , LocalizationManager.GetText("Select 1C export file"))

    If filePath = "False" Then Exit Sub

    resultText = RunPeriodicMatchingFromFile(CStr(filePath))
    MsgBox resultText, vbInformation, LocalizationManager.GetText("1C Integration Results")
End Sub

Public Function RunPeriodicMatchingFromFile(ByVal filePath As String) As String
    Dim wb1C As Workbook
    Dim ws1C As Worksheet
    Dim packageProcessed As Long
    Dim packageFound As Long
    Dim packageSkipped As Long
    Dim packageMultiple As Long
    Dim packageNotFound As Long
    Dim legacyProcessed As Long
    Dim legacyExact As Long
    Dim legacyQueued As Long
    Dim legacyNotFound As Long
    Dim backfillParents As Long
    Dim backfillCreated As Long
    Dim backfillSkipped As Long
    Dim backfillReused As Long

    On Error GoTo RunError

    Call EnsurePackageDocumentsSchema
    Call LegacyMatchReviewManager.EnsureLegacyMatchReviewSchema
    Call LegacyPackageBackfillManager.EnsureLegacyPackageBackfillSchema

    Application.StatusBar = LocalizationManager.GetText("Opening 1C export file...")
    Set wb1C = Workbooks.Open(filePath, ReadOnly:=True)
    Set ws1C = wb1C.Worksheets(1)

    Call ProvodkaIntegrationModule.MassProcessFromWorksheet(ws1C, False, True, False, packageProcessed, packageFound, packageSkipped, packageMultiple, packageNotFound)
    Call LegacyMatchReviewManager.BuildLegacyMatchReviewQueueFromWorksheet(ws1C, True, legacyProcessed, legacyExact, legacyQueued, legacyNotFound)
    Call LegacyPackageBackfillManager.BuildLegacyPackageBackfillQueueFromWorksheet(ws1C, backfillParents, backfillCreated, backfillSkipped, backfillReused)

    wb1C.Close False
    Application.StatusBar = False

    RunPeriodicMatchingFromFile = BuildPeriodicMatchingSummary( _
        packageProcessed, packageFound, packageMultiple, packageNotFound, _
        legacyProcessed, legacyExact, legacyQueued, legacyNotFound, _
        backfillParents, backfillCreated, backfillReused, backfillSkipped)
    Exit Function

RunError:
    On Error Resume Next
    If Not wb1C Is Nothing Then wb1C.Close False
    Application.StatusBar = False
    RunPeriodicMatchingFromFile = LocalizationManager.GetText("Mass processing error:") & " " & Err.Number & " - " & Err.Description
End Function

Private Function BuildPeriodicMatchingSummary(ByVal packageProcessed As Long, ByVal packageFound As Long, ByVal packageMultiple As Long, ByVal packageNotFound As Long, ByVal legacyProcessed As Long, ByVal legacyExact As Long, ByVal legacyQueued As Long, ByVal legacyNotFound As Long, ByVal backfillParents As Long, ByVal backfillCreated As Long, ByVal backfillReused As Long, ByVal backfillSkipped As Long) As String
    BuildPeriodicMatchingSummary = _
        LocalizationManager.GetText("Periodic matching completed.") & vbCrLf & vbCrLf & _
        LocalizationManager.GetText("Package rows processed: ") & packageProcessed & vbCrLf & _
        LocalizationManager.GetText("Package rows matched: ") & packageFound & vbCrLf & _
        LocalizationManager.GetText("Package rows need review: ") & packageMultiple & vbCrLf & _
        LocalizationManager.GetText("Package rows not found: ") & packageNotFound & vbCrLf & vbCrLf & _
        LocalizationManager.GetText("Legacy flat rows processed: ") & legacyProcessed & vbCrLf & _
        LocalizationManager.GetText("Legacy exact matches written: ") & legacyExact & vbCrLf & _
        LocalizationManager.GetText("Legacy review rows queued: ") & legacyQueued & vbCrLf & _
        LocalizationManager.GetText("Legacy flat rows not found: ") & legacyNotFound & vbCrLf & vbCrLf & _
        LocalizationManager.GetText("Legacy package parents processed: ") & backfillParents & vbCrLf & _
        LocalizationManager.GetText("Backfill proposals created: ") & backfillCreated & vbCrLf & _
        LocalizationManager.GetText("Backfill proposals reused: ") & backfillReused & vbCrLf & _
        LocalizationManager.GetText("Backfill rows skipped: ") & backfillSkipped & vbCrLf & vbCrLf & _
        LocalizationManager.GetText("Use Legacy Match Review for flat multiple matches and Legacy Package Backfill for package recovery proposals.")
End Function
