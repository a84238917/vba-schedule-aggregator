' バージョン：v0.5.1
Option Explicit
' このモジュールは、マクロのメイン処理フローを制御します。
' 設定の読み込み、対象ファイルの処理、結果の出力、エラーハンドリングなど、全体の指揮を執ります。

' Public Sub: ExtractDataMain
' マクロのメイン実行プロシージャ
Public Sub ExtractDataMain()
    Dim startTime As Date
    startTime = Now()
    Dim wsConfig As Worksheet ' For initial config reading
    Dim tempConfigSheetName As String
    Dim tempErrorLogSheetName As String

    ' --- Phase 1: Initialize and Prepare Error Logging ASAP ---
    g_nextErrorLogRow = 1 ' Initialize error log row counter

    On Error GoTo ErrorHandler_ExtractDataMain ' Changed label for clarity

    ' 1a. Initialize basic config structure
    Call InitializeConfigStructure(g_configSettings)
    g_configSettings.ScriptFullName = ThisWorkbook.FullName
    tempConfigSheetName = CONFIG_SHEET_DEFAULT_NAME ' Default, can be refined if O46 is read here too

    ' 1b. Get Config Worksheet Object
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Worksheets(tempConfigSheetName)
    On Error GoTo ErrorHandler_ExtractDataMain
    If wsConfig Is Nothing Then
        MsgBox "Configシート「" & tempConfigSheetName & "」が見つかりません。処理を中断します。", vbCritical, "致命的エラー"
        GoTo FinalizeRoutine_ExtractDataMain ' Changed label
    End If
    g_configSettings.configSheetName = wsConfig.Name ' Store actual name
    g_configSettings.ConfigSheetFullName = ThisWorkbook.FullName & " | " & wsConfig.Name


    ' 1c. Read ONLY ErrorLogSheetName from Config sheet
    ' Assuming M02_ConfigReader.ReadStringCell can be used carefully or implement a direct read here
    tempErrorLogSheetName = Trim(CStr(wsConfig.Range("O45").value))
    If tempErrorLogSheetName = "" Then tempErrorLogSheetName = "エラーログ" ' Fallback default
    g_configSettings.ErrorLogSheetName = tempErrorLogSheetName

    ' 1d. Prepare Error Log Sheet (g_errorLogWorksheet will be set globally by this call)
    ' This call must ensure g_errorLogWorksheet is set, even if M04_LogWriter is used internally by M03
    Call M03_SheetManager.PrepareErrorLogSheet(g_configSettings, ThisWorkbook)
    If g_errorLogWorksheet Is Nothing Then
        MsgBox "エラーログシートの準備に失敗しました。指定されたシート名: 「" & tempErrorLogSheetName & "」。" & vbCrLf & _
               "シート名に不正な文字がないか、ブックやシートが保護されていないか等を確認してください。" & vbCrLf & _
               "詳細はVBEのイミディエイトウィンドウの出力を確認してください。処理を中断します。", vbCritical, "致命的エラー"
        GoTo FinalizeRoutine_ExtractDataMain
    End If
    Call M04_LogWriter.WriteErrorLog("INFORMATION", "MainControl", "ExtractDataMain", "エラーログシート準備完了。")


    ' --- Phase 2: Load Full Configuration ---
    ' Now that error logging is set up, proceed with full configuration loading
    If Not M02_ConfigReader.LoadConfiguration(g_configSettings, ThisWorkbook) Then
        Call M04_LogWriter.WriteErrorLog("CRITICAL", "MainControl", "ExtractDataMain", "設定の読み込みに失敗しました。処理を中断します。")
        MsgBox "設定の読み込みに失敗しました。処理を中断します。", vbCritical, "設定エラー"
        GoTo FinalizeRoutine_ExtractDataMain
    End If
    Set g_configSettings.MainWorkbookObject = ThisWorkbook ' Set after successful LoadConfiguration

    ' --- Phase 3: Prepare Remaining Log Sheets and Initial Log Entries ---
    Call M03_SheetManager.PrepareRemainingLogSheets(g_configSettings, ThisWorkbook)

    Call M04_LogWriter.WriteOperationLog(g_configSettings, ThisWorkbook)
    Call M04_LogWriter.WriteFilterLog(g_configSettings, ThisWorkbook)

    ' --- Phase 4: Main Data Extraction Loop ---
    Dim k As Long
    Dim fileSystemObj As Object
    Set fileSystemObj = CreateObject("Scripting.FileSystemObject")
    Dim targetPath As String
    Dim currentFile As Object
    Dim filesInFolder As Object
    Dim targetPattern As String

    If g_configSettings.DebugDetailLevel1Enabled Then
        Call M04_LogWriter.WriteErrorLog("DEBUG_POINT", "MainControl", "ExtractDataMain", "処理対象ファイルループ開始前。TargetFileFolderPaths IsArray: " & IsArray(g_configSettings.TargetFileFolderPaths) & ", Initialized: " & General_IsArrayInitialized(g_configSettings.TargetFileFolderPaths))
        If General_IsArrayInitialized(g_configSettings.TargetFileFolderPaths) Then
             Call M04_LogWriter.WriteErrorLog("DEBUG_POINT", "MainControl", "ExtractDataMain", "TargetFileFolderPaths LBound: " & LBound(g_configSettings.TargetFileFolderPaths) & ", UBound: " & UBound(g_configSettings.TargetFileFolderPaths))
        End If
    End If

    If Not General_IsArrayInitialized(g_configSettings.TargetFileFolderPaths) Or _
       (LBound(g_configSettings.TargetFileFolderPaths) > UBound(g_configSettings.TargetFileFolderPaths)) Then
        Call M04_LogWriter.WriteErrorLog("INFORMATION", "MainControl", "ExtractDataMain", "処理対象ファイル/フォルダパスリスト(TargetFileFolderPaths)が空または未初期化です。処理をスキップします。")
        MsgBox "処理対象のファイルまたはフォルダが設定されていません。", vbInformation, "情報"
        GoTo FinalizeRoutine_ExtractDataMain
    End If

    For k = LBound(g_configSettings.TargetFileFolderPaths) To UBound(g_configSettings.TargetFileFolderPaths)
        If g_configSettings.DebugDetailLevel1Enabled Then Call M04_LogWriter.WriteErrorLog("DEBUG_POINT", "MainControl", "ExtractDataMain", "ループ開始。インデックス k = " & k)
        targetPath = Trim(g_configSettings.TargetFileFolderPaths(k))
        If g_configSettings.DebugDetailLevel1Enabled Then Call M04_LogWriter.WriteErrorLog("DEBUG_POINT", "MainControl", "ExtractDataMain", "targetPath = " & targetPath)

        targetPattern = ""
        If General_IsArrayInitialized(g_configSettings.FilePatternIdentifiers) Then
            Dim l_fp As Long, u_fp As Long
            On Error Resume Next
            l_fp = LBound(g_configSettings.FilePatternIdentifiers)
            u_fp = UBound(g_configSettings.FilePatternIdentifiers)
            If Err.Number = 0 Then
                If k >= l_fp And k <= u_fp Then
                    targetPattern = Trim(g_configSettings.FilePatternIdentifiers(k))
                Else
                    Call M04_LogWriter.WriteErrorLog("WARNING", "MainControl", "ExtractDataMain", "FilePatternIdentifiers のインデックスk=" & k & "が範囲外 (L:" & l_fp & ", U:" & u_fp & ")。空パターン使用。")
                End If
            Else
                Call M04_LogWriter.WriteErrorLog("WARNING", "MainControl", "ExtractDataMain", "FilePatternIdentifiers の LBound/UBound 取得失敗。インデックスk=" & k & "。空パターン使用。")
                Err.Clear
            End If
            On Error GoTo ErrorHandler_ExtractDataMain
        Else
            Call M04_LogWriter.WriteErrorLog("WARNING", "MainControl", "ExtractDataMain", "FilePatternIdentifiers 配列が未初期化。インデックスk=" & k & "。空パターン使用。")
        End If
        If g_configSettings.DebugDetailLevel1Enabled Then Call M04_LogWriter.WriteErrorLog("DEBUG_POINT", "MainControl", "ExtractDataMain", "targetPattern = " & targetPattern)
        g_configSettings.CurrentPatternIdentifier = targetPattern

        If targetPath = "" Then
            Call M04_LogWriter.WriteErrorLog("WARNING", "MainControl", "ExtractDataMain", "処理対象パスが空。スキップ。(インデックス: " & k & ")")
            GoTo NextIteration_ExtractDataMain
        End If

        If g_configSettings.DebugDetailLevel1Enabled Then Call M04_LogWriter.WriteErrorLog("DEBUG_POINT", "MainControl", "ExtractDataMain", "ファイル/フォルダ存在確認前: " & targetPath)
        If fileSystemObj.FolderExists(targetPath) Then
            If g_configSettings.DebugDetailLevel1Enabled Then Call M04_LogWriter.WriteErrorLog("DEBUG_POINT", "MainControl", "ExtractDataMain", "フォルダ処理開始: " & targetPath)
            Set filesInFolder = fileSystemObj.GetFolder(targetPath).Files
            If filesInFolder.Count = 0 Then
                Call M04_LogWriter.WriteErrorLog("INFORMATION", "MainControl", "ExtractDataMain", "対象フォルダにファイルが存在しません: " & targetPath)
                GoTo NextIteration_ExtractDataMain
            End If
            For Each currentFile In filesInFolder
                If g_configSettings.DebugDetailLevel1Enabled Then Call M04_LogWriter.WriteErrorLog("DEBUG_POINT", "MainControl", "ExtractDataMain", "フォルダ内ファイル処理: " & currentFile.Path)
                If IsSupportedExcelFile(currentFile.Path, fileSystemObj) Then
                    Call M06_DataExtractor.ExtractDataFromFile(currentFile.Path, g_configSettings, ThisWorkbook)
                Else
                    Call M04_LogWriter.WriteErrorLog("INFORMATION", "MainControl", "ExtractDataMain", "サポート外ファイル(フォルダ内): " & currentFile.Path)
                End If
            Next currentFile
        ElseIf fileSystemObj.FileExists(targetPath) Then
            If g_configSettings.DebugDetailLevel1Enabled Then Call M04_LogWriter.WriteErrorLog("DEBUG_POINT", "MainControl", "ExtractDataMain", "個別ファイル処理開始: " & targetPath)
             If IsSupportedExcelFile(targetPath, fileSystemObj) Then
                Call M06_DataExtractor.ExtractDataFromFile(targetPath, g_configSettings, ThisWorkbook)
            Else
                Call M04_LogWriter.WriteErrorLog("INFORMATION", "MainControl", "ExtractDataMain", "サポート外ファイル(個別): " & targetPath)
            End If
        Else
            Call M04_LogWriter.WriteErrorLog("ERROR", "MainControl", "ExtractDataMain", "指定パス見つからず: " & targetPath)
        End If
NextIteration_ExtractDataMain:
    Next k
    If g_configSettings.DebugDetailLevel1Enabled Then Call M04_LogWriter.WriteErrorLog("DEBUG_POINT", "MainControl", "ExtractDataMain", "ファイル処理ループ終了。")

    ' 5. 完了処理
    Dim endTime As Date
    endTime = Now()
    Dim timeTaken As String
    timeTaken = Format(endTime - startTime, "hh:mm:ss")

    Call M04_LogWriter.WriteErrorLog("INFORMATION", "MainControl", "ExtractDataMain", "マクロ処理が正常に完了しました。処理時間: " & timeTaken)
    MsgBox "マクロ処理が完了しました。" & vbCrLf & "処理時間: " & timeTaken, vbInformation, "処理完了"

FinalizeRoutine_ExtractDataMain:
    On Error Resume Next
    Set fileSystemObj = Nothing
    Set currentFile = Nothing
    Set filesInFolder = Nothing
    Set g_errorLogWorksheet = Nothing
    Set g_configSettings.MainWorkbookObject = Nothing
    On Error GoTo 0
    Exit Sub

ErrorHandler_ExtractDataMain:
    Dim errorMsg As String
    Dim errModule As String
    Dim errProc As String

    errModule = "MainControl"
    errProc = "ExtractDataMain (or called procedure)"

    errorMsg = "エラーが発生しました。" & vbCrLf & _
               "エラー番号: " & Err.Number & vbCrLf & _
               "エラー内容: " & Err.Description & vbCrLf & _
               "発生モジュール: " & errModule & vbCrLf & _
               "発生プロシージャ: " & errProc

    If Not g_errorLogWorksheet Is Nothing And Not g_configSettings.MainWorkbookObject Is Nothing Then
        Call M04_LogWriter.WriteErrorLog("CRITICAL", errModule, errProc, "エラー番号: " & Err.Number & " - " & Err.Description, Err.Number, Err.Description)
    Else
        Debug.Print Now & " CRITICAL ERROR in " & errModule & "." & errProc & ": " & Err.Number & " - " & Err.Description
        If g_errorLogWorksheet Is Nothing Then Debug.Print "g_errorLogWorksheet is Nothing."
        If g_configSettings.MainWorkbookObject Is Nothing Then Debug.Print "g_configSettings.MainWorkbookObject is Nothing."
    End If

    MsgBox errorMsg, vbCritical, "実行時エラー"
    Resume FinalizeRoutine_ExtractDataMain
End Sub

Private Sub InitializeConfigStructure(ByRef config As tConfigSettings)
    ' A. General Settings
    config.DebugModeFlag = False
    ' config.TraceDebugEnabled = False ' Removed
    config.EnableSheetLogging = True  ' Default for GenericLog
    config.EnableSearchConditionLogSheetOutput = True
    config.EnableErrorLogSheetOutput = True
    config.DebugDetailLevel1Enabled = True  ' Default ON
    config.DebugDetailLevel2Enabled = False ' Default OFF
    config.DebugDetailLevel3Enabled = False ' Default OFF
    config.LogSheetName = "Log"
    config.DefaultFolderPath = vbNullString
    config.OutputSheetName = "抽出結果"
    config.SearchConditionLogSheetName = "検索条件ログ"
    config.ErrorLogSheetName = "エラーログ"
    config.ConfigSheetName = CONFIG_SHEET_DEFAULT_NAME
    config.GetPatternDataMethod = True

    ' B. Schedule File Settings
    Erase config.TargetSheetNames
    config.HeaderRowCount = 0
    config.HeaderColCount = 0
    config.RowsPerDay = 0
    config.MaxDaysPerSheet = 0
    config.YearCellAddress = vbNullString
    config.MonthCellAddress = vbNullString
    config.DayColumnLetter = vbNullString
    config.DayRowOffset = 0
    config.ProcessesPerDay = 0

    ' C. Process Pattern Definition
    config.CurrentPatternIdentifier = vbNullString
    Erase config.ProcessKeys
    Erase config.Kankatsu1List
    Erase config.Kankatsu2List
    Erase config.Bunrui1List
    Erase config.Bunrui2List
    Erase config.Bunrui3List
    Erase config.ProcessColCountSheetHeaders
    Erase config.ProcessColCounts
    Erase config.ProcessDetails
    Erase config.ProcessPatternColNumbers

    ' D. Filter Conditions
    config.WorkerFilterLogic = "AND"
    Erase config.WorkerFilterList
    Erase config.Kankatsu1FilterList
    Erase config.Kankatsu2FilterList
    config.Bunrui1Filter = vbNullString
    config.Bunrui2Filter = vbNullString
    config.Bunrui3Filter = vbNullString
    Erase config.KoujiShuruiFilterList
    Erase config.KoubanFilterList
    Erase config.SagyoushuruiFilterList
    Erase config.TantouFilterList
    config.NinzuFilter = vbNullString
    config.IsNinzuFilterOriginallyEmpty = True
    config.SagyouKashoKindFilter = vbNullString
    Erase config.SagyouKashoFilterList

    ' E. Target File Definition
    Erase config.TargetFileFolderPaths
    Erase config.FilePatternIdentifiers

    ' F. Extraction Data Offset Definition
    Erase config.OffsetItemMasterNames
    Erase config.OffsetDefinitions
    Erase config.IsOffsetOriginallyEmptyFlags
    config.IsOffsetDefinitionsValid = False

    ' G. Output Sheet Settings
    config.OutputHeaderRowCount = 1
    Erase config.OutputHeaderContents
    config.OutputDataOption = "上書き"
    config.HideSheetMethod = "非表示"
    Erase config.HideSheetNames

    ' Additional Members
    config.StartTime = CDate(0)
    config.ScriptFullName = vbNullString
    config.WorkSheetName = "Work"
    config.ConfigSheetFullName = vbNullString
    Set config.MainWorkbookObject = Nothing
End Sub

Private Function IsSupportedExcelFile(ByVal filePath As String, ByVal fso As Object) As Boolean
    Dim extension As String
    extension = LCase(fso.GetExtensionName(filePath))
    Select Case extension
        Case "xls", "xlsx", "xlsm", "xlsb"
            IsSupportedExcelFile = True
        Case Else
            IsSupportedExcelFile = False
    End Select
End Function

Public Function General_IsArrayInitialized(arr As Variant) As Boolean
    If Not IsArray(arr) Then
        General_IsArrayInitialized = False
        Exit Function
    End If
    General_IsArrayInitialized = True
End Function
