' バージョン：v0.5.1
Option Explicit
' このモジュールは、設定シートから情報を読み取り、g_configSettings グローバル変数を設定する役割を担います。
' 主に LoadConfiguration 関数を通じて、M00_GlobalDeclarationsで定義された tConfigSettings 型の変数に値を設定します。

Private Const MODULE_NAME As String = "M02_ConfigReader"

' --- Helper: ParseOffset (moved here from LoadConfiguration's F-Section in prompt for module-level visibility if needed, or keep private if only for LoadConfig)
Private Function ParseOffset(offsetString As String, ByRef resultOffset As tOffset, ByRef overallErrorFlag As Boolean, callerProcName As String, itemDesc As String, ByVal wbForLog As Workbook, ByVal errorLogSheetNameForLog As String) As Boolean
    ' オフセット文字列("行,列")を解析し、tOffset型に格納します。書式不正の場合はエラーを報告しFalseを返します。
    Dim parts() As String
    Dim strVal As String

    ParseOffset = False ' Default to failure
    resultOffset.Row = 0
    resultOffset.Col = 0

    strVal = Trim(offsetString)

    If Len(strVal) = 0 Then
        ' Empty string is not an error for ParseOffset itself, but indicates no offset.
        ' The caller decides if an empty offset is permissible.
        ParseOffset = True
        Exit Function
    End If

    parts = Split(strVal, ",")

    If UBound(parts) <> 1 Then
        Call M04_LogWriter.WriteErrorLog("ERROR", MODULE_NAME, callerProcName, itemDesc & " - オフセット書式不正 (カンマ区切り2要素でない): '" & offsetString & "'", 0, "ParseError")
        overallErrorFlag = True ' Signal error to caller
        Exit Function ' Returns False
    End If

    If Not IsNumeric(Trim(parts(0))) Or Not IsNumeric(Trim(parts(1))) Then
        Call M04_LogWriter.WriteErrorLog("ERROR", MODULE_NAME, callerProcName, itemDesc & " - オフセット値が数値でない: '" & offsetString & "'", 0, "ParseError")
        overallErrorFlag = True ' Signal error to caller
        Exit Function ' Returns False
    End If

    resultOffset.Row = CLng(Trim(parts(0)))
    resultOffset.Col = CLng(Trim(parts(1)))
    ParseOffset = True ' Successfully parsed
End Function


' --- Public Functions ---
Public Function LoadConfiguration(ByRef configStruct As tConfigSettings, ByVal targetWorkbook As Workbook, ByVal configSheetName As String) As Boolean
    Dim wsConfig As Worksheet
    Dim funcName As String: funcName = "LoadConfiguration"
    Dim m_errorOccurred As Boolean: m_errorOccurred = False ' Local error flag for this loading process

    ' Loop Counters
    Dim fSectionReadLoopIdx As Long
    Dim gSectionHeaderReadLoopIdx As Long
    Dim dbgFSectionPrintIdx As Long
    Dim dbgGHeaderPrintIdx As Long

    ' Variables for F-Section Reading
    Dim itemName As String
    Dim offsetStr As String
    Dim tempOffset As tOffset
    Dim actualOffsetCount As Long
    ' Dim currentFatalErrorState As Boolean ' This was used to check m_errorOccurred before and after ParseOffset, now ParseOffset directly modifies m_errorOccurred

    ' Variables for G-Section Reading
    Dim headerCellAddress As String ' Used for logging/debug, actual cell read is direct
    Dim rawHeaderCellVal As Variant ' For reading raw header value
    Dim headerVal As String         ' For processed header string
    Dim outputOpt As String       ' For OutputDataOption

    ' General temp variable for reading values
    Dim tempVal As Variant


    On Error GoTo ErrorHandler_LoadConfiguration

    ' Configシートオブジェクト取得
    On Error Resume Next
    Set wsConfig = targetWorkbook.Worksheets(configSheetName)
    On Error GoTo ErrorHandler_LoadConfiguration

    If wsConfig Is Nothing Then
        Call M04_LogWriter.WriteErrorLog("CRITICAL", MODULE_NAME, funcName, "Configシート「" & configSheetName & "」が見つかりません。", 0, "処理中断")
        LoadConfiguration = False
        Exit Function
    End If
    configStruct.ConfigSheetFullName = targetWorkbook.FullName & " | " & wsConfig.Name

    ' --- A. 一般設定 ---
    If configStruct.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M02_ConfigReader.LoadConfiguration - Reading Section A: General Settings"
    Call LoadGeneralSettings(configStruct, wsConfig) ' Assuming this helper and others (B,C,D,E) remain for now.

    ' --- B. 工程表ファイル内 設定 ---
    Call LoadScheduleFileSettings(configStruct, wsConfig)

    ' --- C. 工程パターン定義 ---
    Call LoadProcessPatternDefinition(configStruct, wsConfig)

    ' --- D. フィルタ条件 ---
    Call LoadFilterConditions(configStruct, wsConfig)

    ' --- E. 処理対象ファイル定義 ---
    Call LoadTargetFileDefinition(configStruct, wsConfig)


    ' --- F. 抽出データオフセット定義 ---
    If configStruct.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M02_ConfigReader.LoadConfiguration - Reading Section F: Extraction Data Offset Definition (Array Method)"
    actualOffsetCount = 0

    Erase configStruct.OffsetItemMasterNames
    Erase configStruct.OffsetDefinitions
    Erase configStruct.IsOffsetOriginallyEmptyFlags

    For fSectionReadLoopIdx = 0 To 10 ' Corresponds to N778+fSectionReadLoopIdx
        itemName = Trim(CStr(wsConfig.Range("N" & (778 + fSectionReadLoopIdx)).Value))
        offsetStr = Trim(CStr(wsConfig.Range("O" & (778 + fSectionReadLoopIdx)).Value))

        If Len(itemName) > 0 Then
            actualOffsetCount = actualOffsetCount + 1
            ReDim Preserve configStruct.OffsetItemMasterNames(1 To actualOffsetCount)
            ReDim Preserve configStruct.OffsetDefinitions(1 To actualOffsetCount)
            ReDim Preserve configStruct.IsOffsetOriginallyEmptyFlags(1 To actualOffsetCount)

            configStruct.OffsetItemMasterNames(actualOffsetCount) = itemName
            configStruct.IsOffsetOriginallyEmptyFlags(actualOffsetCount) = (Len(offsetStr) = 0)

            If Not configStruct.IsOffsetOriginallyEmptyFlags(actualOffsetCount) Then
                If ParseOffset(offsetStr, tempOffset, m_errorOccurred, funcName & " (F-Section)", itemName & " オフセット(O" & (778 + fSectionReadLoopIdx) & ")", targetWorkbook, configStruct.ErrorLogSheetName) Then
                    configStruct.OffsetDefinitions(actualOffsetCount) = tempOffset
                Else
                    configStruct.OffsetDefinitions(actualOffsetCount).Row = 0
                    configStruct.OffsetDefinitions(actualOffsetCount).Col = 0
                    ' m_errorOccurred is set by ParseOffset if parsing failed for non-empty string
                End If
            Else
                configStruct.OffsetDefinitions(actualOffsetCount).Row = 0
                configStruct.OffsetDefinitions(actualOffsetCount).Col = 0
            End If

            If configStruct.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_CONFIG_DETAIL:   F. Offset Item " & actualOffsetCount & " (" & itemName & ", N" & (778 + fSectionReadLoopIdx) & "): '" & offsetStr & "' -> R:" & configStruct.OffsetDefinitions(actualOffsetCount).Row & ", C:" & configStruct.OffsetDefinitions(actualOffsetCount).Col & ", IsEmptyOrig: " & configStruct.IsOffsetOriginallyEmptyFlags(actualOffsetCount)
            If m_errorOccurred Then GoTo FinalConfigCheck ' Error during ParseOffset should lead to exit
        Else
            If Len(offsetStr) > 0 And configStruct.TraceDebugEnabled Then
                Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE: M02_ConfigReader.LoadConfiguration - Offset string '" & offsetStr & "' found in O" & (778 + fSectionReadLoopIdx) & " but no item name in N" & (778 + fSectionReadLoopIdx) & ". Skipping this offset entry."
            End If
        End If
    Next fSectionReadLoopIdx

    If actualOffsetCount = 0 Then
        If configStruct.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE: M02_ConfigReader.LoadConfiguration - No offset items defined with names in N778:N788. Initializing offset arrays as empty."
        ReDim configStruct.OffsetItemMasterNames(1 To 0)
        ReDim configStruct.OffsetDefinitions(1 To 0)
        ReDim configStruct.IsOffsetOriginallyEmptyFlags(1 To 0)
    End If

FinalConfigCheck: ' Label for potential GoTo from F-Section if error occurs
    If m_errorOccurred Then
        LoadConfiguration = False
        Exit Function
    End If

    ' --- G. 出力シート設定 ---
    If configStruct.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M02_ConfigReader.LoadConfiguration - Reading Section G: Output Sheet Settings"
    configStruct.OutputHeaderRowCount = IIf(IsEmpty(wsConfig.Range("O811").Value) Or IsNull(wsConfig.Range("O811").Value), 1, CLng(wsConfig.Range("O811").Value)) ' Default to 1 if empty
    If configStruct.OutputHeaderRowCount <= 0 Then configStruct.OutputHeaderRowCount = 1 ' Ensure at least 1

    If configStruct.OutputHeaderRowCount > 0 Then
        ReDim configStruct.OutputHeaderContents(1 To configStruct.OutputHeaderRowCount)
        For gSectionHeaderReadLoopIdx = 1 To configStruct.OutputHeaderRowCount
            headerCellAddress = "O" & (811 + gSectionHeaderReadLoopIdx)
            rawHeaderCellVal = wsConfig.Range(headerCellAddress).Value
            If IsError(rawHeaderCellVal) Then
                headerVal = ""
                 Call M04_LogWriter.WriteErrorLog("WARNING", MODULE_NAME, funcName, "ヘッダー内容セル (" & headerCellAddress & ") がエラー値「" & CStr(rawHeaderCellVal) & "」。空文字として扱います。", 0, configStruct.ConfigSheetFullName)
            Else
                headerVal = Trim(CStr(rawHeaderCellVal))
            End If
            configStruct.OutputHeaderContents(gSectionHeaderReadLoopIdx) = headerVal
        Next gSectionHeaderReadLoopIdx
    End If

    outputOpt = UCase(Trim(CStr(wsConfig.Range("O1124").Value)))
    If outputOpt = "リセット" Or outputOpt = "追記" Then
        configStruct.OutputDataOption = outputOpt
    Else
        configStruct.OutputDataOption = "リセット" ' Default
    End If
    configStruct.HideSheetMethod = Trim(CStr(wsConfig.Range("O1126").Value))
    configStruct.HideSheetNames = ReadRangeToArray(wsConfig, "O1127:O1146", MODULE_NAME, funcName, "マクロ実行後非表示シートリスト")


    ' --- Final Debug Print ---
    If configStruct.DebugModeFlag Then
        Debug.Print "--- Loaded Configuration Settings (M02_ConfigReader) ---"
        ' ... (A, B, C, D, E sections) ...
        Debug.Print "F. Extraction Data Offsets (Loaded " & IIf(UBound(configStruct.OffsetItemMasterNames) >= LBound(configStruct.OffsetItemMasterNames), UBound(configStruct.OffsetItemMasterNames), 0) & " items):"
        If UBound(configStruct.OffsetItemMasterNames) >= LBound(configStruct.OffsetItemMasterNames) Then
            For dbgFSectionPrintIdx = LBound(configStruct.OffsetItemMasterNames) To UBound(configStruct.OffsetItemMasterNames)
                Debug.Print "  " & dbgFSectionPrintIdx & ". Name: '" & configStruct.OffsetItemMasterNames(dbgFSectionPrintIdx) & _
                              "', Offset: R=" & configStruct.OffsetDefinitions(dbgFSectionPrintIdx).Row & ", C=" & configStruct.OffsetDefinitions(dbgFSectionPrintIdx).Col & _
                              ", IsEmptyOrig: " & configStruct.IsOffsetOriginallyEmptyFlags(dbgFSectionPrintIdx)
            Next dbgFSectionPrintIdx
        End If
        Debug.Print "G-1. OutputHeaderRowCount: " & configStruct.OutputHeaderRowCount
        If configStruct.OutputHeaderRowCount > 0 And General_IsArrayInitialized(configStruct.OutputHeaderContents) Then
            For dbgGHeaderPrintIdx = 1 To configStruct.OutputHeaderRowCount
                 Debug.Print "  G-2. OutputHeaderContents(" & dbgGHeaderPrintIdx & "): [" & configStruct.OutputHeaderContents(dbgGHeaderPrintIdx) & "]"
            Next dbgGHeaderPrintIdx
        End If
        ' ... (other G section items) ...
        Debug.Print "--- End of Loaded Configuration Settings ---"
    End If

    LoadConfiguration = True
    Exit Function

ErrorHandler_LoadConfiguration:
    Call M04_LogWriter.WriteErrorLog("CRITICAL", MODULE_NAME, funcName, "設定読み込み中に予期せぬエラーが発生しました。", Err.Number, Err.Description)
    LoadConfiguration = False
End Function


' --- Private Helper Subroutines ---
' LoadGeneralSettings, LoadScheduleFileSettings, LoadProcessPatternDefinition, LoadFilterConditions, LoadTargetFileDefinition remain
' GetSpecificOffsetFromString is now replaced by ParseOffset (defined at module level or passed in if needed by helpers)
' ReadRangeToArray, ReadStringCell, ReadLongCell, ReadBoolCell, General_IsArrayInitialized remain

' A. 一般設定 (O列)
Private Sub LoadGeneralSettings(ByRef config As tConfigSettings, ByVal ws As Worksheet)
    Dim funcName As String: funcName = "LoadGeneralSettings"
    On Error Resume Next ' 特定のセルアクセスエラーをハンドルするため

    config.DebugModeFlag = ReadBoolCell(ws, "O3", MODULE_NAME, funcName, "デバッグモードフラグ")
    config.DefaultFolderPath = ReadStringCell(ws, "O12", MODULE_NAME, funcName, "デフォルトフォルダパス")
    config.OutputSheetName = ReadStringCell(ws, "O43", MODULE_NAME, funcName, "抽出結果出力シート名", "抽出結果")
    config.SearchConditionLogSheetName = ReadStringCell(ws, "O44", MODULE_NAME, funcName, "検索条件ログシート名", "検索条件ログ")
    config.ErrorLogSheetName = ReadStringCell(ws, "O45", MODULE_NAME, funcName, "エラーログシート名", "エラーログ")
    config.ConfigSheetName = ReadStringCell(ws, "O46", MODULE_NAME, funcName, "設定ファイルシート名", CONFIG_SHEET_DEFAULT_NAME)
    config.GetPatternDataMethod = ReadBoolCell(ws, "O122", MODULE_NAME, funcName, "工程パターンデータ取得方法")

    If Err.Number <> 0 Then Call M04_LogWriter.WriteErrorLog("WARNING", MODULE_NAME, funcName, "一般設定の読み込み中にエラーが発生しました。", Err.Number, Err.Description)
    On Error GoTo 0
End Sub

' B. 工程表ファイル設定 (O列)
Private Sub LoadScheduleFileSettings(ByRef config As tConfigSettings, ByVal ws As Worksheet)
    Dim funcName As String: funcName = "LoadScheduleFileSettings"
    On Error Resume Next

    config.TargetSheetNames = ReadRangeToArray(ws, "O66:O75", MODULE_NAME, funcName, "工程表内 検索対象シート名リスト")
    config.HeaderRowCount = ReadLongCell(ws, "O87", MODULE_NAME, funcName, "工程表ヘッダー行数")
    config.HeaderColCount = ReadLongCell(ws, "O88", MODULE_NAME, funcName, "工程表ヘッダー列数")
    config.RowsPerDay = ReadLongCell(ws, "O89", MODULE_NAME, funcName, "1日のデータが占める行数")
    config.MaxDaysPerSheet = ReadLongCell(ws, "O90", MODULE_NAME, funcName, "1シート内の最大日数")
    config.YearCellAddress = ReadStringCell(ws, "O101", MODULE_NAME, funcName, "「年」のセルアドレス")
    config.MonthCellAddress = ReadStringCell(ws, "O102", MODULE_NAME, funcName, "「月」のセルアドレス")
    config.DayColumnLetter = ReadStringCell(ws, "O103", MODULE_NAME, funcName, "「日」の値がある列文字")
    config.DayRowOffset = ReadLongCell(ws, "O104", MODULE_NAME, funcName, "「日」の値の行オフセット")
    config.ProcessesPerDay = ReadLongCell(ws, "O114", MODULE_NAME, funcName, "1日の工程数")

    If Err.Number <> 0 Then Call M04_LogWriter.WriteErrorLog("WARNING", MODULE_NAME, funcName, "工程表ファイル設定の読み込み中にエラー。", Err.Number, Err.Description)
    On Error GoTo 0
End Sub

' C. 工程パターン定義 (I,J,K,L,M,N列, O-X列)
Private Sub LoadProcessPatternDefinition(ByRef config As tConfigSettings, ByVal ws As Worksheet)
    Dim funcName As String: funcName = "LoadProcessPatternDefinition"
    Dim procPtn_i As Long
    Dim numProcesses As Long
    On Error Resume Next

    config.CurrentPatternIdentifier = ReadStringCell(ws, "O126", MODULE_NAME, funcName, "現在処理中ファイル適用工程パターン識別子")
    numProcesses = config.ProcessesPerDay
    If numProcesses <= 0 Then numProcesses = 10

    config.ProcessKeys = ReadRangeToArray(ws, "I129:I" & (128 + numProcesses), MODULE_NAME, funcName, "工程キーリスト")
    config.Kankatsu1List = ReadRangeToArray(ws, "J129:J" & (128 + numProcesses), MODULE_NAME, funcName, "管内1リスト")
    config.Kankatsu2List = ReadRangeToArray(ws, "K129:K" & (128 + numProcesses), MODULE_NAME, funcName, "管内2リスト")
    config.Bunrui1List = ReadRangeToArray(ws, "L129:L" & (128 + numProcesses), MODULE_NAME, funcName, "分類1リスト")
    config.Bunrui2List = ReadRangeToArray(ws, "M129:M" & (128 + numProcesses), MODULE_NAME, funcName, "分類2リスト")
    config.Bunrui3List = ReadRangeToArray(ws, "N129:N" & (128 + numProcesses), MODULE_NAME, funcName, "分類3リスト")

    Dim headerData As Variant
    headerData = ws.Range("O128:X128").Value
    If IsArray(headerData) Then
        ReDim config.ProcessColCountSheetHeaders(1 To UBound(headerData, 2))
        For procPtn_i = 1 To UBound(headerData, 2)
            config.ProcessColCountSheetHeaders(procPtn_i) = Trim(CStr(headerData(1, procPtn_i)))
        Next procPtn_i
    End If

    config.ProcessColCounts = ws.Range("O129:X" & (128 + numProcesses)).Value

    If General_IsArrayInitialized(config.Kankatsu1List) And General_IsArrayInitialized(config.Kankatsu2List) Then
        Dim k1Count As Long, k2Count As Long, maxCount As Long
        On Error Resume Next
        k1Count = UBound(config.Kankatsu1List)
        k2Count = UBound(config.Kankatsu2List)
        If Err.Number <> 0 Then Err.Clear Else maxCount = IIf(k1Count > k2Count, k1Count, k2Count)
        On Error GoTo 0

        If maxCount > 0 Then
            ReDim config.ProcessDetails(1 To maxCount)
            For procPtn_i = 1 To maxCount
                If procPtn_i <= k1Count Then config.ProcessDetails(procPtn_i).Kankatsu1 = config.Kankatsu1List(procPtn_i)
                If procPtn_i <= k2Count Then config.ProcessDetails(procPtn_i).Kankatsu2 = config.Kankatsu2List(procPtn_i)
            Next procPtn_i
        End If
    End If

    If Err.Number <> 0 Then Call M04_LogWriter.WriteErrorLog("WARNING", MODULE_NAME, funcName, "工程パターン定義の読み込み中にエラー。", Err.Number, Err.Description)
    On Error GoTo 0
End Sub

' D. フィルタ条件 (O列)
Private Sub LoadFilterConditions(ByRef config As tConfigSettings, ByVal ws As Worksheet)
    Dim funcName As String: funcName = "LoadFilterConditions"
    On Error Resume Next

    config.WorkerFilterLogic = ReadStringCell(ws, "O242", MODULE_NAME, funcName, "作業員フィルター検索論理", "AND")
    config.WorkerFilterList = ReadRangeToArray(ws, "O243:O262", MODULE_NAME, funcName, "作業員フィルターリスト")
    config.Kankatsu1FilterList = ReadRangeToArray(ws, "O275:O294", MODULE_NAME, funcName, "管内1フィルターリスト")
    config.Kankatsu2FilterList = ReadRangeToArray(ws, "O305:O334", MODULE_NAME, funcName, "管内2フィルターリスト")
    config.Bunrui1Filter = ReadStringCell(ws, "O346", MODULE_NAME, funcName, "分類1フィルター")
    config.Bunrui2Filter = ReadStringCell(ws, "O367", MODULE_NAME, funcName, "分類2フィルター")
    config.Bunrui3Filter = ReadStringCell(ws, "O388", MODULE_NAME, funcName, "分類3フィルター")
    config.KoujiShuruiFilterList = ReadRangeToArray(ws, "O409:O418", MODULE_NAME, funcName, "工事種類フィルターリスト")
    config.KoubanFilterList = ReadRangeToArray(ws, "O431:O440", MODULE_NAME, funcName, "工番フィルターリスト")
    config.SagyoushuruiFilterList = ReadRangeToArray(ws, "O451:O470", MODULE_NAME, funcName, "作業種類フィルターリスト")
    config.TantouFilterList = ReadRangeToArray(ws, "O481:O490", MODULE_NAME, funcName, "担当の名前フィルターリスト")
    config.NinzuFilter = ReadStringCell(ws, "O503", MODULE_NAME, funcName, "人数フィルター")
    config.IsNinzuFilterOriginallyEmpty = (Trim(config.NinzuFilter) = "")
    config.SagyouKashoKindFilter = ReadStringCell(ws, "O514", MODULE_NAME, funcName, "作業箇所の種類フィルター")
    config.SagyouKashoFilterList = ReadRangeToArray(ws, "O525:O544", MODULE_NAME, funcName, "作業箇所フィルターリスト")

    If Err.Number <> 0 Then Call M04_LogWriter.WriteErrorLog("WARNING", MODULE_NAME, funcName, "フィルタ条件の読み込み中にエラー。", Err.Number, Err.Description)
    On Error GoTo 0
End Sub

' E. 処理対象ファイル定義 (P, Q列)
Private Sub LoadTargetFileDefinition(ByRef config As tConfigSettings, ByVal ws As Worksheet)
    Dim funcName As String: funcName = "LoadTargetFileDefinition"
    On Error Resume Next

    config.TargetFileFolderPaths = ReadRangeToArray(ws, "P557:P756", MODULE_NAME, funcName, "処理対象ファイル/フォルダパスリスト")
    config.FilePatternIdentifiers = ReadRangeToArray(ws, "Q557:Q756", MODULE_NAME, funcName, "各処理対象ファイル適用工程パターン識別子")

    If Err.Number <> 0 Then Call M04_LogWriter.WriteErrorLog("WARNING", MODULE_NAME, funcName, "処理対象ファイル定義の読み込み中にエラー。", Err.Number, Err.Description)
    On Error GoTo 0
End Sub

' --- Reading Helper Functions ---
Private Function ReadStringCell(ws As Worksheet, addr As String, moduleN As String, funcN As String, itemName As String, Optional defaultValue As String = vbNullString) As String
    Dim val As Variant
    On Error Resume Next
    val = ws.Range(addr).Value
    If Err.Number <> 0 Then
        ReadStringCell = defaultValue
        Call M04_LogWriter.WriteErrorLog("WARNING", moduleN, funcN, itemName & " (" & addr & ") 読み取り失敗。デフォルト「" & defaultValue & "」使用。", Err.Number, Err.Description)
    Else
        If IsEmpty(val) Or Trim(CStr(val)) = "" Then
            ReadStringCell = defaultValue
        Else
            ReadStringCell = Trim(CStr(val))
        End If
    End If
    On Error GoTo 0
End Function

Private Function ReadLongCell(ws As Worksheet, addr As String, moduleN As String, funcN As String, itemName As String, Optional defaultValue As Long = 0) As Long
    Dim val As Variant
    On Error Resume Next
    val = ws.Range(addr).Value
    If Err.Number <> 0 Then
        ReadLongCell = defaultValue
        Call M04_LogWriter.WriteErrorLog("WARNING", moduleN, funcN, itemName & " (" & addr & ") 読み取り失敗。デフォルト「" & defaultValue & "」使用。", Err.Number, Err.Description)
    Else
        If IsEmpty(val) Or Not IsNumeric(val) Then
            ReadLongCell = defaultValue
            If Not IsEmpty(val) Then Call M04_LogWriter.WriteErrorLog("WARNING", moduleN, funcN, itemName & " (" & addr & ") が数値でない。デフォルト「" & defaultValue & "」使用。")
        Else
            ReadLongCell = CLng(val)
        End If
    End If
    On Error GoTo 0
End Function

Private Function ReadBoolCell(ws As Worksheet, addr As String, moduleN As String, funcN As String, itemName As String, Optional defaultValue As Boolean = False) As Boolean
    Dim val As Variant
    On Error Resume Next
    val = ws.Range(addr).Value
    If Err.Number <> 0 Then
        ReadBoolCell = defaultValue
        Call M04_LogWriter.WriteErrorLog("WARNING", moduleN, funcN, itemName & " (" & addr & ") 読み取り失敗。デフォルト「" & defaultValue & "」使用。", Err.Number, Err.Description)
    Else
        If IsEmpty(val) Then
            ReadBoolCell = defaultValue
        Else
            ReadBoolCell = (UCase(Trim(CStr(val))) = "TRUE")
        End If
    End If
    On Error GoTo 0
End Function

Private Function ReadRangeToArray(ws As Worksheet, rangeAddress As String, moduleN As String, funcN As String, itemName As String) As String()
    Dim data As Variant, result() As String, arrRead_i As Long, nonEmptyCount As Long
    On Error Resume Next
    data = ws.Range(rangeAddress).Value
    If Err.Number <> 0 Then
        Call M04_LogWriter.WriteErrorLog("WARNING", moduleN, funcN, itemName & " (" & rangeAddress & ") 範囲読み取り失敗。", Err.Number, Err.Description)
        Exit Function
    End If
    On Error GoTo 0

    If IsArray(data) Then
        ReDim result(1 To UBound(data, 1))
        For arrRead_i = 1 To UBound(data, 1)
            If Not IsEmpty(data(arrRead_i, 1)) And Trim(CStr(data(arrRead_i, 1))) <> "" Then
                result(arrRead_i) = Trim(CStr(data(arrRead_i, 1)))
                nonEmptyCount = nonEmptyCount + 1
            Else
                result(arrRead_i) = vbNullString
            End If
        Next arrRead_i
        If nonEmptyCount = 0 Then Erase result
    Else
        If Not IsEmpty(data) And Trim(CStr(data)) <> "" Then
            ReDim result(1 To 1): result(1) = Trim(CStr(data))
        End If
    End If
    ReadRangeToArray = result
End Function

Public Function General_IsArrayInitialized(arr As Variant) As Boolean
    If Not IsArray(arr) Then
        General_IsArrayInitialized = False
        Exit Function
    End If

    ' 配列であれば、ReDimされているとみなし、初期化済みとする
    ' LBoundやUBoundのチェックは、要素が存在するかどうかの判断であり、
    ' 配列が「初期化されているか（DimやReDimされたか）」の判断とは異なる場合がある。
    ' 特にユーザー定義型の配列の場合、LBound等がエラーになることがあるため、
    ' IsArray(arr) が True であれば、ここでは初期化済みと判断する。
    General_IsArrayInitialized = True

    ' もし「要素が実際に存在するか」を確認したい場合は、別途 UBound(arr) >= LBound(arr) のようなチェックを行う。
    ' ここでは「配列として使える状態か」を返すことに注力する。
End Function
