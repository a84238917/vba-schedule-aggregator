' バージョン：v0.5.1
Option Explicit
' このモジュールは、設定シートから情報を読み取り、g_configSettings グローバル変数を設定する役割を担います。
' 主に LoadConfiguration 関数を通じて、M00_GlobalDeclarationsで定義された tConfigSettings 型の変数に値を設定します。

Private Const MODULE_NAME As String = "M02_ConfigReader"
Private m_errorOccurred As Boolean ' Module-level flag for LoadConfiguration

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
Public Function LoadConfiguration(ByRef configStruct As tConfigSettings, ByVal targetWorkbook As Workbook) As Boolean ' Third parameter removed
    Dim wsConfig As Worksheet
    Dim funcName As String: funcName = "LoadConfiguration"
    ' m_errorOccurred is a module-level variable

    m_errorOccurred = False ' Initialize at the start of this specific function call

    ' Configシートオブジェクト取得 (configStruct.configSheetName を使用)
    On Error Resume Next ' Specific handling for wsConfig acquisition
    Set wsConfig = targetWorkbook.Worksheets(configStruct.configSheetName)
    On Error GoTo 0 ' Reset error handling immediately after the Set statement

    If wsConfig Is Nothing Then
        ' This case should be rare if MainControl already set configStruct.configSheetName from a valid sheet.
        ' However, if it happens, log it (M04_LogWriter might not be fully ready if error log sheet itself is the issue)
        Debug.Print Now & " CRITICAL: " & MODULE_NAME & "." & funcName & " - Configシート「" & configStruct.configSheetName & "」が見つかりません。 (モジュール: M02_ConfigReader)"
        m_errorOccurred = True
        ' GoTo FinalConfigCheck_LoadConfig ' Use a specific label for cleanup within this function
        ' For now, let it fall through to the main error check block
    End If

    If Not m_errorOccurred Then ' Proceed only if wsConfig was likely set
        On Error GoTo ErrorHandler_LoadConfiguration ' General error handler for the Load... calls

        configStruct.ConfigSheetFullName = targetWorkbook.FullName & " | " & wsConfig.Name ' Moved here, only if wsConfig is valid

        Call LoadGeneralSettings(configStruct, wsConfig)
        ' After each call, check for errors that might have been raised by subs not setting m_errorOccurred
        If Err.Number <> 0 Then Call PropagateError(MODULE_NAME, funcName, "LoadGeneralSettings", Err.Number, Err.Description)
        If m_errorOccurred Then GoTo FinalConfigCheck_LoadConfig ' Check module flag

        Call LoadScheduleFileSettings(configStruct, wsConfig)
        If Err.Number <> 0 Then Call PropagateError(MODULE_NAME, funcName, "LoadScheduleFileSettings", Err.Number, Err.Description)
        If m_errorOccurred Then GoTo FinalConfigCheck_LoadConfig

        ' Temporarily commented out calls (as per previous subtask)
        ' Call LoadProcessPatternDefinition(configStruct, wsConfig)
        ' If Err.Number <> 0 Then Call PropagateError(MODULE_NAME, funcName, "LoadProcessPatternDefinition", Err.Number, Err.Description)
        ' If m_errorOccurred Then GoTo FinalConfigCheck_LoadConfig

        Call LoadFilterConditions(configStruct, wsConfig) ' Focus on this one
        If Err.Number <> 0 Then Call PropagateError(MODULE_NAME, funcName, "LoadFilterConditions", Err.Number, Err.Description)
        If m_errorOccurred Then GoTo FinalConfigCheck_LoadConfig

        ' Call LoadTargetFileDefinition(configStruct, wsConfig)
        ' If Err.Number <> 0 Then Call PropagateError(MODULE_NAME, funcName, "LoadTargetFileDefinition", Err.Number, Err.Description)
        ' If m_errorOccurred Then GoTo FinalConfigCheck_LoadConfig

        ' --- F. 抽出データオフセット定義 ---
        Dim fSectionReadLoopIdx As Long ' Moved relevant Dim statements inside "If Not m_errorOccurred"
        Dim gSectionHeaderReadLoopIdx As Long
        Dim dbgFSectionPrintIdx As Long
        Dim dbgGHeaderPrintIdx As Long
        Dim itemName As String
        Dim offsetStr As String
        Dim tempOffset As tOffset
        Dim actualOffsetCount As Long
        Dim headerCellAddress As String
        Dim rawHeaderCellVal As Variant
        Dim headerVal As String
        Dim outputOpt As String

        actualOffsetCount = 0 ' Initialize for F-Section processing

        If configStruct.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M02_ConfigReader.LoadConfiguration - Reading Section F: Extraction Data Offset Definition (Array Method)"
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
                    ' Note: ParseOffset itself sets the module-level m_errorOccurred on failure
                    If ParseOffset(offsetStr, tempOffset, m_errorOccurred, funcName & " (F-Section)", itemName & " オフセット(O" & (778 + fSectionReadLoopIdx) & ")", targetWorkbook, configStruct.ErrorLogSheetName) Then
                        configStruct.OffsetDefinitions(actualOffsetCount) = tempOffset
                    Else
                        configStruct.OffsetDefinitions(actualOffsetCount).Row = 0 ' Ensure default on parse fail
                        configStruct.OffsetDefinitions(actualOffsetCount).Col = 0
                    End If
                Else
                    configStruct.OffsetDefinitions(actualOffsetCount).Row = 0
                    configStruct.OffsetDefinitions(actualOffsetCount).Col = 0
                End If

                If configStruct.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_CONFIG_DETAIL:   F. Offset Item " & actualOffsetCount & " (" & itemName & ", N" & (778 + fSectionReadLoopIdx) & "): '" & offsetStr & "' -> R:" & configStruct.OffsetDefinitions(actualOffsetCount).Row & ", C:" & configStruct.OffsetDefinitions(actualOffsetCount).Col & ", IsEmptyOrig: " & configStruct.IsOffsetOriginallyEmptyFlags(actualOffsetCount)
                If m_errorOccurred Then GoTo FinalConfigCheck_LoadConfig ' Check module flag after ParseOffset
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

        ' Determine if F-Section was successful and set IsOffsetDefinitionsValid
        If Not m_errorOccurred Then ' No parsing errors during F-Section
            configStruct.IsOffsetDefinitionsValid = True ' Consider valid even if actualOffsetCount = 0 (arrays are ReDim'd)
        Else
            configStruct.IsOffsetDefinitionsValid = False ' Error occurred during F-Section processing
        End If
        If m_errorOccurred Then GoTo FinalConfigCheck_LoadConfig ' Re-check before G-section

        ' --- G. 出力シート設定 ---
        If configStruct.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M02_ConfigReader.LoadConfiguration - Reading Section G: Output Sheet Settings"
        configStruct.OutputHeaderRowCount = IIf(IsEmpty(wsConfig.Range("O811").Value) Or IsNull(wsConfig.Range("O811").Value), 1, CLng(wsConfig.Range("O811").Value))
        If configStruct.OutputHeaderRowCount <= 0 Then configStruct.OutputHeaderRowCount = 1

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
        ' configStruct.HideSheetNames = ReadRangeToArray(wsConfig, "O1127:O1146", MODULE_NAME, funcName, "マクロ実行後非表示シートリスト") ' Still commented out

    End If ' End of "If Not m_errorOccurred Then" for wsConfig check

FinalConfigCheck_LoadConfig:
    If m_errorOccurred Then
        If Err.Number = 0 Then ' If m_errorOccurred was set by a sub but no active error here
            Call M04_LogWriter.WriteErrorLog("CRITICAL", MODULE_NAME, funcName, "設定読み込み中にエラーが発生しました。詳細は直前のログを確認してください。")
        Else ' An error is active, log it
            Call M04_LogWriter.WriteErrorLog("CRITICAL", MODULE_NAME, funcName, "設定読み込み中にエラーが発生しました (伝播または新規)。", Err.Number, Err.Description)
        End If
        LoadConfiguration = False
    Else
        If configStruct.DebugModeFlag Then
            ' --- Final Debug Print --- (Moved inside success condition)
            Dim dbgFSectionPrintIdx_Renamed As Long ' Declare here as it's only used in this block
            Dim dbgGHeaderPrintIdx_Renamed As Long
            Debug.Print "--- Loaded Configuration Settings (M02_ConfigReader) ---"
            Debug.Print "F. IsOffsetDefinitionsValid: " & configStruct.IsOffsetDefinitionsValid
            If configStruct.IsOffsetDefinitionsValid And UBound(configStruct.OffsetItemMasterNames) >= LBound(configStruct.OffsetItemMasterNames) Then
                 For dbgFSectionPrintIdx_Renamed = LBound(configStruct.OffsetItemMasterNames) To UBound(configStruct.OffsetItemMasterNames)
                    Debug.Print "  F Item " & dbgFSectionPrintIdx_Renamed & ". Name: '" & configStruct.OffsetItemMasterNames(dbgFSectionPrintIdx_Renamed) & _
                                  "', Offset: R=" & configStruct.OffsetDefinitions(dbgFSectionPrintIdx_Renamed).Row & ", C=" & configStruct.OffsetDefinitions(dbgFSectionPrintIdx_Renamed).Col & _
                                  ", IsEmptyOrig: " & configStruct.IsOffsetOriginallyEmptyFlags(dbgFSectionPrintIdx_Renamed)
                Next dbgFSectionPrintIdx_Renamed
            ElseIf configStruct.IsOffsetDefinitionsValid Then
                Debug.Print "  F. No Offset Items Loaded."
            Else
                Debug.Print "  F. Offset Definitions are NOT valid."
            End If
            Debug.Print "G-1. OutputHeaderRowCount: " & configStruct.OutputHeaderRowCount
            If configStruct.OutputHeaderRowCount > 0 And General_IsArrayInitialized(configStruct.OutputHeaderContents) Then
                For dbgGHeaderPrintIdx_Renamed = 1 To configStruct.OutputHeaderRowCount
                     Debug.Print "  G-2. OutputHeaderContents(" & dbgGHeaderPrintIdx_Renamed & "): [" & configStruct.OutputHeaderContents(dbgGHeaderPrintIdx_Renamed) & "]"
                Next dbgGHeaderPrintIdx_Renamed
            End If
            Debug.Print "--- End of Loaded Configuration Settings ---"
        End If
        LoadConfiguration = True
    End If
    Exit Function

ErrorHandler_LoadConfiguration: ' Catches unhandled errors from Load... calls
    Call PropagateError(MODULE_NAME, funcName, "LoadConfigurationメイン処理", Err.Number, Err.Description)
    Resume FinalConfigCheck_LoadConfig
End Function

' --- Private Helper Subroutines ---
Private Sub LoadGeneralSettings(ByRef config As tConfigSettings, ByVal ws As Worksheet)
    Dim funcName As String: funcName = "LoadGeneralSettings"
    On Error Resume Next ' 特定のセルアクセスエラーをハンドルするため

    config.DebugModeFlag = ReadBoolCell(ws, "O3", MODULE_NAME, funcName, "デバッグモードフラグ")
    config.TraceDebugEnabled = ReadBoolCell(ws, "O4", MODULE_NAME, funcName, "詳細トレースデバッグ有効フラグ", True) ' Assuming O4 is TraceDebugEnabled, Default True
    config.EnableSheetLogging = ReadBoolCell(ws, "O5", MODULE_NAME, funcName, "汎用ログシートへの出力有効フラグ", True) ' Default True
    config.EnableSearchConditionLogSheetOutput = ReadBoolCell(ws, "O6", MODULE_NAME, funcName, "検索条件ログシート出力有効フラグ", True) ' ★追加, Default True
    config.EnableErrorLogSheetOutput = ReadBoolCell(ws, "O7", MODULE_NAME, funcName, "エラーログシート出力有効フラグ", True) ' ★追加, Default True
    config.DefaultFolderPath = ReadStringCell(ws, "O12", MODULE_NAME, funcName, "デフォルトフォルダパス")
    config.LogSheetName = ReadStringCell(ws, "O42", MODULE_NAME, funcName, "汎用ログシート名", "Log") ' Default to "Log" if O42 is empty
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
    Dim currentItem As String
    Dim rawData As Variant
    Dim tempList() As String
    Dim i As Long, r As Long ' Added r for row iteration
    Dim count As Long
    Dim lBound1 As Long, uBound1 As Long, lBound2 As Long, uBound2 As Long

    On Error GoTo ErrorHandler_LoadFilterConditions

    currentItem = "WorkerFilterLogic (O242)"
    config.WorkerFilterLogic = ReadStringCell(ws, "O242", MODULE_NAME, funcName, "作業員フィルター検索論理", "AND")

    currentItem = "WorkerFilterList (O243:O262)"
    rawData = ReadRangeToArray(ws, "O243:O262", MODULE_NAME, funcName, "作業員フィルターリスト")

    If IsEmpty(rawData) Then
        ReDim config.WorkerFilterList(1 To 0)
    ElseIf Not IsArray(rawData) Then
        If Trim(CStr(rawData)) <> "" Then
            ReDim config.WorkerFilterList(1 To 1)
            config.WorkerFilterList(1) = Trim(CStr(rawData))
        Else
            ReDim config.WorkerFilterList(1 To 0)
        End If
    Else ' IsArray(rawData) is True
        On Error Resume Next ' For LBound/UBound calls
        lBound1 = LBound(rawData, 1)
        uBound1 = UBound(rawData, 1)
        lBound2 = LBound(rawData, 2) ' Attempt to get 2nd dimension
        uBound2 = UBound(rawData, 2)
        If Err.Number <> 0 Then ' It's a 1D array
            Err.Clear
            If uBound1 >= lBound1 Then
                ReDim tempList(lBound1 To uBound1)
                count = 0
                For i = lBound1 To uBound1
                    If Not IsEmpty(rawData(i)) And Trim(CStr(rawData(i))) <> "" Then
                        count = count + 1
                        tempList(count) = Trim(CStr(rawData(i)))
                    End If
                Next i
                If count > 0 Then
                    ReDim Preserve tempList(1 To count)
                    config.WorkerFilterList = tempList
                Else
                    ReDim config.WorkerFilterList(1 To 0)
                End If
            Else ' Array like (1 To 0)
                ReDim config.WorkerFilterList(1 To 0)
            End If
        Else ' It's a 2D array
            If uBound1 >= lBound1 And uBound2 >= lBound2 Then
                If lBound2 = 1 And uBound2 = 1 Then ' N rows x 1 column
                    ReDim tempList(1 To uBound1 - lBound1 + 1)
                    count = 0
                    For r = lBound1 To uBound1
                        If Not IsEmpty(rawData(r, 1)) And Trim(CStr(rawData(r, 1))) <> "" Then
                            count = count + 1
                            tempList(count) = Trim(CStr(rawData(r, 1)))
                        End If
                    Next r
                    If count > 0 Then
                        ReDim Preserve tempList(1 To count)
                        config.WorkerFilterList = tempList
                    Else
                        ReDim config.WorkerFilterList(1 To 0)
                    End If
                Else ' Not a single column 2D array, treat as error for this list type
                    Call M04_LogWriter.WriteErrorLog("WARNING", MODULE_NAME, funcName, currentItem & " - 予期しない2D配列構造。リストは空になります。")
                    ReDim config.WorkerFilterList(1 To 0)
                End If
            Else ' Array like (1 To 0, 1 To 0)
                ReDim config.WorkerFilterList(1 To 0)
            End If
        End If
        On Error GoTo ErrorHandler_LoadFilterConditions ' Restore error handler
    End If
    ' --- TEMPORARILY COMMENT OUT OTHER LISTS FOR FOCUSED TESTING ---
    ' currentItem = "Kankatsu1FilterList (O275:O294)"
    ' config.Kankatsu1FilterList = ReadRangeToArray(ws, "O275:O294", MODULE_NAME, funcName, "管内1フィルターリスト")
    ' ... (and so on for all other ReadRangeToArray calls in this sub)

    ' --- Temporarily comment out string reads as well to isolate array issue ---
    ' currentItem = "Bunrui1Filter (O346)"
    ' config.Bunrui1Filter = ReadStringCell(ws, "O346", MODULE_NAME, funcName, "分類1フィルター")
    ' ... (and so on for other ReadStringCell calls) ...

    ' currentItem = "NinzuFilter (O503)"
    ' config.NinzuFilter = ReadStringCell(ws, "O503", MODULE_NAME, funcName, "人数フィルター")
    ' config.IsNinzuFilterOriginallyEmpty = (Trim(config.NinzuFilter) = "")
    ' currentItem = "SagyouKashoKindFilter (O514)"
    ' config.SagyouKashoKindFilter = ReadStringCell(ws, "O514", MODULE_NAME, funcName, "作業箇所の種類フィルター")

    Exit Sub ' Normal exit for this phase

ErrorHandler_LoadFilterConditions:
    Call M04_LogWriter.WriteErrorLog("ERROR", MODULE_NAME, funcName, "フィルター条件「" & currentItem & "」の処理中にエラー。", Err.Number, Err.Description)
    ' Propagate error by not handling it with Resume or Exit Sub
End Sub

' E. 処理対象ファイル定義 (P, Q列)
Private Sub LoadTargetFileDefinition(ByRef config As tConfigSettings, ByVal ws As Worksheet)
    Dim funcName As String: funcName = "LoadTargetFileDefinition"
    On Error Resume Next ' Keep this for the overall sub, individual reads have their own handling

    config.TargetFileFolderPaths = ReadRangeToArray(ws, "P557:P756", MODULE_NAME, funcName, "処理対象ファイル/フォルダパスリスト")
    config.FilePatternIdentifiers = ReadRangeToArray(ws, "Q557:Q756", MODULE_NAME, funcName, "各処理対象ファイル適用工程パターン識別子")

    ' Validate FilePatternIdentifiers
    Dim isValidArray As Boolean
    isValidArray = False
    If IsArray(config.FilePatternIdentifiers) Then
        On Error Resume Next ' Check LBound/UBound safely
        Dim l As Long, u As Long
        l = LBound(config.FilePatternIdentifiers)
        u = UBound(config.FilePatternIdentifiers)
        If Err.Number = 0 Then
            isValidArray = True ' LBound/UBound succeeded, it's a proper array
        Else
            Err.Clear
        End If
        On Error GoTo 0
    End If

    If Not isValidArray Then
        Call M04_LogWriter.WriteErrorLog("WARNING", MODULE_NAME, funcName, "FilePatternIdentifiers (Q557:Q756) が有効な配列として読み込めませんでした。空の配列として初期化します。")
        ReDim config.FilePatternIdentifiers(1 To 0)
    End If
    ' A similar check could be added for TargetFileFolderPaths if deemed necessary

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

Private Function ReadRangeToArray(ws As Worksheet, rangeAddress As String, moduleN As String, funcN As String, itemName As String) As Variant
    Dim data As Variant
    On Error GoTo ReadRangeErrorHandler

    data = ws.Range(rangeAddress).value
    ReadRangeToArray = data
    Exit Function

ReadRangeErrorHandler:
    Call M04_LogWriter.WriteErrorLog("ERROR", moduleN, funcN, itemName & " (" & rangeAddress & ") の範囲読み取り自体に失敗。", Err.Number, Err.Description)
    ReadRangeToArray = Empty ' Return Empty on error
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

Private Sub PropagateError(ByVal moduleN As String, ByVal callerProcName As String, ByVal failedSubName As String, ByVal errNum As Long, ByVal errDesc As String)
    ' This sub is called when a Load... sub finishes and Err.Number is not 0,
    ' meaning an error occurred in the sub and was not handled by Resume Next or Exit Sub within it.
    m_errorOccurred = True ' Set the module-level flag
    Call M04_LogWriter.WriteErrorLog("ERROR", moduleN, callerProcName, failedSubName & " からエラーが伝播 (または新規発生)。", errNum, errDesc)
    ' Do not Clear Err here, let LoadConfiguration handle it or GoTo
End Sub
