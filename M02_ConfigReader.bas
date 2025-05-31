' バージョン：v0.5.2
Option Explicit
' このモジュールは、設定シートから情報を読み取り、g_configSettings グローバル変数を設定する役割を担います。
' 主に LoadConfiguration 関数を通じて、M00_GlobalDeclarationsで定義された tConfigSettings 型の変数に値を設定します。

Private Const MODULE_NAME As String = "M02_ConfigReader"
Private m_errorOccurred As Boolean ' Module-level flag for LoadConfiguration

' --- Helper: ParseOffset
Private Function ParseOffset(offsetString As String, ByRef resultOffset As tOffset, ByRef overallErrorFlag As Boolean, callerProcName As String, itemDesc As String, ByVal wbForLog As Workbook, ByVal errorLogSheetNameForLog As String) As Boolean
    Dim parts() As String
    Dim strVal As String
    ParseOffset = False
    resultOffset.Row = 0
    resultOffset.Col = 0
    strVal = Trim(offsetString)
    If Len(strVal) = 0 Then
        ParseOffset = True
        Exit Function
    End If
    parts = Split(strVal, ",")
    If UBound(parts) <> 1 Then
        Call M04_LogWriter.WriteErrorLog("ERROR", MODULE_NAME, callerProcName, itemDesc & " - オフセット書式不正 (カンマ区切り2要素でない): '" & offsetString & "'", 0, "ParseError")
        overallErrorFlag = True
        Exit Function
    End If
    If Not IsNumeric(Trim(parts(0))) Or Not IsNumeric(Trim(parts(1))) Then
        Call M04_LogWriter.WriteErrorLog("ERROR", MODULE_NAME, callerProcName, itemDesc & " - オフセット値が数値でない: '" & offsetString & "'", 0, "ParseError")
        overallErrorFlag = True
        Exit Function
    End If
    resultOffset.Row = CLng(Trim(parts(0)))
    resultOffset.Col = CLng(Trim(parts(1)))
    ParseOffset = True
End Function

' --- Public Functions ---
Public Function LoadConfiguration(ByRef configStruct As tConfigSettings, ByVal targetWorkbook As Workbook) As Boolean
    Dim wsConfig As Worksheet
    Dim funcName As String: funcName = "LoadConfiguration"

    m_errorOccurred = False ' Initialize module-level flag

    On Error Resume Next
    Set wsConfig = targetWorkbook.Worksheets(configStruct.configSheetName)
    On Error GoTo 0

    If wsConfig Is Nothing Then
        Debug.Print Now & " CRITICAL: " & MODULE_NAME & "." & funcName & " - Configシート「" & configStruct.configSheetName & "」が見つかりません。"
        m_errorOccurred = True
    End If

    If Not m_errorOccurred Then
        On Error GoTo ErrorHandler_LoadConfiguration
        configStruct.ConfigSheetFullName = targetWorkbook.FullName & " | " & wsConfig.Name

        Call LoadGeneralSettings(configStruct, wsConfig)
        If Err.Number <> 0 Then Call HandleConfigLoadingError(MODULE_NAME, funcName, "LoadGeneralSettings", Err.Number, Err.Description)
        If m_errorOccurred Then GoTo FinalConfigCheck_LoadConfig

        Call LoadScheduleFileSettings(configStruct, wsConfig)
        If Err.Number <> 0 Then Call HandleConfigLoadingError(MODULE_NAME, funcName, "LoadScheduleFileSettings", Err.Number, Err.Description)
        If m_errorOccurred Then GoTo FinalConfigCheck_LoadConfig

        Call LoadProcessPatternDefinition(configStruct, wsConfig)
        If Err.Number <> 0 Then Call HandleConfigLoadingError(MODULE_NAME, funcName, "LoadProcessPatternDefinition", Err.Number, Err.Description)
        If m_errorOccurred Then GoTo FinalConfigCheck_LoadConfig

        Call LoadFilterConditions(configStruct, wsConfig)
        If Err.Number <> 0 Then Call HandleConfigLoadingError(MODULE_NAME, funcName, "LoadFilterConditions", Err.Number, Err.Description)
        If m_errorOccurred Then GoTo FinalConfigCheck_LoadConfig

        Call LoadTargetFileDefinition(configStruct, wsConfig)
        If Err.Number <> 0 Then Call HandleConfigLoadingError(MODULE_NAME, funcName, "LoadTargetFileDefinition", Err.Number, Err.Description)
        If m_errorOccurred Then GoTo FinalConfigCheck_LoadConfig

        ' --- F. 抽出データオフセット定義 ---
        Dim fSectionReadLoopIdx As Long
        Dim itemName As String
        Dim offsetStr As String
        Dim tempOffset As tOffset
        Dim actualOffsetCount As Long
        actualOffsetCount = 0

        If configStruct.DebugDetailLevel2Enabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL_L2: M02_ConfigReader.LoadConfiguration - Reading Section F"
        Erase configStruct.OffsetItemMasterNames
        Erase configStruct.OffsetDefinitions
        Erase configStruct.IsOffsetOriginallyEmptyFlags
        For fSectionReadLoopIdx = 0 To 10
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
                    If ParseOffset(offsetStr, tempOffset, m_errorOccurred, funcName & " (F-Section)", itemName, targetWorkbook, configStruct.ErrorLogSheetName) Then
                        configStruct.OffsetDefinitions(actualOffsetCount) = tempOffset
                    Else
                        configStruct.OffsetDefinitions(actualOffsetCount).Row = 0
                        configStruct.OffsetDefinitions(actualOffsetCount).Col = 0
                    End If
                Else
                    configStruct.OffsetDefinitions(actualOffsetCount).Row = 0
                    configStruct.OffsetDefinitions(actualOffsetCount).Col = 0
                End If
                If configStruct.DebugDetailLevel2Enabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_CONFIG_DETAIL_L2:   F. Offset Item " & actualOffsetCount & " (" & itemName & ", N" & (778 + fSectionReadLoopIdx) & "): '" & offsetStr & "' -> R:" & configStruct.OffsetDefinitions(actualOffsetCount).Row & ", C:" & configStruct.OffsetDefinitions(actualOffsetCount).Col & ", IsEmptyOrig: " & configStruct.IsOffsetOriginallyEmptyFlags(actualOffsetCount)
                If m_errorOccurred Then GoTo FinalConfigCheck_LoadConfig
            Else
                If Len(offsetStr) > 0 And configStruct.DebugDetailLevel2Enabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE_L2: Offset string in O but no item name in N" & (778 + fSectionReadLoopIdx)
            End If
        Next fSectionReadLoopIdx
        If actualOffsetCount = 0 Then
            ReDim configStruct.OffsetItemMasterNames(1 To 0)
            ReDim configStruct.OffsetDefinitions(1 To 0)
            ReDim configStruct.IsOffsetOriginallyEmptyFlags(1 To 0)
        End If
        If Not m_errorOccurred Then configStruct.IsOffsetDefinitionsValid = True Else configStruct.IsOffsetDefinitionsValid = False
        If m_errorOccurred Then GoTo FinalConfigCheck_LoadConfig

        ' --- G. 出力シート設定 ---
        Dim gSectionHeaderReadLoopIdx As Long
        Dim headerCellAddress As String
        Dim rawHeaderCellVal As Variant
        Dim headerVal As String
        Dim outputOpt As String
        Dim rawHideSheetNames As Variant

        If configStruct.DebugDetailLevel2Enabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL_L2: M02_ConfigReader.LoadConfiguration - Reading Section G"
        configStruct.OutputHeaderRowCount = IIf(IsEmpty(wsConfig.Range("O811").value) Or IsNull(wsConfig.Range("O811").value), 1, CLng(wsConfig.Range("O811").value))
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
        If outputOpt = "リセット" Or outputOpt = "追記" Then configStruct.OutputDataOption = outputOpt Else configStruct.OutputDataOption = "リセット"
        configStruct.HideSheetMethod = Trim(CStr(wsConfig.Range("O1126").Value))

        Dim currentItemContext As String
        currentItemContext = "HideSheetNames (O1127:O1146)"
        rawHideSheetNames = ReadRangeToArray(wsConfig, "O1127:O1146", MODULE_NAME, funcName, currentItemContext)
        configStruct.HideSheetNames = ConvertRawVariantToStringArray(rawHideSheetNames, MODULE_NAME, funcName, currentItemContext, configStruct)
        Call DebugPrintArrayState(configStruct.HideSheetNames, currentItemContext, configStruct)

    End If

FinalConfigCheck_LoadConfig:
    If m_errorOccurred Then
        If Err.Number = 0 Then Call M04_LogWriter.WriteErrorLog("CRITICAL", MODULE_NAME, funcName, "設定読み込み中にエラーが発生しました。詳細は直前のログを確認してください。") Else Call M04_LogWriter.WriteErrorLog("CRITICAL", MODULE_NAME, funcName, "設定読み込み中にエラーが発生しました (伝播または新規)。", Err.Number, Err.Description)
        LoadConfiguration = False
    Else
        If configStruct.DebugDetailLevel3Enabled Then
            Dim dbgFSectionPrintIdx As Long, dbgGHeaderPrintIdx As Long
            Debug.Print "--- Loaded Configuration Settings (M02_ConfigReader) ---"

            Debug.Print "F. IsOffsetDefinitionsValid: " & configStruct.IsOffsetDefinitionsValid
            Dim hasFSectionItemsToPrint As Boolean
            hasFSectionItemsToPrint = False
            If IsArray(configStruct.OffsetItemMasterNames) Then
                On Error Resume Next
                Dim lboundTest As Long, uboundTest As Long
                lboundTest = LBound(configStruct.OffsetItemMasterNames)
                uboundTest = UBound(configStruct.OffsetItemMasterNames)
                If Err.Number = 0 Then
                    If lboundTest <= uboundTest Then
                        If Not (lboundTest = 1 And uboundTest = 0) Then
                            hasFSectionItemsToPrint = True
                        End If
                    End If
                End If
                On Error GoTo 0
            End If

            If hasFSectionItemsToPrint Then
                Debug.Print "  F. Extraction Data Offsets (Loaded " & (UBound(configStruct.OffsetItemMasterNames) - LBound(configStruct.OffsetItemMasterNames) + 1) & " potential items):"
                For dbgFSectionPrintIdx = LBound(configStruct.OffsetItemMasterNames) To UBound(configStruct.OffsetItemMasterNames)
                    If dbgFSectionPrintIdx <= UBound(configStruct.OffsetDefinitions) And dbgFSectionPrintIdx <= UBound(configStruct.IsOffsetOriginallyEmptyFlags) And _
                       dbgFSectionPrintIdx >= LBound(configStruct.OffsetDefinitions) And dbgFSectionPrintIdx >= LBound(configStruct.IsOffsetOriginallyEmptyFlags) Then
                        Debug.Print "    Item " & dbgFSectionPrintIdx & ". Name: '" & configStruct.OffsetItemMasterNames(dbgFSectionPrintIdx) & _
                                      "', Offset: R=" & configStruct.OffsetDefinitions(dbgFSectionPrintIdx).Row & ", C=" & configStruct.OffsetDefinitions(dbgFSectionPrintIdx).Col & _
                                      ", IsEmptyOrig: " & configStruct.IsOffsetOriginallyEmptyFlags(dbgFSectionPrintIdx)
                    Else
                        Debug.Print "    Item " & dbgFSectionPrintIdx & ". Name: '" & configStruct.OffsetItemMasterNames(dbgFSectionPrintIdx) & "' - (Offset data array bounds mismatch)"
                    End If
                Next dbgFSectionPrintIdx
            Else
                If configStruct.IsOffsetDefinitionsValid Then
                    Debug.Print "  F. No Offset Items Loaded (OffsetItemMasterNames is empty or not properly dimensioned, but OffsetDefinitions structure was marked valid)."
                Else
                    Debug.Print "  F. No Offset Items Loaded (OffsetItemMasterNames is empty or not properly dimensioned, and OffsetDefinitions structure was NOT marked valid or no items defined)."
                End If
            End If

            Debug.Print "G-1. OutputHeaderRowCount: " & configStruct.OutputHeaderRowCount
            If configStruct.OutputHeaderRowCount > 0 And General_IsArrayInitialized(configStruct.OutputHeaderContents) Then
                For dbgGHeaderPrintIdx = 1 To configStruct.OutputHeaderRowCount
                     Debug.Print "  G-2. OutputHeaderContents(" & dbgGHeaderPrintIdx & "): [" & configStruct.OutputHeaderContents(dbgGHeaderPrintIdx) & "]"
                Next dbgGHeaderPrintIdx
            End If
            Debug.Print "--- End of Loaded Configuration Settings ---"
        End If
        LoadConfiguration = True
    End If
    Exit Function
ErrorHandler_LoadConfiguration:
    Call HandleConfigLoadingError(MODULE_NAME, funcName, "LoadConfigurationメイン処理", Err.Number, Err.Description) ' Updated Call
    Resume FinalConfigCheck_LoadConfig
End Function

Private Sub HandleConfigLoadingError(ByVal moduleN As String, ByVal callerProcName As String, ByVal failedSubName As String, ByVal errNum As Long, ByVal errDesc As String) ' Renamed Sub
    m_errorOccurred = True
    Call M04_LogWriter.WriteErrorLog("ERROR", moduleN, callerProcName, failedSubName & " からエラーが伝播 (または新規発生)。", errNum, errDesc)
End Sub

' --- Private Helper Subroutines ---
Private Sub LoadGeneralSettings(ByRef config As tConfigSettings, ByVal ws As Worksheet)
    Dim funcName As String: funcName = "LoadGeneralSettings"
    Dim currentItem As String
    On Error GoTo ErrorHandler_LoadGeneralSettings

    currentItem = "DebugModeFlag (O3)"
    config.DebugModeFlag = ReadBoolCell(ws, "O3", MODULE_NAME, funcName, currentItem)
    currentItem = "DebugDetailLevel1Enabled (O4)"
    config.DebugDetailLevel1Enabled = ReadBoolCell(ws, "O4", MODULE_NAME, funcName, currentItem, True)
    currentItem = "EnableSheetLogging (O5)"
    config.EnableSheetLogging = ReadBoolCell(ws, "O5", MODULE_NAME, funcName, currentItem, True)
    currentItem = "EnableSearchConditionLogSheetOutput (O6)"
    config.EnableSearchConditionLogSheetOutput = ReadBoolCell(ws, "O6", MODULE_NAME, funcName, currentItem, True)
    currentItem = "EnableErrorLogSheetOutput (O7)"
    config.EnableErrorLogSheetOutput = ReadBoolCell(ws, "O7", MODULE_NAME, funcName, currentItem, True)
    currentItem = "DebugDetailLevel2Enabled (O8)"
    config.DebugDetailLevel2Enabled = ReadBoolCell(ws, "O8", MODULE_NAME, funcName, currentItem, False)
    currentItem = "DebugDetailLevel3Enabled (O9)"
    config.DebugDetailLevel3Enabled = ReadBoolCell(ws, "O9", MODULE_NAME, funcName, currentItem, False)

    currentItem = "DefaultFolderPath (O12)"
    config.DefaultFolderPath = ReadStringCell(ws, "O12", MODULE_NAME, funcName, currentItem)
    currentItem = "LogSheetName (O42)"
    config.LogSheetName = ReadStringCell(ws, "O42", MODULE_NAME, funcName, currentItem, "Log")
    currentItem = "OutputSheetName (O43)"
    config.OutputSheetName = ReadStringCell(ws, "O43", MODULE_NAME, funcName, currentItem, "抽出結果")
    currentItem = "SearchConditionLogSheetName (O44)"
    config.SearchConditionLogSheetName = ReadStringCell(ws, "O44", MODULE_NAME, funcName, currentItem, "検索条件ログ")
    currentItem = "ErrorLogSheetName (O45)"
    config.ErrorLogSheetName = ReadStringCell(ws, "O45", MODULE_NAME, funcName, currentItem, "エラーログ")
    currentItem = "ConfigSheetName (O46)"
    config.ConfigSheetName = ReadStringCell(ws, "O46", MODULE_NAME, funcName, currentItem, CONFIG_SHEET_DEFAULT_NAME)
    currentItem = "GetPatternDataMethod (O122)"
    config.GetPatternDataMethod = ReadBoolCell(ws, "O122", MODULE_NAME, funcName, currentItem)
    Exit Sub
ErrorHandler_LoadGeneralSettings:
    Call HandleConfigLoadingError(MODULE_NAME, funcName, "一般設定「" & currentItem & "」読込エラー", Err.Number, Err.Description) ' Updated Call
End Sub

Private Sub LoadScheduleFileSettings(ByRef config As tConfigSettings, ByVal ws As Worksheet)
    Dim funcName As String: funcName = "LoadScheduleFileSettings"
    Dim currentItem As String
    Dim rawData As Variant

    On Error GoTo ErrorHandler_LoadScheduleFileSettings ' Ensure error handler is set at the beginning

    currentItem = "TargetSheetNames (O66:O75)" ' Set currentItem for context
    rawData = ReadRangeToArray(ws, "O66:O75", MODULE_NAME, funcName, currentItem)

    ' Added logging for rawData
    Debug.Print Now & " LoadScheduleFileSettings: Before calling ConvertRawVariantToStringArray for " & currentItem
    Debug.Print Now & " LoadScheduleFileSettings: TypeName(rawData) = " & TypeName(rawData)
    If IsArray(rawData) Then
        On Error Resume Next ' Handle potential errors if rawData is not a valid array for LBound/UBound
        Debug.Print Now & " LoadScheduleFileSettings: LBound(rawData, 1) = " & LBound(rawData, 1) & ", UBound(rawData, 1) = " & UBound(rawData, 1)
        If Err.Number <> 0 Then
            Debug.Print Now & " LoadScheduleFileSettings: Error accessing LBound/UBound for rawData (1st dim): " & Err.Description
            Err.Clear
        End If
        ' Corrected VarType usage and check for multi-dimensional array
        ' Check if rawData is an array and if it has more than one dimension.
        ' VarType(rawData) will return vbArray (8192) ORed with the base type if it's an array.
        ' To check for dimensions, we need to attempt to get UBound of a second dimension.
        Dim numDimensions As Integer
        numDimensions = 1 ' Assume 1 dimension by default
        On Error Resume Next ' Temporarily for UBound on 2nd dimension
        Dim uboundDim2Test As Long
        uboundDim2Test = UBound(rawData, 2)
        If Err.Number = 0 Then
            numDimensions = 2
            Debug.Print Now & " LoadScheduleFileSettings: LBound(rawData, 2) = " & LBound(rawData, 2) & ", UBound(rawData, 2) = " & uboundDim2Test
        Else
            ' It's a 1-dimensional array if UBound on 2nd dimension fails.
             Debug.Print Now & " LoadScheduleFileSettings: rawData appears to be 1-dimensional (Error accessing UBound for 2nd dim): " & Err.Description
        End If
        Err.Clear
        On Error GoTo ErrorHandler_LoadScheduleFileSettings ' Restore original error handler
    ElseIf IsEmpty(rawData) Then
        Debug.Print Now & " LoadScheduleFileSettings: rawData is Empty."
    Else
        Debug.Print Now & " LoadScheduleFileSettings: rawData is a scalar value: " & CStr(rawData)
    End If

    config.TargetSheetNames = ConvertRawVariantToStringArray(rawData, MODULE_NAME, funcName, currentItem, config)
    Call DebugPrintArrayState(config.TargetSheetNames, currentItem, config)

    currentItem = "HeaderRowCount (O87)"
    config.HeaderRowCount = ReadLongCell(ws, "O87", MODULE_NAME, funcName, currentItem)
    currentItem = "HeaderColCount (O88)"
    config.HeaderColCount = ReadLongCell(ws, "O88", MODULE_NAME, funcName, currentItem)
    currentItem = "RowsPerDay (O89)"
    config.RowsPerDay = ReadLongCell(ws, "O89", MODULE_NAME, funcName, currentItem)
    currentItem = "MaxDaysPerSheet (O90)"
    config.MaxDaysPerSheet = ReadLongCell(ws, "O90", MODULE_NAME, funcName, currentItem)
    currentItem = "YearCellAddress (O101)"
    config.YearCellAddress = ReadStringCell(ws, "O101", MODULE_NAME, funcName, currentItem)
    currentItem = "MonthCellAddress (O102)"
    config.MonthCellAddress = ReadStringCell(ws, "O102", MODULE_NAME, funcName, currentItem)
    currentItem = "DayColumnLetter (O103)"
    config.DayColumnLetter = ReadStringCell(ws, "O103", MODULE_NAME, funcName, currentItem)
    currentItem = "DayRowOffset (O104)"
    config.DayRowOffset = ReadLongCell(ws, "O104", MODULE_NAME, funcName, currentItem)
    currentItem = "ProcessesPerDay (O114)"
    config.ProcessesPerDay = ReadLongCell(ws, "O114", MODULE_NAME, funcName, currentItem)
    Exit Sub
ErrorHandler_LoadScheduleFileSettings:
    Call M02_LogError(MODULE_NAME, funcName, "工程表ファイル設定「" & currentItem & "」読込エラー", Err.Number, Err.Description)
End Sub

Private Sub LoadProcessPatternDefinition(ByRef config As tConfigSettings, ByVal ws As Worksheet)
    Dim funcName As String: funcName = "LoadProcessPatternDefinition"
    Dim procPtn_i As Long
    Dim numProcesses As Long
    Dim currentItem As String
    Dim rawData As Variant
    On Error GoTo ErrorHandler_LoadProcessPatternDefinition

    currentItem = "CurrentPatternIdentifier (O126)"
    config.CurrentPatternIdentifier = ReadStringCell(ws, "O126", MODULE_NAME, funcName, currentItem)

    numProcesses = config.ProcessesPerDay
    If numProcesses <= 0 Then numProcesses = 10

    currentItem = "ProcessKeys (I129:I" & (128 + numProcesses) & ")"
    rawData = ReadRangeToArray(ws, "I129:I" & (128 + numProcesses), MODULE_NAME, funcName, currentItem)
    config.ProcessKeys = ConvertRawVariantToStringArray(rawData, MODULE_NAME, funcName, currentItem, config)
    Call DebugPrintArrayState(config.ProcessKeys, currentItem, config)

    currentItem = "Kankatsu1List (J129:J" & (128 + numProcesses) & ")"
    rawData = ReadRangeToArray(ws, "J129:J" & (128 + numProcesses), MODULE_NAME, funcName, currentItem)
    config.Kankatsu1List = ConvertRawVariantToStringArray(rawData, MODULE_NAME, funcName, currentItem, config)
    Call DebugPrintArrayState(config.Kankatsu1List, currentItem, config)

    currentItem = "Kankatsu2List (K129:K" & (128 + numProcesses) & ")"
    rawData = ReadRangeToArray(ws, "K129:K" & (128 + numProcesses), MODULE_NAME, funcName, currentItem)
    config.Kankatsu2List = ConvertRawVariantToStringArray(rawData, MODULE_NAME, funcName, currentItem, config)
    Call DebugPrintArrayState(config.Kankatsu2List, currentItem, config)

    currentItem = "Bunrui1List (L129:L" & (128 + numProcesses) & ")"
    rawData = ReadRangeToArray(ws, "L129:L" & (128 + numProcesses), MODULE_NAME, funcName, currentItem)
    config.Bunrui1List = ConvertRawVariantToStringArray(rawData, MODULE_NAME, funcName, currentItem, config)
    Call DebugPrintArrayState(config.Bunrui1List, currentItem, config)

    currentItem = "Bunrui2List (M129:M" & (128 + numProcesses) & ")"
    rawData = ReadRangeToArray(ws, "M129:M" & (128 + numProcesses), MODULE_NAME, funcName, currentItem)
    config.Bunrui2List = ConvertRawVariantToStringArray(rawData, MODULE_NAME, funcName, currentItem, config)
    Call DebugPrintArrayState(config.Bunrui2List, currentItem, config)

    currentItem = "Bunrui3List (N129:N" & (128 + numProcesses) & ")"
    rawData = ReadRangeToArray(ws, "N129:N" & (128 + numProcesses), MODULE_NAME, funcName, currentItem)
    config.Bunrui3List = ConvertRawVariantToStringArray(rawData, MODULE_NAME, funcName, currentItem, config)
    Call DebugPrintArrayState(config.Bunrui3List, currentItem, config)

    currentItem = "ProcessColCountSheetHeaders (O128:X128)"
    Dim headerData As Variant
    headerData = ws.Range("O128:X128").Value
    If IsArray(headerData) Then
        On Error Resume Next
        Dim ub As Long: ub = UBound(headerData, 2)
        If Err.Number = 0 Then
            ReDim config.ProcessColCountSheetHeaders(1 To ub)
            For procPtn_i = 1 To ub
                config.ProcessColCountSheetHeaders(procPtn_i) = Trim(CStr(headerData(1, procPtn_i)))
            Next procPtn_i
        Else
            ReDim config.ProcessColCountSheetHeaders(1 To 0)
            Call M04_LogWriter.WriteErrorLog("WARNING", MODULE_NAME, funcName, currentItem & " - UBound(headerData, 2) failed. Headers not loaded.")
        End If
        On Error GoTo ErrorHandler_LoadProcessPatternDefinition
    ElseIf Not IsEmpty(headerData) Then
        ReDim config.ProcessColCountSheetHeaders(1 To 1)
        config.ProcessColCountSheetHeaders(1) = Trim(CStr(headerData))
    Else
        ReDim config.ProcessColCountSheetHeaders(1 To 0)
    End If
    Call DebugPrintArrayState(config.ProcessColCountSheetHeaders, currentItem, config)

    currentItem = "ProcessColCounts (O129:X" & (128 + numProcesses) & ")"
    config.ProcessColCounts = ws.Range("O129:X" & (128 + numProcesses)).Value

    If General_IsArrayInitialized(config.Kankatsu1List) And General_IsArrayInitialized(config.Kankatsu2List) Then
        Dim k1Count As Long, k2Count As Long, maxCount As Long
        k1Count = 0: If UBound(config.Kankatsu1List) >= LBound(config.Kankatsu1List) Then k1Count = UBound(config.Kankatsu1List)
        k2Count = 0: If UBound(config.Kankatsu2List) >= LBound(config.Kankatsu2List) Then k2Count = UBound(config.Kankatsu2List)
        maxCount = IIf(k1Count > k2Count, k1Count, k2Count)
        If maxCount > 0 Then
            ReDim config.ProcessDetails(1 To maxCount)
            For procPtn_i = 1 To maxCount
                If procPtn_i <= k1Count Then config.ProcessDetails(procPtn_i).Kankatsu1 = config.Kankatsu1List(procPtn_i)
                If procPtn_i <= k2Count Then config.ProcessDetails(procPtn_i).Kankatsu2 = config.Kankatsu2List(procPtn_i)
            Next procPtn_i
        Else
            ReDim config.ProcessDetails(1 To 0)
        End If
    Else
        ReDim config.ProcessDetails(1 To 0)
    End If
    Exit Sub
ErrorHandler_LoadProcessPatternDefinition:
    Call M02_LogError(MODULE_NAME, funcName, "工程パターン定義「" & currentItem & "」読込エラー", Err.Number, Err.Description) ' Updated Call
End Sub

Private Sub LoadFilterConditions(ByRef config As tConfigSettings, ByVal ws As Worksheet)
    Dim funcName As String: funcName = "LoadFilterConditions"
    Dim currentItem As String
    Dim rawData As Variant
    On Error GoTo ErrorHandler_LoadFilterConditions

    currentItem = "WorkerFilterLogic (O242)"
    config.WorkerFilterLogic = ReadStringCell(ws, "O242", MODULE_NAME, funcName, "作業員フィルター検索論理", "AND")

    currentItem = "WorkerFilterList (O243:O262)"
    rawData = ReadRangeToArray(ws, "O243:O262", MODULE_NAME, funcName, currentItem)
    config.WorkerFilterList = ConvertRawVariantToStringArray(rawData, MODULE_NAME, funcName, currentItem, config)
    Call DebugPrintArrayState(config.WorkerFilterList, currentItem, config)

    currentItem = "Kankatsu1FilterList (O275:O294)"
    rawData = ReadRangeToArray(ws, "O275:O294", MODULE_NAME, funcName, currentItem)
    config.Kankatsu1FilterList = ConvertRawVariantToStringArray(rawData, MODULE_NAME, funcName, currentItem, config)
    Call DebugPrintArrayState(config.Kankatsu1FilterList, currentItem, config)

    currentItem = "Kankatsu2FilterList (O305:O334)"
    rawData = ReadRangeToArray(ws, "O305:O334", MODULE_NAME, funcName, currentItem)
    config.Kankatsu2FilterList = ConvertRawVariantToStringArray(rawData, MODULE_NAME, funcName, currentItem, config)
    Call DebugPrintArrayState(config.Kankatsu2FilterList, currentItem, config)

    currentItem = "Bunrui1Filter (O346)"
    config.Bunrui1Filter = ReadStringCell(ws, "O346", MODULE_NAME, funcName, "分類1フィルター")
    currentItem = "Bunrui2Filter (O367)"
    config.Bunrui2Filter = ReadStringCell(ws, "O367", MODULE_NAME, funcName, "分類2フィルター")
    currentItem = "Bunrui3Filter (O388)"
    config.Bunrui3Filter = ReadStringCell(ws, "O388", MODULE_NAME, funcName, "分類3フィルター")

    currentItem = "KoujiShuruiFilterList (O409:O418)"
    rawData = ReadRangeToArray(ws, "O409:O418", MODULE_NAME, funcName, currentItem)
    config.KoujiShuruiFilterList = ConvertRawVariantToStringArray(rawData, MODULE_NAME, funcName, currentItem, config)
    Call DebugPrintArrayState(config.KoujiShuruiFilterList, currentItem, config)

    currentItem = "KoubanFilterList (O431:O440)"
    rawData = ReadRangeToArray(ws, "O431:O440", MODULE_NAME, funcName, currentItem)
    config.KoubanFilterList = ConvertRawVariantToStringArray(rawData, MODULE_NAME, funcName, currentItem, config)
    Call DebugPrintArrayState(config.KoubanFilterList, currentItem, config)

    currentItem = "SagyoushuruiFilterList (O451:O470)"
    rawData = ReadRangeToArray(ws, "O451:O470", MODULE_NAME, funcName, currentItem)
    config.SagyoushuruiFilterList = ConvertRawVariantToStringArray(rawData, MODULE_NAME, funcName, currentItem, config)
    Call DebugPrintArrayState(config.SagyoushuruiFilterList, currentItem, config)

    currentItem = "TantouFilterList (O481:O490)"
    rawData = ReadRangeToArray(ws, "O481:O490", MODULE_NAME, funcName, currentItem)
    config.TantouFilterList = ConvertRawVariantToStringArray(rawData, MODULE_NAME, funcName, currentItem, config)
    Call DebugPrintArrayState(config.TantouFilterList, currentItem, config)

    currentItem = "NinzuFilter (O503)"
    config.NinzuFilter = ReadStringCell(ws, "O503", MODULE_NAME, funcName, "人数フィルター")
    config.IsNinzuFilterOriginallyEmpty = (Trim(config.NinzuFilter) = "")

    currentItem = "SagyouKashoKindFilter (O514)"
    config.SagyouKashoKindFilter = ReadStringCell(ws, "O514", MODULE_NAME, funcName, "作業箇所の種類フィルター")

    currentItem = "SagyouKashoFilterList (O525:O544)"
    rawData = ReadRangeToArray(ws, "O525:O544", MODULE_NAME, funcName, currentItem)
    config.SagyouKashoFilterList = ConvertRawVariantToStringArray(rawData, MODULE_NAME, funcName, currentItem, config)
    Call DebugPrintArrayState(config.SagyouKashoFilterList, currentItem, config)

    Exit Sub
ErrorHandler_LoadFilterConditions:
    Call M02_LogError(MODULE_NAME, funcName, "フィルター条件「" & currentItem & "」読込エラー", Err.Number, Err.Description) ' Updated Call
End Sub

Private Sub LoadTargetFileDefinition(ByRef config As tConfigSettings, ByVal ws As Worksheet)
    Dim funcName As String: funcName = "LoadTargetFileDefinition"
    Dim currentItem As String
    Dim rawData As Variant
    On Error GoTo ErrorHandler_LoadTargetFileDefinition

    currentItem = "TargetFileFolderPaths (P557:P756)"
    rawData = ReadRangeToArray(ws, "P557:P756", MODULE_NAME, funcName, currentItem)
    config.TargetFileFolderPaths = ConvertRawVariantToStringArray(rawData, MODULE_NAME, funcName, currentItem, config)
    Call DebugPrintArrayState(config.TargetFileFolderPaths, currentItem, config)

    currentItem = "FilePatternIdentifiers (Q557:Q756)"
    rawData = ReadRangeToArray(ws, "Q557:Q756", MODULE_NAME, funcName, currentItem)
    config.FilePatternIdentifiers = ConvertRawVariantToStringArray(rawData, MODULE_NAME, funcName, currentItem, config)
    Call DebugPrintArrayState(config.FilePatternIdentifiers, currentItem, config)

    Exit Sub
ErrorHandler_LoadTargetFileDefinition:
    Call M02_LogError(MODULE_NAME, funcName, "処理対象ファイル定義「" & currentItem & "」読込エラー", Err.Number, Err.Description) ' Updated Call
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
        If IsEmpty(val) Or Trim(CStr(val)) = "" Then ReadStringCell = defaultValue Else ReadStringCell = Trim(CStr(val))
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
        If IsEmpty(val) Or Not IsNumeric(val) Then ReadLongCell = defaultValue Else ReadLongCell = CLng(val)
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
        If IsEmpty(val) Then ReadBoolCell = defaultValue Else ReadBoolCell = (UCase(Trim(CStr(val))) = "TRUE")
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

Private Function ConvertRawVariantToStringArray(ByVal rawData As Variant, ByVal moduleN As String, ByVal funcN_from_caller As String, ByVal itemName As String, ByRef currentConfig As tConfigSettings) As String()
    Debug.Print Now & " CVTSA_CALLED for: '" & itemName & "' from '" & funcN_from_caller & "'"

    Debug.Print Now & " CVTSA_Point_001: Before Dim cvtsaTempArray()"
    Dim cvtsaTempArray() As String
    Debug.Print Now & " CVTSA_Point_002: After Dim cvtsaTempArray(), Before Dim i As Long"
    Dim i As Long
    Debug.Print Now & " CVTSA_Point_003: After Dim i As Long, Before Dim r As Long"
    Dim r As Long
    Debug.Print Now & " CVTSA_Point_004: After Dim r As Long, Before Dim count As Long"
    Dim count As Long
    Debug.Print Now & " CVTSA_Point_005: After Dim count As Long, Before Dim lBound1 As Long"
    Dim lBound1 As Long
    Debug.Print Now & " CVTSA_Point_006: After Dim lBound1 As Long, Before Dim uBound1 As Long"
    Dim uBound1 As Long
    Debug.Print Now & " CVTSA_Point_007: After Dim uBound1 As Long, Before Dim lBound2 As Long"
    Dim lBound2 As Long
    Debug.Print Now & " CVTSA_Point_008: After Dim lBound2 As Long, Before Dim uBound2 As Long"
    Dim uBound2 As Long
    Debug.Print Now & " CVTSA_Point_009: After Dim uBound2 As Long, Before Dim tempMsg As String"
    Dim tempMsg As String
    Debug.Print Now & " CVTSA_Point_010: After Dim tempMsg As String, Before Dim localFuncName As String"
    Dim localFuncName As String
    Debug.Print Now & " CVTSA_Point_011: After Dim localFuncName As String, Before localFuncName assignment"
    localFuncName = "ConvertRawVariantToStringArray"
    Debug.Print Now & " CVTSA_Point_012: After localFuncName assignment, Before On Error GoTo 0"

    On Error GoTo 0
    Debug.Print Now & " CVTSA_Point_013: After On Error GoTo 0, Before On Error GoTo ErrorHandler_CVTSA_Direct"
    On Error GoTo ErrorHandler_CVTSA_Direct
    Debug.Print Now & " CVTSA_Point_014: After On Error GoTo ErrorHandler_CVTSA_Direct, Before Err.Clear and ReDim cvtsaTempArray(1 To 0)"

    Err.Clear ' Added Err.Clear
    Debug.Print Now & " CVTSA_Point_014a: After Err.Clear. Err.Number=" & Err.Number & ", Err.Description='" & Err.Description & "'"
    Debug.Print Now & " CVTSA_Point_014b: Before ReDim cvtsaTempArray(1 To 0). TypeName(cvtsaTempArray)=" & TypeName(cvtsaTempArray)
    ' Initial ReDim cvtsaTempArray(1 To 0) REMOVED as per new logic

    Debug.Print Now & " CVTSA_Point_015: After initial ReDim was removed. TypeName(cvtsaTempArray)=" & TypeName(cvtsaTempArray) & ". Err.Number=" & Err.Number & ", Err.Description='" & Err.Description & "'"
    ' Note: LBound/UBound checks here would fail if array is not yet dimensioned by data population.

    ' This is the first point where currentConfig is accessed for its members
    Debug.Print Now & " CVTSA_Point_015a: Before checking currentConfig.DebugDetailLevel2Enabled for " & itemName
    If currentConfig.DebugDetailLevel2Enabled Then
        Debug.Print Now & " CVTSA_Point_016: Before First DEBUG_L2 Log Call for " & itemName
        tempMsg = "START for item: '" & itemName & "' (called by " & funcN_from_caller & "). TypeName(rawData): " & TypeName(rawData)
        Call M04_LogWriter.WriteErrorLog("DEBUG_L2", moduleN, localFuncName, tempMsg)
        Debug.Print Now & " CVTSA_Point_017: After First DEBUG_L2 Log Call for " & itemName
    End If

    If IsEmpty(rawData) Then
        If currentConfig.DebugDetailLevel2Enabled Then Call M04_LogWriter.WriteErrorLog("DEBUG_L2", moduleN, localFuncName, itemName & " - rawData is Empty (called by " & funcN_from_caller & "). Returning empty array (1 To 0).")
        ' cvtsaTempArray is already (1 To 0)
    ElseIf Not IsArray(rawData) Then
        If currentConfig.DebugDetailLevel2Enabled Then Call M04_LogWriter.WriteErrorLog("DEBUG_L2", moduleN, localFuncName, itemName & " - rawData is Scalar (called by " & funcN_from_caller & "). Value: '" & CStr(rawData) & "'")
        If Trim(CStr(rawData)) <> "" Then
            ReDim cvtsaTempArray(1 To 1)
            cvtsaTempArray(1) = Trim(CStr(rawData))
            If currentConfig.DebugDetailLevel2Enabled Then Call M04_LogWriter.WriteErrorLog("DEBUG_L2", moduleN, localFuncName, itemName & " - Scalar converted to 1-element array. LBound=" & LBound(cvtsaTempArray) & ", UBound=" & UBound(cvtsaTempArray))
        Else
            ' cvtsaTempArray is already (1 To 0)
            If currentConfig.DebugDetailLevel2Enabled Then Call M04_LogWriter.WriteErrorLog("DEBUG_L2", moduleN, localFuncName, itemName & " - Scalar is empty. Returning (1 To 0) array.")
        End If
    Else ' IsArray(rawData) is True
        If currentConfig.DebugDetailLevel2Enabled Then Call M04_LogWriter.WriteErrorLog("DEBUG_L2", moduleN, localFuncName, itemName & " - rawData is Array (called by " & funcN_from_caller & ").")

        Dim dbg_l1 As String, dbg_u1 As String, dbg_l2 As String, dbg_u2 As String
        Dim numDimensions As Integer

        On Error Resume Next ' Temporarily for LBound/UBound calls
        lBound1 = LBound(rawData, 1): dbg_l1 = CStr(lBound1)
        uBound1 = UBound(rawData, 1): dbg_u1 = CStr(uBound1)
        If Err.Number <> 0 Then
            Debug.Print Now & " " & localFuncName & ": DEBUG - Failed LBound/UBound on 1st dimension for '" & itemName & "'. Error: " & Err.Description
            dbg_l1 = "Err:" & Err.Number: dbg_u1 = "Err:" & Err.Description ' Capture error for logging if needed
            Err.Clear
            On Error GoTo ErrorHandler_CVTSA_Direct ' Restore before GoTo
            GoTo InvalidArrayStructure_CVTSA_Direct ' Treat as invalid structure
        End If

        lBound2 = LBound(rawData, 2): dbg_l2 = CStr(lBound2)
        uBound2 = UBound(rawData, 2): dbg_u2 = CStr(uBound2)
        If Err.Number = 0 Then
            numDimensions = 2
        Else
            numDimensions = 1
            Err.Clear
        End If
        On Error GoTo ErrorHandler_CVTSA_Direct ' Restore main error handler immediately

        If currentConfig.DebugDetailLevel2Enabled Then
            If numDimensions = 1 Then
                Call M04_LogWriter.WriteErrorLog("DEBUG_L2", moduleN, localFuncName, itemName & " - Detected 1D Array. Bounds: " & dbg_l1 & " To " & dbg_u1)
            Else ' numDimensions = 2
                Call M04_LogWriter.WriteErrorLog("DEBUG_L2", moduleN, localFuncName, itemName & " - Detected 2D Array. Bounds1: " & dbg_l1 & " To " & dbg_u1 & ". Bounds2: " & dbg_l2 & " To " & dbg_u2)
            End If
        End If

        If numDimensions = 1 Then
            If uBound1 >= lBound1 Then
                Dim newUBound1D As Long: newUBound1D = uBound1 - lBound1 + 1
                If currentConfig.DebugDetailLevel2Enabled Then Call M04_LogWriter.WriteErrorLog("DEBUG_L2", moduleN, localFuncName, itemName & " - 1D: Calc UBound for cvtsaTempArray: " & newUBound1D)
                ReDim cvtsaTempArray(1 To newUBound1D)
                If currentConfig.DebugDetailLevel2Enabled Then Call M04_LogWriter.WriteErrorLog("DEBUG_L2", moduleN, localFuncName, itemName & " - 1D: cvtsaTempArray ReDim'd. LBound=" & LBound(cvtsaTempArray) & ", UBound=" & UBound(cvtsaTempArray))
                count = 0
                For i = lBound1 To uBound1
                    If currentConfig.DebugDetailLevel2Enabled Then Call M04_LogWriter.WriteErrorLog("DEBUG_L2", moduleN, localFuncName, itemName & " - 1D Loop i=" & i & ", rawData(i)='" & CStr(rawData(i)) & "'. current count=" & count & ". Target cvtsaTempArray(" & count + 1 & ")")
                    If Not IsEmpty(rawData(i)) And Trim(CStr(rawData(i))) <> "" Then
                        count = count + 1
                        cvtsaTempArray(count) = Trim(CStr(rawData(i)))
                        If currentConfig.DebugDetailLevel2Enabled Then Call M04_LogWriter.WriteErrorLog("DEBUG_L2", moduleN, localFuncName, itemName & " - 1D Loop: Added to cvtsaTempArray(" & count & ") = '" & cvtsaTempArray(count) & "'")
                    Else
                        If currentConfig.DebugDetailLevel2Enabled Then Call M04_LogWriter.WriteErrorLog("DEBUG_L2", moduleN, localFuncName, itemName & " - 1D Loop: Skipped empty/whitespace rawData(" & i & ")")
                    End If
                Next i
                If currentConfig.DebugDetailLevel2Enabled Then Call M04_LogWriter.WriteErrorLog("DEBUG_L2", moduleN, localFuncName, itemName & " - 1D Loop END. count=" & count & ". UBound(cvtsaTempArray)=" & UBound(cvtsaTempArray) & ". Attempting ReDim Preserve to (1 To " & count & ") if count < UBound and count > 0.")
                If count > 0 Then
                    If count < UBound(cvtsaTempArray) Then ReDim Preserve cvtsaTempArray(1 To count)
                Else
                    ReDim cvtsaTempArray(1 To 0)
                End If
                If currentConfig.DebugDetailLevel2Enabled Then Call M04_LogWriter.WriteErrorLog("DEBUG_L2", moduleN, localFuncName, itemName & " - 1D: cvtsaTempArray final state post-preserve/reset. LBound=" & LBound(cvtsaTempArray) & ", UBound=" & UBound(cvtsaTempArray))
            End If
        ElseIf numDimensions = 2 And lBound2 = 1 And uBound2 = 1 Then
            If uBound1 >= lBound1 Then
                Dim newUBound2D As Long: newUBound2D = uBound1 - lBound1 + 1
                If currentConfig.DebugDetailLevel2Enabled Then Call M04_LogWriter.WriteErrorLog("DEBUG_L2", moduleN, localFuncName, itemName & " - 2D-Vert: Calc UBound for cvtsaTempArray: " & newUBound2D)
                ReDim cvtsaTempArray(1 To newUBound2D)
                If currentConfig.DebugDetailLevel2Enabled Then Call M04_LogWriter.WriteErrorLog("DEBUG_L2", moduleN, localFuncName, itemName & " - 2D-Vert: cvtsaTempArray ReDim'd. LBound=" & LBound(cvtsaTempArray) & ", UBound=" & UBound(cvtsaTempArray))
                count = 0
                For r = lBound1 To uBound1
                    If currentConfig.DebugDetailLevel2Enabled Then Call M04_LogWriter.WriteErrorLog("DEBUG_L2", moduleN, localFuncName, itemName & " - 2D-Vert Loop r=" & r & ", rawData(r,1)='" & CStr(rawData(r, lBound2)) & "'. current count=" & count & ". Target cvtsaTempArray(" & count + 1 & ")")
                    If Not IsEmpty(rawData(r, lBound2)) And Trim(CStr(rawData(r, lBound2))) <> "" Then
                        count = count + 1
                        cvtsaTempArray(count) = Trim(CStr(rawData(r, lBound2)))
                        If currentConfig.DebugDetailLevel2Enabled Then Call M04_LogWriter.WriteErrorLog("DEBUG_L2", moduleN, localFuncName, itemName & " - 2D-Vert Loop: Added to cvtsaTempArray(" & count & ") = '" & cvtsaTempArray(count) & "'")
                    Else
                         If currentConfig.DebugDetailLevel2Enabled Then Call M04_LogWriter.WriteErrorLog("DEBUG_L2", moduleN, localFuncName, itemName & " - 2D-Vert Loop: Skipped empty/whitespace rawData(" & r & ", " & lBound2 & ")")
                    End If
                Next r
                If currentConfig.DebugDetailLevel2Enabled Then Call M04_LogWriter.WriteErrorLog("DEBUG_L2", moduleN, localFuncName, itemName & " - 2D-Vert Loop END. count=" & count & ". UBound(cvtsaTempArray)=" & UBound(cvtsaTempArray) & ". Attempting ReDim Preserve to (1 To " & count & ") if count < UBound and count > 0.")
                If count > 0 Then
                    If count < UBound(cvtsaTempArray) Then ReDim Preserve cvtsaTempArray(1 To count)
                Else
                    ReDim cvtsaTempArray(1 To 0)
                End If
                If currentConfig.DebugDetailLevel2Enabled Then Call M04_LogWriter.WriteErrorLog("DEBUG_L2", moduleN, localFuncName, itemName & " - 2D-Vert: cvtsaTempArray final state post-preserve/reset. LBound=" & LBound(cvtsaTempArray) & ", UBound=" & UBound(cvtsaTempArray))
            End If
        Else
InvalidArrayStructure_CVTSA_Direct:
            On Error GoTo ErrorHandler_CVTSA_Direct ' Ensure handler is set before logging (though WriteErrorLog has its own)
            If currentConfig.DebugDetailLevel2Enabled Then
                tempMsg = itemName & " - 予期しない配列構造またはエラーのある配列 (呼び出し元: " & funcN_from_caller & "). L1:" & dbg_l1 & ", U1:" & dbg_u1 & ", L2:" & dbg_l2 & ", U2:" & dbg_u2 & ". 空として扱います。"
                Call M04_LogWriter.WriteErrorLog("WARNING_L2", moduleN, localFuncName, tempMsg)
            End If
            ReDim cvtsaTempArray(1 To 0) ' Ensure defined as empty
        End If
    End If

    If Not General_IsArrayInitialized(cvtsaTempArray) Then
        Debug.Print Now & " CVTSA_INFO: cvtsaTempArray was not initialized by data population. Attempting ReDim (1 To 0)."
        On Error Resume Next ' To catch the problematic ReDim
        ReDim cvtsaTempArray(1 To 0)
        If Err.Number <> 0 Then
            Debug.Print Now & " CVTSA_ERR_FINAL_REDIM: Error " & Err.Number & " during final ReDim (1 To 0): " & Err.Description
            ' As a fallback, create a truly empty array that should be safe
            cvtsaTempArray = Split(vbNullString, Chr(1)) ' Chr(1) is unlikely to be in vbNullString, ensuring a 0-element array with LBound 0
        End If
        On Error GoTo ErrorHandler_CVTSA_Direct ' Restore original handler
    End If

    ConvertRawVariantToStringArray = cvtsaTempArray
    Exit Function

ErrorHandler_CVTSA_Direct:
    Debug.Print Now & " CVTSA_Point_ERR: ENTERING ErrorHandler_CVTSA_Direct for " & itemName
    Debug.Print Now & " CVTSA_Point_ERR_DETAIL: Err.Source='" & Err.Source & "', Err.HelpFile='" & Err.HelpFile & "', Err.HelpContext=" & Err.HelpContext
    Debug.Print Now & " CVTSA_Point_ERR_CONTEXT: TypeName(cvtsaTempArray) before ReDim in Handler = " & TypeName(cvtsaTempArray)

    tempMsg = itemName & " の変換中に予期せぬエラー (呼び出し元: " & funcN_from_caller & ")"
    Debug.Print Now & " CRITICAL_ERROR in " & localFuncName & ": " & tempMsg & " Err# " & Err.Number & " - " & Err.Description

    On Error Resume Next ' Temporarily disable error handling
    Debug.Print Now & " CVTSA_Point_ERR_HANDLER_BEFORE_REDIM: Attempting ReDim cvtsaTempArray(1 To 0) within error handler."
    ReDim cvtsaTempArray(1 To 0)
    Debug.Print Now & " CVTSA_Point_ERR_HANDLER_AFTER_REDIM: After ReDim in handler. TypeName(cvtsaTempArray)=" & TypeName(cvtsaTempArray) & ", LBound=" & LBound(cvtsaTempArray) & ", UBound=" & UBound(cvtsaTempArray) & ", Err.Number (after ReDim in handler)=" & Err.Number & ", Err.Description (after ReDim in handler)='" & Err.Description & "'"
    Err.Clear ' Clear any error that might have occurred during ReDim in handler
    ' On Error GoTo 0 ' This will be restored after the check or by the final On Error GoTo 0 if an error occurs during the check

    If Not General_IsArrayInitialized(cvtsaTempArray) Then
        Debug.Print Now & " CVTSA_INFO_HANDLER: cvtsaTempArray not initialized in error handler. Attempting ReDim (1 To 0)."
        On Error Resume Next ' To catch the problematic ReDim
        ReDim cvtsaTempArray(1 To 0)
        If Err.Number <> 0 Then
            Debug.Print Now & " CVTSA_ERR_HANDLER_FINAL_REDIM: Error " & Err.Number & " during final ReDim (1 To 0) in handler: " & Err.Description
            cvtsaTempArray = Split(vbNullString, Chr(1))
        End If
        ' No need to restore ErrorHandler_CVTSA_Direct as we are already in it.
        ' The next On Error GoTo 0 will handle subsequent errors.
    End If
    On Error GoTo 0 ' Restore default error handling (or the main one for this handler scope)

    ConvertRawVariantToStringArray = cvtsaTempArray
    Debug.Print Now & " CVTSA_Point_ERR_EXIT: Exiting ErrorHandler_CVTSA_Direct for " & itemName
End Function

Private Sub DebugPrintArrayState(ByRef arr As Variant, ByVal arrName As String, ByRef currentConfig As tConfigSettings)
    Dim l As Long, u As Long
    Dim msg As String
    If Not IsArray(arr) Then
        msg = arrName & " is not an array. TypeName: " & TypeName(arr)
        If currentConfig.DebugDetailLevel2Enabled Then Debug.Print msg
        Call M04_LogWriter.WriteErrorLog("DEBUG_ARRAY_STATE", "M02_ConfigReader", "DebugPrintArrayState", msg)
        Exit Sub
    End If
    On Error Resume Next
    l = LBound(arr)
    u = UBound(arr)
    If Err.Number <> 0 Then
        msg = arrName & " IsArray=True, but LBound/UBound failed. Err: " & Err.Description
        If currentConfig.DebugDetailLevel2Enabled Then Debug.Print msg
        Call M04_LogWriter.WriteErrorLog("DEBUG_ARRAY_STATE", "M02_ConfigReader", "DebugPrintArrayState", msg)
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
    msg = arrName & " IsArray=True, LBound=" & l & ", UBound=" & u
    If currentConfig.DebugDetailLevel2Enabled Then Debug.Print msg
    Call M04_LogWriter.WriteErrorLog("DEBUG_ARRAY_STATE", "M02_ConfigReader", "DebugPrintArrayState", msg)
End Sub

Public Function General_IsArrayInitialized(arr As Variant) As Boolean
    If Not IsArray(arr) Then
        General_IsArrayInitialized = False
        Exit Function
    End If

    Dim l As Long, u As Long
    On Error Resume Next
    l = LBound(arr)
    u = UBound(arr)
    If Err.Number = 0 Then
        General_IsArrayInitialized = True
    Else
        General_IsArrayInitialized = False
    End If
    On Error GoTo 0
End Function

Private Sub M02_LogError(ByVal moduleN As String, ByVal callerProcName As String, ByVal failedSubName As String, ByVal errNum As Long, ByVal errDesc As String)
    ' This sub is called when a Load... sub finishes and Err.Number is not 0,
    ' meaning an error occurred in the sub and was not handled by Resume Next or Exit Sub within it.
    m_errorOccurred = True ' Set the module-level flag
    Call M04_LogWriter.WriteErrorLog("ERROR", moduleN, callerProcName, failedSubName & " からエラーが伝播 (または新規発生)。", errNum, errDesc)
    ' Do not Clear Err here, let LoadConfiguration handle it or GoTo
End Sub
