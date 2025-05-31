' バージョン：v0.5.0
Option Explicit
' このモジュールは、個々の工程表ファイルを開き、指定されたシートからデータを抽出し、フィルター条件を適用（将来的に）し、結果を出力シートに書き込む（将来的に）処理を担当します。このステップでは、年月と日付の基本情報抽出とログ出力を実装します。

' Private Function LogExtractor_IsArrayInitialized(arr As Variant) As Boolean
    ' 配列が有効に初期化されているか（少なくとも1つの要素を持つか）を確認します。
    ' Variant型が配列でない場合、または配列であっても要素が割り当てられていない場合（Dim arr() のみでReDimされていない状態など）はFalseを返します。
    ' On Error GoTo NotAnArrayOrNotInitialized_M06
    ' If IsArray(arr) Then
    '     Dim lBoundCheck As Long
    '     lBoundCheck = LBound(arr)
    '     LogExtractor_IsArrayInitialized = True
    '     Exit Function
    ' End If
' NotAnArrayOrNotInitialized_M06:
    ' LogExtractor_IsArrayInitialized = False
' End Function

Private Function GetNextFilterLogRow(ByVal logSheetName As String, ByVal mainWB As Workbook) As Long
    ' 指定されたログシート名とワークブックに基づき、次に書き込むべき行番号を取得します。
    If mainWB Is Nothing Or Len(logSheetName) = 0 Then
        GetNextFilterLogRow = 1 ' Fallback
        If DEBUG_MODE_ERROR Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - ERROR: M06_DataExtractor.GetNextFilterLogRow - mainWB is Nothing or logSheetName is empty."
        Exit Function
    End If

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = mainWB.Sheets(logSheetName)
    On Error GoTo 0 ' Or a local error handler for this function

    If ws Is Nothing Then
        GetNextFilterLogRow = 1 ' Fallback if sheet not found
        If DEBUG_MODE_ERROR Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - ERROR: M06_DataExtractor.GetNextFilterLogRow - Sheet '" & logSheetName & "' not found in workbook '" & mainWB.Name & "'."
        Exit Function
    End If

    If Application.WorksheetFunction.CountA(ws.Columns(1)) = 0 Then
        GetNextFilterLogRow = 1
    Else
        GetNextFilterLogRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    End If
End Function

Private Function CreateOffset(r As Long, c As Long) As tOffset
    ' tOffset UDTを作成して返します。
    Dim tempOff As tOffset
    tempOff.Row = r
    tempOff.Col = c
    CreateOffset = tempOff
End Function

Private Function GetValueFromOffset(wsKouteiSheet As Worksheet, baseProcessRow As Long, baseProcessCol As Long, offsetToApply As tOffset, itemDebugName As String, ByRef config As tConfigSettings, mainWB As Workbook) As Variant
    ' 基準セルとオフセットに基づき、工程表シートから値を読み取ります。範囲外やエラー時はログ記録し空文字を返します。
    Dim targetRow As Long, targetCol As Long
    Dim val As Variant

    GetValueFromOffset = "" ' Default return
    On Error GoTo GetValueFromOffset_Error

    targetRow = baseProcessRow + offsetToApply.Row
    targetCol = baseProcessCol + offsetToApply.Col

    ' Validate target coordinates
    If targetRow <= 0 Or targetRow > wsKouteiSheet.Rows.Count Or targetCol <= 0 Or targetCol > wsKouteiSheet.Columns.Count Then
        If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE: M06_DataExtractor.GetValueFromOffset - Offset for '" & itemDebugName & "' is outside sheet bounds. Base(R" & baseProcessRow & "C" & baseProcessCol & ") + Offset(R" & offsetToApply.Row & "C" & offsetToApply.Col & ") -> Target(R" & targetRow & "C" & targetCol & ")"
        Call M04_LogWriter.SafeWriteErrorLog("WARNING", mainWB, config.ErrorLogSheetName, "M06_DataExtractor", "GetValueFromOffset", "オフセットがシート範囲外: " & itemDebugName & " (Target R" & targetRow & "C" & targetCol & ")", 0, wsKouteiSheet.Parent.Name & "/" & wsKouteiSheet.Name)
        Exit Function
    End If

    val = wsKouteiSheet.Cells(targetRow, targetCol).Value

    If IsError(val) Then
        If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE: M06_DataExtractor.GetValueFromOffset - Cell " & itemDebugName & " (R" & targetRow & "C" & targetCol & ") contains an error value: " & CStr(val)
        Call M04_LogWriter.SafeWriteErrorLog("WARNING", mainWB, config.ErrorLogSheetName, "M06_DataExtractor", "GetValueFromOffset", "セルがエラー値: " & itemDebugName & " (R" & targetRow & "C" & targetCol & ") Value: " & CStr(val), 0, wsKouteiSheet.Parent.Name & "/" & wsKouteiSheet.Name)
        Exit Function ' Return ""
    End If

    GetValueFromOffset = Trim(CStr(val))
    Exit Function
GetValueFromOffset_Error:
    If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE: M06_DataExtractor.GetValueFromOffset - Runtime error for '" & itemDebugName & "' at R" & targetRow & "C" & targetCol & ". Error: " & Err.Description
    Call M04_LogWriter.SafeWriteErrorLog("WARNING", mainWB, config.ErrorLogSheetName, "M06_DataExtractor", "GetValueFromOffset", "セル値読取エラー: " & itemDebugName & " (R" & targetRow & "C" & targetCol & ") - " & Err.Description, Err.Number, wsKouteiSheet.Parent.Name & "/" & wsKouteiSheet.Name)
    GetValueFromOffset = "" ' Ensure empty string on error
End Function

Private Function PerformFilterCheck(dataRowArray As Variant, ByRef config As tConfigSettings) As Boolean
    ' 抽出されたデータ行がフィルター条件を満たすか確認します。(スタブ)
    ' 将来的にはconfig内の各種フィルター設定に基づいて判定します。
    PerformFilterCheck = True ' Placeholder, always passes filter
End Function

Public Function ExtractDataFromFile(kouteiFilePath As String, ByRef config As tConfigSettings, ByVal mainWorkbook As Workbook, Optional wsOutput As Worksheet = Nothing, Optional ByRef outputNextRow As Long = 0, Optional ByVal currentFileNum As Long = 0, Optional ByRef totalExtractedCount As Long = 0) As Boolean

    Static s_lastSuccessfullyProcessedFilePath As String
    Static s_lastValidYearInFileAsLong As Long
    Static s_lastValidMonthInFileAsLong As Long
    Dim anyDateExtractedSuccessfullyInFile As Boolean
    Dim yearMonthEstablishedForThisFile As Boolean

    Dim wbKoutei As Workbook, wsKoutei As Worksheet
    Dim currentYear As Long, currentMonth As Long, dayIdx As Long
    Dim dayVal As Variant, dateInLoop As Date
    Dim dayCellRow As Long, tempStr As String
    Dim eachSheetName As Variant
    Dim actualTargetSheetName As String
    Dim targetSheetProcessed As Boolean: targetSheetProcessed = False
    Dim yearVal As Variant, monthVal As Variant
    Dim yearStr As String, monthStr As String

    Dim outputActualHeaderNames() As String ' 0-based from Split
    Dim workerHeaderMap As Object ' Scripting.Dictionary
    Dim workerHeaderSequence As Long
    Dim sagyoinMasterIndex As Long
    Dim baseSagyoinOffset As tOffset
    Dim baseSagyoinOffsetIsEmpty As Boolean
    Dim oneRowOfExtractedData() As Variant ' 0-based
    Dim outputColIdx As Long, colIdxForWorkerMap As Long
    Dim currentOutputHeader As String
    Dim currentDataValue As Variant
    Dim masterIdx As Long
    Dim foundMasterOffset As Boolean
    Dim hasKeyDataOtherThanWorkers As Boolean
    Dim keyHeaderName As Variant
    Dim keyHeaderNamesForBlankCheck As Variant

    Dim processIdx As Long
    Dim currentBaseRowForProcess As Long, currentBaseColForProcess As Long
    Dim colOffsetAccumulator As Long
    Dim workerName As String, actualExtractedWorkerCount As Long
    Dim isRowAllEmpty As Boolean
    Dim maxWorkersForThisProcess As Long
    Dim i_m06 As Long ' Generic loop counter

    anyDateExtractedSuccessfullyInFile = False
    ExtractDataFromFile = False

    If s_lastSuccessfullyProcessedFilePath <> kouteiFilePath Then
        If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE: M06_DataExtractor.ExtractDataFromFile - New file detected. Resetting year/month fallback. Old Path: '" & s_lastSuccessfullyProcessedFilePath & "', New Path: '" & kouteiFilePath & "'"
        s_lastValidYearInFileAsLong = 0
        s_lastValidMonthInFileAsLong = 0
        s_lastSuccessfullyProcessedFilePath = kouteiFilePath
        yearMonthEstablishedForThisFile = False
    Else
        yearMonthEstablishedForThisFile = (s_lastValidYearInFileAsLong <> 0 And s_lastValidMonthInFileAsLong <> 0)
        If yearMonthEstablishedForThisFile Then
             currentYear = s_lastValidYearInFileAsLong
             currentMonth = s_lastValidMonthInFileAsLong
             If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE: M06_DataExtractor.ExtractDataFromFile - Using previously established Year/Month for file '" & kouteiFilePath & "': " & currentYear & "/" & currentMonth
        End If
    End If

    On Error GoTo ExtractDataFromFile_Error

    If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M06_DataExtractor.ExtractDataFromFile - Opening file: '" & kouteiFilePath & "'"
    Set wbKoutei = Workbooks.Open(Filename:=kouteiFilePath, UpdateLinks:=0, ReadOnly:=True)
    If wbKoutei Is Nothing Then GoTo ExtractDataFromFile_Finally

    ' --- I. Preparation and Header Analysis ---
    If Not (config.OutputHeaderRowCount > 0 And UBound(config.OutputHeaderContents) >= LBound(config.OutputHeaderContents) And config.OutputHeaderRowCount <= UBound(config.OutputHeaderContents) And config.OutputHeaderRowCount >= LBound(config.OutputHeaderContents)) Then
        Call M04_LogWriter.SafeWriteErrorLog("CRITICAL", mainWorkbook, config.ErrorLogSheetName, "M06_DataExtractor", "ExtractDataFromFile", "出力ヘッダーがConfig G-2で定義されていないか行数指定が不正なため、データマッピングできません。", 0, kouteiFilePath)
        ExtractDataFromFile = False
        GoTo ExtractDataFromFile_Finally
    End If
    outputActualHeaderNames = Split(config.OutputHeaderContents(config.OutputHeaderRowCount), vbTab) ' 0-based array

    If config.TraceDebugEnabled Then
        Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE: M06_DataExtractor.ExtractDataFromFile - Parsed OutputActualHeaderNames from config.OutputHeaderContents(" & config.OutputHeaderRowCount & "):"
        Dim tempHeadIdx_M06_Debug As Long ' Use a unique variable name for the loop counter

        ' Robust check for array state after Split
        Dim lBoundVal As Long, uBoundVal As Long
        Dim arrayIsValidForLoop As Boolean
        arrayIsValidForLoop = False
        On Error Resume Next ' Handle UBound on uninitialized array, though Split should always return an array
        lBoundVal = LBound(outputActualHeaderNames)
        uBoundVal = UBound(outputActualHeaderNames)
        If Err.Number = 0 Then ' Successfully got bounds
            On Error GoTo ExtractDataFromFile_Error ' Restore main error handler if it was active
            arrayIsValidForLoop = (uBoundVal >= lBoundVal)
        Else ' Failed to get bounds, array likely not what we expect
            On Error GoTo ExtractDataFromFile_Error ' Restore main error handler
            Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE:   outputActualHeaderNames is not a valid/initialized array after Split (Error getting bounds)."
            Err.Clear
        End If

        If arrayIsValidForLoop Then
            For tempHeadIdx_M06_Debug = lBoundVal To uBoundVal
                Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE:   Header Index " & tempHeadIdx_M06_Debug & " (" & lBoundVal & "-" & uBoundVal & "): [" & outputActualHeaderNames(tempHeadIdx_M06_Debug) & "]"
            Next tempHeadIdx_M06_Debug
        Else
            Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE:   outputActualHeaderNames array is empty or invalid after Split (LBound > UBound or error getting bounds)."
        End If
        Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE: --- End Parsed OutputActualHeaderNames ---"
    End If
    ' The original simple Join print is now covered by the loop above if TraceDebugEnabled.
    ' If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE: Output headers for mapping (last row of config headers): " & Join(outputActualHeaderNames, "|")


    Set workerHeaderMap = CreateObject("Scripting.Dictionary")
    workerHeaderSequence = 0
    If UBound(outputActualHeaderNames) >= LBound(outputActualHeaderNames) Then ' Check if array is valid
        For colIdxForWorkerMap = LBound(outputActualHeaderNames) To UBound(outputActualHeaderNames)
            Dim tempHeaderNameForWorkerCheck As String: tempHeaderNameForWorkerCheck = Trim(outputActualHeaderNames(colIdxForWorkerMap))
            If Left(tempHeaderNameForWorkerCheck, DEFAULT_WORKER_HEADER_PREFIX_LENGTH) = DEFAULT_WORKER_HEADER_PREFIX Then
                workerHeaderSequence = workerHeaderSequence + 1
                workerHeaderMap(tempHeaderNameForWorkerCheck) = workerHeaderSequence
            End If
        Next colIdxForWorkerMap
    End If
    If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE: Worker header map created with " & workerHeaderMap.Count & " items."

    sagyoinMasterIndex = -1
    baseSagyoinOffsetIsEmpty = True
    If UBound(config.OffsetItemMasterNames) >= LBound(config.OffsetItemMasterNames) Then
        For i_m06 = LBound(config.OffsetItemMasterNames) To UBound(config.OffsetItemMasterNames)
            If config.OffsetItemMasterNames(i_m06) = DEFAULT_WORKER_HEADER_PREFIX Then
                sagyoinMasterIndex = i_m06
                If sagyoinMasterIndex >= LBound(config.OffsetDefinitions) And sagyoinMasterIndex <= UBound(config.OffsetDefinitions) Then
                    baseSagyoinOffset = config.OffsetDefinitions(sagyoinMasterIndex)
                Else
                    baseSagyoinOffset.Row = 0: baseSagyoinOffset.Col = 0 ' Default
                    If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - WARNING: sagyoinMasterIndex out of bounds for OffsetDefinitions."
                End If
                If sagyoinMasterIndex >= LBound(config.IsOffsetOriginallyEmptyFlags) And sagyoinMasterIndex <= UBound(config.IsOffsetOriginallyEmptyFlags) Then
                    baseSagyoinOffsetIsEmpty = config.IsOffsetOriginallyEmptyFlags(sagyoinMasterIndex)
                Else
                    baseSagyoinOffsetIsEmpty = True ' Default
                    If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - WARNING: sagyoinMasterIndex out of bounds for IsOffsetOriginallyEmptyFlags."
                End If
                If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE: Base sagyoin offset '" & DEFAULT_WORKER_HEADER_PREFIX & "' found at master index " & sagyoinMasterIndex & ", R=" & baseSagyoinOffset.Row & ", C=" & baseSagyoinOffset.Col & ", IsEmptyOrig=" & baseSagyoinOffsetIsEmpty
                Exit For
            End If
        Next i_m06
    End If

    If sagyoinMasterIndex = -1 Then
        If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE: M06_DataExtractor.ExtractDataFromFile - WARNING: Base '" & DEFAULT_WORKER_HEADER_PREFIX & "' offset not found in Config Section F. Worker names cannot be extracted by offset."
    ElseIf baseSagyoinOffsetIsEmpty Then
        If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE: M06_DataExtractor.ExtractDataFromFile - WARNING: Base '" & DEFAULT_WORKER_HEADER_PREFIX & "' offset is defined but its value is empty. Worker names cannot be extracted by offset."
    End If

    ' --- 対象シートループ ---
    If UBound(config.TargetSheetNames) >= LBound(config.TargetSheetNames) Then
        For Each eachSheetName In config.TargetSheetNames
            actualTargetSheetName = Trim(CStr(eachSheetName))
            targetSheetProcessed = True

            If Len(actualTargetSheetName) = 0 Then
                If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE: M06_DataExtractor.ExtractDataFromFile - Empty sheet name found in TargetSheetNames array. Skipping."
                GoTo NextSheetInLoop_M06
            End If

            If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE: M06_DataExtractor.ExtractDataFromFile - Attempting to process sheet: '" & actualTargetSheetName & "' in file '" & kouteiFilePath & "'"

            Set wsKoutei = Nothing
            On Error Resume Next
            Set wsKoutei = wbKoutei.Worksheets(actualTargetSheetName)
            On Error GoTo ExtractDataFromFile_Error

            If wsKoutei Is Nothing Then
                tempStr = "シートが見つかりません: '" & actualTargetSheetName & "'"
                Call M04_LogWriter.SafeWriteErrorLog("ERROR", mainWorkbook, config.ErrorLogSheetName, "M06_DataExtractor", "ExtractDataFromFile (SheetCheck)", tempStr, 0, kouteiFilePath)
                GoTo NextSheetInLoop_M06
            End If

            If yearMonthEstablishedForThisFile Then
                If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE: M06_DataExtractor.ExtractDataFromFile - Sheet '" & actualTargetSheetName & "': Using pre-established Year/Month: " & currentYear & "/" & currentMonth
            Else
                Dim cellAccessErr As Boolean: cellAccessErr = False
                On Error Resume Next
                yearVal = wsKoutei.Range(config.YearCellAddress).Value
                If Err.Number <> 0 Then cellAccessErr = True: tempStr = "年セル(" & config.YearCellAddress & ")アクセスエラー: " & Err.Description: Err.Clear

                monthVal = wsKoutei.Range(config.MonthCellAddress).Value
                If Err.Number <> 0 And Not cellAccessErr Then cellAccessErr = True: tempStr = "月セル(" & config.MonthCellAddress & ")アクセスエラー: " & Err.Description: Err.Clear
                On Error GoTo ExtractDataFromFile_Error

                If cellAccessErr Then
                    If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE: M06_DataExtractor.ExtractDataFromFile - Sheet '" & actualTargetSheetName & "': " & tempStr
                    Call M04_LogWriter.WriteFilterLogEntry(config.LogSheetName, GetNextFilterLogRow(config.LogSheetName, mainWorkbook), "年月取得失敗(セルアクセス)", kouteiFilePath & "/" & actualTargetSheetName & "/" & tempStr)
                    GoTo NextSheetInLoop_M06
                End If

                If IsError(yearVal) Then yearVal = Empty
                If IsError(monthVal) Then monthVal = Empty

                yearStr = Trim(CStr(yearVal))
                monthStr = Trim(CStr(monthVal))
                Dim tempYearMonthValid As Boolean: tempYearMonthValid = False

                If Len(yearStr) > 0 And IsNumeric(yearStr) And CLng(yearStr) >= 1900 And CLng(yearStr) <= 2999 Then
                    If Len(monthStr) > 0 And IsNumeric(monthStr) And CLng(monthStr) >= 1 And CLng(monthStr) <= 12 Then
                        currentYear = CLng(yearStr)
                        currentMonth = CLng(monthStr)
                        s_lastValidYearInFileAsLong = currentYear
                        s_lastValidMonthInFileAsLong = currentMonth
                        yearMonthEstablishedForThisFile = True
                        tempYearMonthValid = True
                        tempStr = "ファイル「" & kouteiFilePath & "」の年/月を " & currentYear & "/" & currentMonth & " に確定 (シート「" & actualTargetSheetName & "」より取得)"
                        Call M04_LogWriter.WriteFilterLogEntry(config.LogSheetName, GetNextFilterLogRow(config.LogSheetName, mainWorkbook), "年月確定", tempStr)
                        If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE: M06_DataExtractor.ExtractDataFromFile - " & tempStr
                    End If
                End If

                If Not tempYearMonthValid Then
                    tempStr = "シート「" & actualTargetSheetName & "」の年/月セルの値が不正です。 Y (" & config.YearCellAddress & "):'" & CStr(yearVal) & "', M (" & config.MonthCellAddress & "):'" & CStr(monthVal) & "'"
                    If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE: M06_DataExtractor.ExtractDataFromFile - " & tempStr
                    Call M04_LogWriter.WriteFilterLogEntry(config.LogSheetName, GetNextFilterLogRow(config.LogSheetName, mainWorkbook), "年月取得失敗(値不正)", kouteiFilePath & "/" & actualTargetSheetName & "/" & tempStr)
                    GoTo NextSheetInLoop_M06
                End If
            End If

            If Not yearMonthEstablishedForThisFile Then
                If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE: M06_DataExtractor.ExtractDataFromFile - Sheet '" & actualTargetSheetName & "': Year/Month not established, skipping day processing."
                GoTo NextSheetInLoop_M06
            End If

            For dayIdx = 1 To config.MaxDaysPerSheet
                dayCellRow = config.HeaderRowCount + (dayIdx - 1) * config.RowsPerDay + config.DayRowOffset
                dayVal = wsKoutei.Cells(dayCellRow, config.DayColumnLetter).Value

                If IsEmpty(dayVal) Or Len(Trim(CStr(dayVal))) = 0 Then
                    If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE: M06_DataExtractor.ExtractDataFromFile - Day value at " & wsKoutei.Name & "!" & config.DayColumnLetter & dayCellRow & " is empty. Skipping."
                    GoTo NextDayInLoop_M06
                End If

                If Not IsNumeric(dayVal) Then
                    tempStr = "日付セルの値が数値ではありません (" & config.DayColumnLetter & dayCellRow & "): '" & CStr(dayVal) & "'"
                    If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE: M06_DataExtractor.ExtractDataFromFile - " & tempStr & ". Skipping."
                    Call M04_LogWriter.WriteFilterLogEntry(config.LogSheetName, GetNextFilterLogRow(config.LogSheetName, mainWorkbook), "日付取得失敗(非数値)", kouteiFilePath & "/" & actualTargetSheetName & "/" & tempStr)
                    GoTo NextDayInLoop_M06
                End If

                Dim dayLong As Long
                dayLong = CLng(dayVal)

                If dayLong <= 0 Or dayLong > 31 Then
                    tempStr = "日付セルの値が範囲外(1-31)です (" & config.DayColumnLetter & dayCellRow & "): " & dayLong
                    If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE: M06_DataExtractor.ExtractDataFromFile - " & tempStr & ". Skipping."
                    Call M04_LogWriter.WriteFilterLogEntry(config.LogSheetName, GetNextFilterLogRow(config.LogSheetName, mainWorkbook), "日付取得失敗(範囲外)", kouteiFilePath & "/" & actualTargetSheetName & "/" & tempStr)
                    GoTo NextDayInLoop_M06
                End If

                On Error Resume Next
                dateInLoop = DateSerial(currentYear, currentMonth, dayLong)
                If Err.Number <> 0 Then
                    On Error GoTo ExtractDataFromFile_Error
                    tempStr = "DateSerialで無効な日付です (" & currentYear & "/" & currentMonth & "/" & dayLong & " at " & config.DayColumnLetter & dayCellRow & "). Error: " & Err.Description
                    If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE: M06_DataExtractor.ExtractDataFromFile - " & tempStr & ". Skipping."
                    Call M04_LogWriter.WriteFilterLogEntry(config.LogSheetName, GetNextFilterLogRow(config.LogSheetName, mainWorkbook), "日付検証エラー(DateSerial)", kouteiFilePath & "/" & actualTargetSheetName & "/" & tempStr)
                    Err.Clear
                    GoTo NextDayInLoop_M06
                End If
                On Error GoTo ExtractDataFromFile_Error

                For processIdx = 0 To config.ProcessesPerDay - 1
                    If UBound(outputActualHeaderNames) >= LBound(outputActualHeaderNames) Then ' Check if outputActualHeaderNames is populated
                        ReDim oneRowOfExtractedData(LBound(outputActualHeaderNames) To UBound(outputActualHeaderNames)) '0-based
                    Else
                        Call M04_LogWriter.SafeWriteErrorLog("CRITICAL", mainWorkbook, config.ErrorLogSheetName, "M06_DataExtractor", "ExtractDataFromFile", "outputActualHeaderNames が初期化されていないか空です。致命的なエラー。", 0, kouteiFilePath & "/" & actualTargetSheetName)
                        ExtractDataFromFile = False
                        GoTo ExtractDataFromFile_Finally ' Critical error, cannot proceed
                    End If

                    actualExtractedWorkerCount = 0
                    colOffsetAccumulator = 0
                    If processIdx > 0 Then
                        Dim k As Long
                        For k = 0 To processIdx - 1
                            If UBound(config.ProcessPatternColNumbers) >= 1 And LBound(config.ProcessPatternColNumbers) <= 1 Then ' Check outer array
                                If UBound(config.ProcessPatternColNumbers(1)) >= k And LBound(config.ProcessPatternColNumbers(1)) <= k Then ' Check inner array
                                    colOffsetAccumulator = colOffsetAccumulator + config.ProcessPatternColNumbers(1)(k)
                                Else
                                    If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - WARNING: ProcessPatternColNumbers(1) inner array not valid for k=" & k
                                End If
                            Else
                                If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - WARNING: ProcessPatternColNumbers outer array not valid"
                            End If
                        Next k
                    End If
                    currentBaseRowForProcess = config.HeaderRowCount + (dayIdx - 1) * config.RowsPerDay + 1
                    currentBaseColForProcess = config.HeaderColCount + 1 + colOffsetAccumulator

                    If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE: Processing Day " & Format(dateInLoop, "yyyy-mm-dd") & ", ProcessIdx " & processIdx & ", BaseCell R" & currentBaseRowForProcess & "C" & currentBaseColForProcess

                    For outputColIdx = LBound(outputActualHeaderNames) To UBound(outputActualHeaderNames) ' 0-based
                        currentOutputHeader = Trim(outputActualHeaderNames(outputColIdx))
                        currentDataValue = ""

                        If currentOutputHeader = "日付" Then
                            currentDataValue = Format(dateInLoop, "yyyy/mm/dd(aaa)")
                        ElseIf currentOutputHeader = "年" Then
                            currentDataValue = currentYear
                        ElseIf currentOutputHeader = "月" Then
                            currentDataValue = currentMonth
                        ElseIf currentOutputHeader = "シート名" Then
                            currentDataValue = actualTargetSheetName
                        ElseIf currentOutputHeader = "管内1" Then
                            If config.ProcessesPerDay > 0 And UBound(config.ProcessDetails) >= LBound(config.ProcessDetails) And processIdx >= LBound(config.ProcessDetails) And processIdx <= UBound(config.ProcessDetails) Then
                                currentDataValue = config.ProcessDetails(processIdx).Kankatsu1
                            End If
                        ElseIf currentOutputHeader = "管内2" Then
                             If config.ProcessesPerDay > 0 And UBound(config.ProcessDetails) >= LBound(config.ProcessDetails) And processIdx >= LBound(config.ProcessDetails) And processIdx <= UBound(config.ProcessDetails) Then
                                currentDataValue = config.ProcessDetails(processIdx).Kankatsu2
                            End If
                        ElseIf workerHeaderMap.Exists(currentOutputHeader) Then
                            Dim workerSequenceNum As Long: workerSequenceNum = workerHeaderMap(currentOutputHeader)
                            If sagyoinMasterIndex <> -1 And Not baseSagyoinOffsetIsEmpty Then
                                maxWorkersForThisProcess = 0
                                If UBound(config.ProcessPatternColNumbers) >= 1 And LBound(config.ProcessPatternColNumbers) <= 1 Then
                                    If UBound(config.ProcessPatternColNumbers(1)) >= processIdx And LBound(config.ProcessPatternColNumbers(1)) <= processIdx Then
                                        maxWorkersForThisProcess = config.ProcessPatternColNumbers(1)(processIdx)
                                    End If
                                End If

                                If workerSequenceNum > 0 And workerSequenceNum <= Application.WorksheetFunction.Min(MAX_WORKERS_TO_EXTRACT, maxWorkersForThisProcess) Then
                                    Dim sagyoinActualOffset As tOffset
                                    sagyoinActualOffset.Row = baseSagyoinOffset.Row
                                    sagyoinActualOffset.Col = baseSagyoinOffset.Col + (workerSequenceNum - 1)
                                    currentDataValue = GetValueFromOffset(wsKoutei, currentBaseRowForProcess, currentBaseColForProcess, sagyoinActualOffset, currentOutputHeader, config, mainWorkbook)
                                    If Len(CStr(currentDataValue)) > 0 Then actualExtractedWorkerCount = actualExtractedWorkerCount + 1
                                End If
                            End If
                        Else
                            foundMasterOffset = False
                            If UBound(config.OffsetItemMasterNames) >= LBound(config.OffsetItemMasterNames) Then
                                For masterIdx = LBound(config.OffsetItemMasterNames) To UBound(config.OffsetItemMasterNames)
                                    If config.OffsetItemMasterNames(masterIdx) = currentOutputHeader Then
                                        If UBound(config.IsOffsetOriginallyEmptyFlags) >= masterIdx And LBound(config.IsOffsetOriginallyEmptyFlags) <= masterIdx Then
                                            If Not config.IsOffsetOriginallyEmptyFlags(masterIdx) Then
                                                If UBound(config.OffsetDefinitions) >= masterIdx And LBound(config.OffsetDefinitions) <= masterIdx Then
                                                    currentDataValue = GetValueFromOffset(wsKoutei, currentBaseRowForProcess, currentBaseColForProcess, config.OffsetDefinitions(masterIdx), currentOutputHeader, config, mainWorkbook)
                                                End If
                                            End If
                                        End If
                                        foundMasterOffset = True
                                        Exit For
                                    End If
                                Next masterIdx
                            End If
                            If Not foundMasterOffset And config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE: M06_DataExtractor.ExtractDataFromFile - Header '" & currentOutputHeader & "' not found in OffsetItemMasterNames or fixed fields for ProcessIdx " & processIdx & ", Day " & Format(dateInLoop, "dd")
                        End If
                        oneRowOfExtractedData(outputColIdx) = currentDataValue
                    Next outputColIdx

                    hasKeyDataOtherThanWorkers = False
                    keyHeaderNamesForBlankCheck = Array("工番", "変電所", "作業名1", "作業名2", "担当の名前", "工事種類", "人数")

                    If UBound(outputActualHeaderNames) >= LBound(outputActualHeaderNames) And UBound(oneRowOfExtractedData) >= LBound(oneRowOfExtractedData) Then
                        For outputColIdx = LBound(outputActualHeaderNames) To UBound(outputActualHeaderNames)
                            currentOutputHeader = outputActualHeaderNames(outputColIdx)
                            Dim isKeyToCheck As Boolean: isKeyToCheck = False
                            For Each keyHeaderName In keyHeaderNamesForBlankCheck
                                If currentOutputHeader = CStr(keyHeaderName) Then
                                    isKeyToCheck = True
                                    Exit For
                                End If
                            Next keyHeaderName

                            If isKeyToCheck Then
                                If Len(CStr(oneRowOfExtractedData(outputColIdx))) > 0 Then
                                    hasKeyDataOtherThanWorkers = True
                                    Exit For
                                End If
                            End If
                        Next outputColIdx
                    End If

                    If config.ExtractIfWorkersEmpty Then
                        isRowAllEmpty = Not hasKeyDataOtherThanWorkers
                    Else
                        isRowAllEmpty = Not (hasKeyDataOtherThanWorkers And actualExtractedWorkerCount > 0)
                    End If

                    If isRowAllEmpty Then
                        Call M04_LogWriter.WriteFilterLogEntry(config.LogSheetName, GetNextFilterLogRow(config.LogSheetName, mainWorkbook), "空白行スキップ(O241:" & config.ExtractIfWorkersEmpty & ")", kouteiFilePath & "/" & actualTargetSheetName & "/Day" & Format(dateInLoop, "dd") & "/Proc" & processIdx)
                        GoTo NextProcessInDay_M06
                    End If

                    If Not PerformFilterCheck(oneRowOfExtractedData, config) Then GoTo NextProcessInDay_M06

                    If Not wsOutput Is Nothing Then
                        For outputColIdx = LBound(oneRowOfExtractedData) To UBound(oneRowOfExtractedData) ' 0-based
                            wsOutput.Cells(outputNextRow, outputColIdx + 1).Value = oneRowOfExtractedData(outputColIdx) ' Write to 1-based Excel columns
                        Next outputColIdx
                        outputNextRow = outputNextRow + 1
                        totalExtractedCount = totalExtractedCount + 1
                        anyDateExtractedSuccessfullyInFile = True
                    Else
                        If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE: M06_DataExtractor.ExtractDataFromFile - wsOutput is Nothing. Cannot write extracted data row."
                        anyDateExtractedSuccessfullyInFile = True
                    End If

                    tempStr = kouteiFilePath & "/" & actualTargetSheetName & "/" & Format(dateInLoop, "yyyy-mm-dd") & "/Proc" & processIdx
                    Call M04_LogWriter.WriteFilterLogEntry(config.LogSheetName, GetNextFilterLogRow(config.LogSheetName, mainWorkbook), "行抽出成功", tempStr)
                    If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE: M06_DataExtractor.ExtractDataFromFile - Successfully extracted row for Date: " & Format(dateInLoop, "yyyy-mm-dd") & ", Process: " & processIdx

        NextProcessInDay_M06:
                Next processIdx
NextDayInLoop_M06:
            Next dayIdx
NextSheetInLoop_M06:
        Next eachSheetName

        If Not targetSheetProcessed And UBound(config.TargetSheetNames) >= LBound(config.TargetSheetNames) Then
             Call M04_LogWriter.SafeWriteErrorLog("WARNING", mainWorkbook, config.ErrorLogSheetName, "M06_DataExtractor", "ExtractDataFromFile", "処理対象シート名リスト(config.TargetSheetNames)に有効なシート名がありませんでした。", 0, kouteiFilePath)
             GoTo ExtractDataFromFile_Finally
        End If
    Else
        Call M04_LogWriter.SafeWriteErrorLog("ERROR", mainWorkbook, config.ErrorLogSheetName, "M06_DataExtractor", "ExtractDataFromFile", "処理対象シート名リスト(config.TargetSheetNames)が空または未初期化です。", 0, kouteiFilePath)
        GoTo ExtractDataFromFile_Finally
    End If

    If Not yearMonthEstablishedForThisFile And targetSheetProcessed Then
        tempStr = "ファイル「" & kouteiFilePath & "」内のどの指定シートからも有効な年/月を取得できませんでした。"
        Call M04_LogWriter.SafeWriteErrorLog("ERROR", mainWorkbook, config.ErrorLogSheetName, "M06_DataExtractor", "ExtractDataFromFile (YearMonthEstablishment)", tempStr, 0, "ファイル処理エラー")
        anyDateExtractedSuccessfullyInFile = False
    End If
    ExtractDataFromFile = anyDateExtractedSuccessfullyInFile

ExtractDataFromFile_Finally:
    If Not wbKoutei Is Nothing Then wbKoutei.Close SaveChanges:=False
    Set wsKoutei = Nothing
    Set wbKoutei = Nothing
    If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M06_DataExtractor.ExtractDataFromFile - Closed file: '" & kouteiFilePath & "', Result: " & ExtractDataFromFile
    Exit Function

ExtractDataFromFile_Error:
    tempStr = "実行時エラー " & Err.Number & ": " & Err.Description & ", Procedure: ExtractDataFromFile, File: " & kouteiFilePath
    Call M04_LogWriter.SafeWriteErrorLog("ERROR", mainWorkbook, config.ErrorLogSheetName, "M06_DataExtractor", "ExtractDataFromFile", tempStr, Err.Number, Err.Description)
    ExtractDataFromFile = False
    Resume ExtractDataFromFile_Finally
End Function
