' バージョン：v0.5.0
Option Explicit
' このモジュールは、個々の工程表ファイルを開き、指定されたシートからデータを抽出し、フィルター条件を適用（将来的に）し、結果を出力シートに書き込む（将来的に）処理を担当します。このステップでは、年月と日付の基本情報抽出とログ出力を実装します。

Private Function LogExtractor_IsArrayInitialized(arr As Variant) As Boolean
    ' 配列が有効に初期化されているか（少なくとも1つの要素を持つか）を確認します。
    ' Variant型が配列でない場合、または配列であっても要素が割り当てられていない場合はFalseを返します。
    On Error GoTo NotAnArrayOrNotInitialized_M06
    If IsArray(arr) Then
        Dim lBoundCheck As Long
        lBoundCheck = LBound(arr)
        LogExtractor_IsArrayInitialized = True
        Exit Function
    End If
NotAnArrayOrNotInitialized_M06:
    LogExtractor_IsArrayInitialized = False
End Function

Private Function GetNextFilterLogRow(filterLogSheet As Worksheet) As Long
    ' 指定されたフィルターログシートのA列で、次に書き込むべき行番号を取得します。
    If filterLogSheet Is Nothing Then
        GetNextFilterLogRow = 1 ' フォールバック
        Exit Function
    End If
    If Application.WorksheetFunction.CountA(filterLogSheet.Columns(1)) = 0 Then
        GetNextFilterLogRow = 1
    Else
        GetNextFilterLogRow = filterLogSheet.Cells(filterLogSheet.Rows.Count, 1).End(xlUp).Row + 1
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
    ' 個別工程表ファイルからデータを抽出し、フィルターログに年月日の基本情報を記録します。
    ' Arguments:
    '   kouteiFilePath: (I) String型。処理対象の工程表ファイルパス。
    '   config: (I) tConfigSettings型。各種設定情報。
    '   mainWorkbook: (I) Workbook型。マクロ本体（ログシートが存在する）のワークブックオブジェクト。
    '   wsOutput: (I/O) Worksheet型 (Optional)。抽出データの出力先シート。
    '   outputNextRow: (I/O) Long型 (Optional)。出力シートの次の書き込み行。
    '   currentFileNum: (I) Long型 (Optional)。現在処理中のファイル番号 (ログ用)。
    '   totalExtractedCount: (I/O) Long型 (Optional)。抽出成功した総件数。
    ' Returns:
    '   Boolean: データ抽出処理が（部分的にでも）成功した場合はTrue、それ以外はFalse。

    Static s_lastSuccessfullyProcessedFilePath As String
    Static s_lastValidYearInFileAsLong As Long ' Renamed
    Static s_lastValidMonthInFileAsLong As Long ' Renamed
    Dim anyDateExtractedSuccessfullyInFile As Boolean ' For function return value
    Dim yearMonthEstablishedForThisFile As Boolean ' New local variable

    Dim wbKoutei As Workbook, wsKoutei As Worksheet
    Dim currentYear As Long, currentMonth As Long, dayIdx As Long
    Dim dayVal As Variant, dateInLoop As Date 
    Dim dayCellRow As Long, tempStr As String
    Dim filterLogSht As Worksheet
    Dim eachSheetName As Variant ' For iterating through target sheets
    Dim actualTargetSheetName As String ' To hold the trimmed sheet name
    Dim targetSheetProcessed As Boolean: targetSheetProcessed = False ' Flag to see if at least one sheet was attempted
    Dim yearVal As Variant, monthVal As Variant 
    Dim yearStr As String, monthStr As String
    ' Dim yearMonthValid As Boolean ' Replaced by yearMonthEstablishedForThisFile at a higher scope

    ' Constants for extractedData array
    Const COL_DATE = 1, COL_YEAR = 2, COL_MONTH = 3, COL_SHEETNAME = 4
    Const COL_KANKATSU1 = 5, COL_KANKATSU2 = 6
    Const COL_BUNRUI1 = 7, COL_BUNRUI2 = 8, COL_BUNRUI3 = 9 ' Placeholders, as Bunrui not in ProcessDetails yet
    Const COL_KOUBAN = 10, COL_HENSHOUSHO = 11, COL_SAGYOMEI1 = 12, COL_SAGYOMEI2 = 13
    Const COL_TANTOU = 14, COL_KOUJISHURUI = 15, COL_NINZU = 16
    Const COL_SAGYOIN_START = 17, COL_SAGYOIN_END = 26 ' For 10 workers
    Const COL_SONOTA = 27, COL_SHUURYOJIKAN = 28, COL_BUNRUI1EXTSRC = 29
    Const MAX_EXTRACT_COLS = 29 
    Dim extractedData(1 To MAX_EXTRACT_COLS) As Variant
    Dim processIdx As Long
    Dim currentBaseRowForProcess As Long, currentBaseColForProcess As Long
    Dim colOffsetAccumulator As Long
    Dim workerIdx As Long, workerActualCol As Long, workerName As String, actualExtractedWorkerCount As Long
    Dim isRowAllEmpty As Boolean
    Dim maxWorkersForThisProcess As Long

    anyDateExtractedSuccessfullyInFile = False
    ExtractDataFromFile = False ' Default to failure

    If s_lastSuccessfullyProcessedFilePath <> kouteiFilePath Then
        If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE: M06_DataExtractor.ExtractDataFromFile - New file detected. Resetting year/month fallback. Old Path: '" & s_lastSuccessfullyProcessedFilePath & "', New Path: '" & kouteiFilePath & "'"
        s_lastValidYearInFileAsLong = 0 ' Reset for new file
        s_lastValidMonthInFileAsLong = 0
        s_lastSuccessfullyProcessedFilePath = kouteiFilePath
        yearMonthEstablishedForThisFile = False ' Reset for new file
    Else
        ' Same file as last call in this macro run, check if Y/M was already found
        yearMonthEstablishedForThisFile = (s_lastValidYearInFileAsLong <> 0 And s_lastValidMonthInFileAsLong <> 0)
        If yearMonthEstablishedForThisFile Then
             currentYear = s_lastValidYearInFileAsLong ' Use previously established Y/M
             currentMonth = s_lastValidMonthInFileAsLong
             If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE: M06_DataExtractor.ExtractDataFromFile - Using previously established Year/Month for file '" & kouteiFilePath & "': " & currentYear & "/" & currentMonth
        End If
    End If
    
    On Error GoTo ExtractDataFromFile_Error ' Main error handler for the function

    If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M06_DataExtractor.ExtractDataFromFile - Opening file: '" & kouteiFilePath & "'"
    Set wbKoutei = Workbooks.Open(Filename:=kouteiFilePath, UpdateLinks:=0, ReadOnly:=True)
    If wbKoutei Is Nothing Then Exit Function ' Should be caught by error handler, but as a safeguard

    ' --- フィルターログシート取得 (最初に取得試行) ---
    On Error Resume Next ' Temporarily disable error handling for sheet existence check
    Set filterLogSht = mainWorkbook.Worksheets(config.SearchConditionLogSheetName)
    On Error GoTo ExtractDataFromFile_Error ' Restore main error handler
    
    If filterLogSht Is Nothing Then
        If DEBUG_MODE_ERROR Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - ERROR: M06_DataExtractor.ExtractDataFromFile - Filter log sheet '" & config.SearchConditionLogSheetName & "' not found. Date extraction logging will be skipped."
        ' Proceeding without filter logging for this file if sheet is missing, but error log should have info from PrepareSheets
    End If

    ' --- 対象シートループ ---
    If LogExtractor_IsArrayInitialized(config.TargetSheetNames) And UBound(config.TargetSheetNames) >= LBound(config.TargetSheetNames) Then
        For Each eachSheetName In config.TargetSheetNames
            actualTargetSheetName = Trim(CStr(eachSheetName))
            targetSheetProcessed = True ' Mark that we are attempting to process a sheet

            If Len(actualTargetSheetName) = 0 Then
                If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE: M06_DataExtractor.ExtractDataFromFile - Empty sheet name found in TargetSheetNames array. Skipping."
                GoTo NextSheetInLoop_M06 ' Skip if sheet name is empty
            End If

            If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE: M06_DataExtractor.ExtractDataFromFile - Attempting to process sheet: '" & actualTargetSheetName & "' in file '" & kouteiFilePath & "'"
            
            Set wsKoutei = Nothing ' Reset for current sheet
            On Error Resume Next
            Set wsKoutei = wbKoutei.Worksheets(actualTargetSheetName)
            On Error GoTo ExtractDataFromFile_Error ' Restore main error handler
            
            If wsKoutei Is Nothing Then
                tempStr = "シートが見つかりません: '" & actualTargetSheetName & "'"
                Call M04_LogWriter.SafeWriteErrorLog("ERROR", mainWorkbook, config.ErrorLogSheetName, "M06_DataExtractor", "ExtractDataFromFile (SheetCheck)", tempStr, 0, kouteiFilePath)
                GoTo NextSheetInLoop_M06 ' Skip to next sheet
            End If

            If yearMonthEstablishedForThisFile Then
                ' currentYear and currentMonth are already set from static vars for this file
                If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE: M06_DataExtractor.ExtractDataFromFile - Sheet '" & actualTargetSheetName & "': Using pre-established Year/Month: " & currentYear & "/" & currentMonth
            Else
                ' Attempt to read and establish Year/Month from this sheet
                Dim cellAccessErr As Boolean: cellAccessErr = False
                On Error Resume Next ' For reading cell values specifically
                yearVal = wsKoutei.Range(config.YearCellAddress).Value
                If Err.Number <> 0 Then cellAccessErr = True: tempStr = "年セル(" & config.YearCellAddress & ")アクセスエラー: " & Err.Description: Err.Clear
                
                monthVal = wsKoutei.Range(config.MonthCellAddress).Value
                If Err.Number <> 0 And Not cellAccessErr Then cellAccessErr = True: tempStr = "月セル(" & config.MonthCellAddress & ")アクセスエラー: " & Err.Description: Err.Clear
                On Error GoTo ExtractDataFromFile_Error ' Restore main error handler

                If cellAccessErr Then
                    If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE: M06_DataExtractor.ExtractDataFromFile - Sheet '" & actualTargetSheetName & "': " & tempStr
                    If Not filterLogSht Is Nothing Then Call M04_LogWriter.WriteFilterLogEntry(filterLogSht, GetNextFilterLogRow(filterLogSht), "年月取得失敗(セルアクセス)", kouteiFilePath & "/" & actualTargetSheetName & "/" & tempStr)
                    GoTo NextSheetInLoop_M06 ' Try next sheet
                End If
                
                If IsError(yearVal) Then yearVal = Empty ' Convert Error type to Empty for consistent checks
                If IsError(monthVal) Then monthVal = Empty

                yearStr = Trim(CStr(yearVal))
                monthStr = Trim(CStr(monthVal))
                Dim tempYearMonthValid As Boolean: tempYearMonthValid = False

                If Len(yearStr) > 0 And IsNumeric(yearStr) And CLng(yearStr) >= 1900 And CLng(yearStr) <= 2999 Then
                    If Len(monthStr) > 0 And IsNumeric(monthStr) And CLng(monthStr) >= 1 And CLng(monthStr) <= 12 Then
                        currentYear = CLng(yearStr)
                        currentMonth = CLng(monthStr)
                        s_lastValidYearInFileAsLong = currentYear ' Update static vars
                        s_lastValidMonthInFileAsLong = currentMonth
                        yearMonthEstablishedForThisFile = True ' Mark as established for this file
                        tempYearMonthValid = True
                        tempStr = "ファイル「" & kouteiFilePath & "」の年/月を " & currentYear & "/" & currentMonth & " に確定 (シート「" & actualTargetSheetName & "」より取得)"
                        If Not filterLogSht Is Nothing Then Call M04_LogWriter.WriteFilterLogEntry(filterLogSht, GetNextFilterLogRow(filterLogSht), "年月確定", tempStr)
                        If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE: M06_DataExtractor.ExtractDataFromFile - " & tempStr
                    End If
                End If

                If Not tempYearMonthValid Then
                    tempStr = "シート「" & actualTargetSheetName & "」の年/月セルの値が不正です。 Y (" & config.YearCellAddress & "):'" & CStr(yearVal) & "', M (" & config.MonthCellAddress & "):'" & CStr(monthVal) & "'"
                    If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE: M06_DataExtractor.ExtractDataFromFile - " & tempStr
                    If Not filterLogSht Is Nothing Then Call M04_LogWriter.WriteFilterLogEntry(filterLogSht, GetNextFilterLogRow(filterLogSht), "年月取得失敗(値不正)", kouteiFilePath & "/" & actualTargetSheetName & "/" & tempStr)
                    GoTo NextSheetInLoop_M06 ' Try next sheet
                End If
            End If
            
            ' --- 日処理ループ ---
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
                    If Not filterLogSht Is Nothing Then Call M04_LogWriter.WriteFilterLogEntry(filterLogSht, GetNextFilterLogRow(filterLogSht), "日付取得失敗(非数値)", kouteiFilePath & "/" & actualTargetSheetName & "/" & tempStr)
                    GoTo NextDayInLoop_M06
                End If

                Dim dayLong As Long
                dayLong = CLng(dayVal) ' Convert to Long once

                If dayLong <= 0 Or dayLong > 31 Then ' Basic day range check
                    tempStr = "日付セルの値が範囲外(1-31)です (" & config.DayColumnLetter & dayCellRow & "): " & dayLong
                    If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE: M06_DataExtractor.ExtractDataFromFile - " & tempStr & ". Skipping."
                    If Not filterLogSht Is Nothing Then Call M04_LogWriter.WriteFilterLogEntry(filterLogSht, GetNextFilterLogRow(filterLogSht), "日付取得失敗(範囲外)", kouteiFilePath & "/" & actualTargetSheetName & "/" & tempStr)
                    GoTo NextDayInLoop_M06
                End If

                On Error Resume Next ' For DateSerial error only
                dateInLoop = DateSerial(currentYear, currentMonth, dayLong)
                If Err.Number <> 0 Then
                    On Error GoTo ExtractDataFromFile_Error ' Restore main error handler
                    tempStr = "DateSerialで無効な日付です (" & currentYear & "/" & currentMonth & "/" & dayLong & " at " & config.DayColumnLetter & dayCellRow & "). Error: " & Err.Description
                    If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE: M06_DataExtractor.ExtractDataFromFile - " & tempStr & ". Skipping."
                    If Not filterLogSht Is Nothing Then Call M04_LogWriter.WriteFilterLogEntry(filterLogSht, GetNextFilterLogRow(filterLogSht), "日付検証エラー(DateSerial)", kouteiFilePath & "/" & actualTargetSheetName & "/" & tempStr)
                    Err.Clear
                    GoTo NextDayInLoop_M06 
                End If
                On Error GoTo ExtractDataFromFile_Error ' Restore main error handler
                
                ' --- Process Loop for each process block within the current day ---
                For processIdx = 0 To config.ProcessesPerDay - 1
                    Erase extractedData ' Clear for current process block
                    colOffsetAccumulator = 0
                    If processIdx > 0 Then
                        Dim k As Long
                        For k = 0 To processIdx - 1
                            colOffsetAccumulator = colOffsetAccumulator + config.ProcessPatternColNumbers(1)(k)
                        Next k
                    End If
                    currentBaseRowForProcess = config.HeaderRowCount + (dayIdx - 1) * config.RowsPerDay + 1
                    currentBaseColForProcess = config.HeaderColCount + 1 + colOffsetAccumulator

                    If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE: Processing Day " & Format(dateInLoop, "yyyy-mm-dd") & ", ProcessIdx " & processIdx & ", BaseCell R" & currentBaseRowForProcess & "C" & currentBaseColForProcess

                    ' Populate fixed data
                    extractedData(COL_DATE) = Format(dateInLoop, "yyyy/mm/dd(aaa)")
                    extractedData(COL_YEAR) = currentYear
                    extractedData(COL_MONTH) = currentMonth
                    extractedData(COL_SHEETNAME) = actualTargetSheetName
                    If config.ProcessesPerDay > 0 And processIdx <= UBound(config.ProcessDetails) Then ' Check bounds
                        extractedData(COL_KANKATSU1) = config.ProcessDetails(processIdx).Kankatsu1
                        extractedData(COL_KANKATSU2) = config.ProcessDetails(processIdx).Kankatsu2
                    End If
                    ' Bunrui 1-3 currently not in ProcessDetails, so extractedData(COL_BUNRUI1-3) remain empty for now

                    ' Populate data using offsets
                    If Not config.IsOffsetKoubanOriginallyEmpty Then extractedData(COL_KOUBAN) = GetValueFromOffset(wsKoutei, currentBaseRowForProcess, currentBaseColForProcess, config.OffsetKouban, "工番", config, mainWorkbook)
                    If Not config.IsOffsetHensendenjoOriginallyEmpty Then extractedData(COL_HENSHOUSHO) = GetValueFromOffset(wsKoutei, currentBaseRowForProcess, currentBaseColForProcess, config.OffsetHensendenjo, "変電所", config, mainWorkbook)
                    If Not config.IsOffsetSagyomei1OriginallyEmpty Then extractedData(COL_SAGYOMEI1) = GetValueFromOffset(wsKoutei, currentBaseRowForProcess, currentBaseColForProcess, config.OffsetSagyomei1, "作業名1", config, mainWorkbook)
                    If Not config.IsOffsetSagyomei2OriginallyEmpty Then extractedData(COL_SAGYOMEI2) = GetValueFromOffset(wsKoutei, currentBaseRowForProcess, currentBaseColForProcess, config.OffsetSagyomei2, "作業名2", config, mainWorkbook)
                    If Not config.IsOffsetTantouOriginallyEmpty Then extractedData(COL_TANTOU) = GetValueFromOffset(wsKoutei, currentBaseRowForProcess, currentBaseColForProcess, config.OffsetTantou, "担当", config, mainWorkbook)
                    If Not config.IsOffsetKoujiShuruiOriginallyEmpty Then extractedData(COL_KOUJISHURUI) = GetValueFromOffset(wsKoutei, currentBaseRowForProcess, currentBaseColForProcess, config.OffsetKoujiShurui, "工事種類", config, mainWorkbook)
                    If Not config.IsOffsetNinzuOriginallyEmpty Then extractedData(COL_NINZU) = GetValueFromOffset(wsKoutei, currentBaseRowForProcess, currentBaseColForProcess, config.OffsetNinzu, "人数", config, mainWorkbook)
                    If Not config.IsOffsetSonotaOriginallyEmpty Then extractedData(COL_SONOTA) = GetValueFromOffset(wsKoutei, currentBaseRowForProcess, currentBaseColForProcess, config.OffsetSonota, "その他", config, mainWorkbook)
                    If Not config.IsOffsetShuuryoJikanOriginallyEmpty Then extractedData(COL_SHUURYOJIKAN) = GetValueFromOffset(wsKoutei, currentBaseRowForProcess, currentBaseColForProcess, config.OffsetShuuryoJikan, "終了時間", config, mainWorkbook)
                    If Not config.IsOffsetBunrui1ExtSrcOriginallyEmpty Then extractedData(COL_BUNRUI1EXTSRC) = GetValueFromOffset(wsKoutei, currentBaseRowForProcess, currentBaseColForProcess, config.OffsetBunrui1ExtSrc, "分類1抽出元", config, mainWorkbook)
                    
                    ' Worker Names
                    actualExtractedWorkerCount = 0
                    If config.ProcessesPerDay > 0 And processIdx <= UBound(config.ProcessPatternColNumbers(1)) Then
                        maxWorkersForThisProcess = config.ProcessPatternColNumbers(1)(processIdx)
                        If Not config.IsOffsetSagyoinStartOriginallyEmpty Then ' Only extract workers if offset is defined
                            For workerIdx = 1 To Application.WorksheetFunction.Min(10, maxWorkersForThisProcess) ' Max 10 workers
                                Dim sagyoinOffsetCol As Long: sagyoinOffsetCol = config.OffsetSagyoinStart.Col + (workerIdx - 1)
                                workerName = GetValueFromOffset(wsKoutei, currentBaseRowForProcess, currentBaseColForProcess, CreateOffset(config.OffsetSagyoinStart.Row, sagyoinOffsetCol), "作業員" & workerIdx, config, mainWorkbook)
                                extractedData(COL_SAGYOIN_START + workerIdx - 1) = workerName
                                If Len(workerName) > 0 Then actualExtractedWorkerCount = actualExtractedWorkerCount + 1
                            Next workerIdx
                        End If
                    Else
                         maxWorkersForThisProcess = 0 
                    End If

                    ' Blank Row Check
                    isRowAllEmpty = True
                    If Len(CStr(extractedData(COL_KOUBAN))) > 0 Then isRowAllEmpty = False
                    If Len(CStr(extractedData(COL_HENSHOUSHO))) > 0 Then isRowAllEmpty = False
                    If Len(CStr(extractedData(COL_SAGYOMEI1))) > 0 Then isRowAllEmpty = False
                    If Len(CStr(extractedData(COL_SAGYOMEI2))) > 0 Then isRowAllEmpty = False
                    If actualExtractedWorkerCount > 0 Then isRowAllEmpty = False

                    If isRowAllEmpty Then
                        If Not filterLogSht Is Nothing Then Call M04_LogWriter.WriteFilterLogEntry(filterLogSht, GetNextFilterLogRow(filterLogSht), "空白行スキップ", kouteiFilePath & "/" & actualTargetSheetName & "/Day" & Format(dateInLoop, "dd") & "/Proc" & processIdx)
                        GoTo NextProcessInDay_M06
                    End If

                    ' Filter (Stub)
                    If Not PerformFilterCheck(extractedData, config) Then GoTo NextProcessInDay_M06

                    ' Write to wsOutput (ensure wsOutput is valid and outputNextRow is managed)
                    If Not wsOutput Is Nothing Then
                        Dim colIdx As Long
                        For colIdx = LBound(extractedData) To UBound(extractedData)
                            wsOutput.Cells(outputNextRow, colIdx).Value = extractedData(colIdx)
                        Next colIdx
                        outputNextRow = outputNextRow + 1
                        totalExtractedCount = totalExtractedCount + 1
                        anyDateExtractedSuccessfullyInFile = True 
                    Else
                        If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE: M06_DataExtractor.ExtractDataFromFile - wsOutput is Nothing. Cannot write extracted data row."
                        anyDateExtractedSuccessfullyInFile = True 
                    End If
                    
                    ' Log successful full row extraction to filter log
                    If Not filterLogSht Is Nothing Then
                        tempStr = kouteiFilePath & "/" & actualTargetSheetName & "/" & Format(dateInLoop, "yyyy-mm-dd") & "/Proc" & processIdx
                        Call M04_LogWriter.WriteFilterLogEntry(filterLogSht, GetNextFilterLogRow(filterLogSht), "行抽出成功", tempStr)
                        If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE: M06_DataExtractor.ExtractDataFromFile - Successfully extracted row for Date: " & Format(dateInLoop, "yyyy-mm-dd") & ", Process: " & processIdx
                    End If

        NextProcessInDay_M06: ' Label for skipping to next process
                Next processIdx
NextDayInLoop_M06:
            Next dayIdx
NextSheetInLoop_M06:
        Next eachSheetName
        
        If Not targetSheetProcessed And UBound(config.TargetSheetNames) >= LBound(config.TargetSheetNames) Then ' Loop ran but all sheet names were empty
             Call M04_LogWriter.SafeWriteErrorLog("WARNING", mainWorkbook, config.ErrorLogSheetName, "M06_DataExtractor", "ExtractDataFromFile", "処理対象シート名リスト(config.TargetSheetNames)に有効なシート名がありませんでした。", 0, kouteiFilePath)
             GoTo ExtractDataFromFile_Finally
        End If
    Else
        Call M04_LogWriter.SafeWriteErrorLog("ERROR", mainWorkbook, config.ErrorLogSheetName, "M06_DataExtractor", "ExtractDataFromFile", "処理対象シート名リスト(config.TargetSheetNames)が空または未初期化です。", 0, kouteiFilePath)
        GoTo ExtractDataFromFile_Finally
    End If

    If Not yearMonthEstablishedForThisFile And targetSheetProcessed Then ' Check if Y/M was never established despite processing sheets
        tempStr = "ファイル「" & kouteiFilePath & "」内のどの指定シートからも有効な年/月を取得できませんでした。"
        Call M04_LogWriter.SafeWriteErrorLog("ERROR", mainWorkbook, config.ErrorLogSheetName, "M06_DataExtractor", "ExtractDataFromFile (YearMonthEstablishment)", tempStr, 0, "ファイル処理エラー")
        anyDateExtractedSuccessfullyInFile = False ' Ensure this is False
    End If
    ExtractDataFromFile = anyDateExtractedSuccessfullyInFile ' Assign final determined value

ExtractDataFromFile_Finally:
    If Not wbKoutei Is Nothing Then wbKoutei.Close SaveChanges:=False
    Set wsKoutei = Nothing
    Set wbKoutei = Nothing
    Set filterLogSht = Nothing
    If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M06_DataExtractor.ExtractDataFromFile - Closed file: '" & kouteiFilePath & "', Result: " & ExtractDataFromFile
    Exit Function

ExtractDataFromFile_Error:
    tempStr = "実行時エラー " & Err.Number & ": " & Err.Description & ", Procedure: ExtractDataFromFile, File: " & kouteiFilePath
    Call M04_LogWriter.SafeWriteErrorLog("ERROR", mainWorkbook, config.ErrorLogSheetName, "M06_DataExtractor", "ExtractDataFromFile", tempStr, Err.Number, Err.Description)
    ExtractDataFromFile = False ' Ensure False on error
    Resume ExtractDataFromFile_Finally
End Function
