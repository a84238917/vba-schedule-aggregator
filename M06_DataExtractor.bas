' v0.4.0
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

Private Function GetValueFromOffset(targetCell As Range, rowOffset As Long, colOffset As Long, Optional itemDesc As String = "") As Variant
    ' 指定されたセルからのオフセット位置にあるセルの値を取得します。(スタブ)
    ' 将来的にはエラーハンドリングや型変換などを追加する可能性があります。
    GetValueFromOffset = "" ' Placeholder
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
                
                ' --- Successful Date Extraction ---
                anyDateExtractedSuccessfullyInFile = True ' Add this line
                If Not filterLogSht Is Nothing Then
                    tempStr = kouteiFilePath & "/" & actualTargetSheetName & "/" & Format(dateInLoop, "yyyy-mm-dd")
                    Call M04_LogWriter.WriteFilterLogEntry(filterLogSht, GetNextFilterLogRow(filterLogSht), "日付抽出成功", tempStr)
                    If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE: M06_DataExtractor.ExtractDataFromFile - Extracted Date: " & Format(dateInLoop, "yyyy-mm-dd") & " from cell " & config.DayColumnLetter & dayCellRow
                End If
                ' このステップでは工程処理ループと詳細データ項目抽出は行わない
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

    ExtractDataFromFile = anyDateExtractedSuccessfullyInFile

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
