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

Public Function ExtractDataFromFile(kouteiFilePath As String, ByRef config As tConfigSettings, Optional wsOutput As Worksheet = Nothing, Optional ByRef outputNextRow As Long = 0, Optional ByVal currentFileNum As Long = 0, Optional ByRef totalExtractedCount As Long = 0) As Boolean
    ' 個別工程表ファイルからデータを抽出し、フィルターログに年月日の基本情報を記録します。
    ' Arguments:
    '   kouteiFilePath: (I) String型。処理対象の工程表ファイルパス。
    '   config: (I) tConfigSettings型。各種設定情報。
    '   wsOutput: (I/O) Worksheet型 (Optional)。抽出データの出力先シート。
    '   outputNextRow: (I/O) Long型 (Optional)。出力シートの次の書き込み行。
    '   currentFileNum: (I) Long型 (Optional)。現在処理中のファイル番号 (ログ用)。
    '   totalExtractedCount: (I/O) Long型 (Optional)。抽出成功した総件数。
    ' Returns:
    '   Boolean: データ抽出処理が（部分的にでも）成功した場合はTrue、それ以外はFalse。

    Dim wbKoutei As Workbook, wsKoutei As Worksheet
    Dim currentYear As Long, currentMonth As Long, dayIdx As Long
    Dim dayVal As Variant, dateInLoop As Date, targetSheetName As String
    Dim dayCellRow As Long, tempStr As String
    Dim filterLogSht As Worksheet

    ExtractDataFromFile = False ' Default to failure
    On Error GoTo ExtractDataFromFile_Error

    If DEBUG_MODE_DETAIL Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M06_DataExtractor.ExtractDataFromFile - Opening file: '" & kouteiFilePath & "'"
    Set wbKoutei = Workbooks.Open(Filename:=kouteiFilePath, UpdateLinks:=0, ReadOnly:=True)
    If wbKoutei Is Nothing Then Exit Function ' Should be caught by error handler, but as a safeguard

    ' --- 対象シート特定 (リストの最初のシートのみ) ---
    If LogExtractor_IsArrayInitialized(config.TargetSheetNames) Then
        If UBound(config.TargetSheetNames) >= LBound(config.TargetSheetNames) Then
            targetSheetName = config.TargetSheetNames(LBound(config.TargetSheetNames))
        Else
            Call M04_LogWriter.SafeWriteErrorLog(ActiveWorkbook, config.ErrorLogSheetName, "M06_DataExtractor", "ExtractDataFromFile", "処理対象シート名リスト(config.TargetSheetNames)が空です。", 0, kouteiFilePath)
            GoTo ExtractDataFromFile_Finally ' Graceful exit
        End If
    Else
        Call M04_LogWriter.SafeWriteErrorLog(ActiveWorkbook, config.ErrorLogSheetName, "M06_DataExtractor", "ExtractDataFromFile", "処理対象シート名リスト(config.TargetSheetNames)が初期化されていません。", 0, kouteiFilePath)
        GoTo ExtractDataFromFile_Finally
    End If

    If DEBUG_MODE_DETAIL Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M06_DataExtractor.ExtractDataFromFile - Attempting to access sheet: '" & targetSheetName & "' in file '" & kouteiFilePath & "'"
    On Error Resume Next ' Temporarily disable error handling for sheet existence check
    Set wsKoutei = wbKoutei.Worksheets(targetSheetName)
    On Error GoTo ExtractDataFromFile_Error ' Restore main error handler
    
    If wsKoutei Is Nothing Then
        Call M04_LogWriter.SafeWriteErrorLog(ActiveWorkbook, config.ErrorLogSheetName, "M06_DataExtractor", "ExtractDataFromFile", "シートが見つかりません: " & targetSheetName, 0, kouteiFilePath)
        GoTo ExtractDataFromFile_Finally
    End If

    ' --- 年月取得 ---
    Dim yearVal As Variant, monthVal As Variant
    yearVal = wsKoutei.Range(config.YearCellAddress).Value
    monthVal = wsKoutei.Range(config.MonthCellAddress).Value
    
    If Not IsNumeric(yearVal) Or Not IsNumeric(monthVal) Or CLng(yearVal) < 1900 Or CLng(yearVal) > 2999 Or CLng(monthVal) < 1 Or CLng(monthVal) > 12 Then
        tempStr = "年セル(" & config.YearCellAddress & ")または月セル(" & config.MonthCellAddress & ")の値が不正です。Year='" & CStr(yearVal) & "', Month='" & CStr(monthVal) & "'"
        Call M04_LogWriter.SafeWriteErrorLog(ActiveWorkbook, config.ErrorLogSheetName, "M06_DataExtractor", "ExtractDataFromFile", tempStr, 0, kouteiFilePath & "/" & targetSheetName)
        GoTo ExtractDataFromFile_Finally
    End If
    currentYear = CLng(yearVal)
    currentMonth = CLng(monthVal)
    If DEBUG_MODE_DETAIL Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M06_DataExtractor.ExtractDataFromFile - Year: " & currentYear & ", Month: " & currentMonth

    ' --- フィルターログシート取得 ---
    On Error Resume Next ' Temporarily disable error handling for sheet existence check
    Set filterLogSht = ActiveWorkbook.Worksheets(config.SearchConditionLogSheetName)
    On Error GoTo ExtractDataFromFile_Error ' Restore main error handler
    
    If filterLogSht Is Nothing Then
        If DEBUG_MODE_ERROR Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - ERROR: M06_DataExtractor.ExtractDataFromFile - Filter log sheet '" & config.SearchConditionLogSheetName & "' not found. Cannot log date extractions."
        ' Proceeding without filter logging for this file if sheet is missing, but error log should have info from PrepareSheets
    End If
        
    ' --- 日処理ループ ---
    For dayIdx = 1 To config.MaxDaysPerSheet
        dayCellRow = config.HeaderRowCount + (dayIdx - 1) * config.RowsPerDay + config.DayRowOffset
        dayVal = wsKoutei.Cells(dayCellRow, config.DayColumnLetter).Value

        If Not IsNumeric(dayVal) Or Len(Trim(CStr(dayVal))) = 0 Or CLng(dayVal) <= 0 Then
            If DEBUG_MODE_WARNING And Not IsEmpty(dayVal) And Len(Trim(CStr(dayVal))) > 0 Then ' Log only if cell wasn't truly blank, but invalid
                tempStr = "日付セルの値が不正または0以下です (" & config.DayColumnLetter & dayCellRow & "): '" & CStr(dayVal) & "'"
                If Not filterLogSht Is Nothing Then Call M04_LogWriter.WriteFilterLogEntry(filterLogSht, GetNextFilterLogRow(filterLogSht), "日付取得エラー(スキップ)", kouteiFilePath & "/" & targetSheetName & "/" & tempStr)
            End If
            GoTo NextDayInLoop_M06 ' Skip this day
        End If

        On Error Resume Next ' For DateSerial error
        dateInLoop = DateSerial(currentYear, currentMonth, CLng(dayVal))
        If Err.Number <> 0 Then
            On Error GoTo ExtractDataFromFile_Error ' Restore error handler
            tempStr = "DateSerialで無効な日付 (" & currentYear & "/" & currentMonth & "/" & CLng(dayVal) & "). Error: " & Err.Description
            If Not filterLogSht Is Nothing Then Call M04_LogWriter.WriteFilterLogEntry(filterLogSht, GetNextFilterLogRow(filterLogSht), "日付検証エラー(スキップ)", kouteiFilePath & "/" & targetSheetName & "/" & tempStr)
            Err.Clear
            GoTo NextDayInLoop_M06 ' Skip this day
        End If
        On Error GoTo ExtractDataFromFile_Error ' Restore error handler

        If Not filterLogSht Is Nothing Then
            tempStr = kouteiFilePath & "/" & targetSheetName & "/" & Format(dateInLoop, "yyyy-mm-dd")
            Call M04_LogWriter.WriteFilterLogEntry(filterLogSht, GetNextFilterLogRow(filterLogSht), "日付抽出成功", tempStr)
            If DEBUG_MODE_DETAIL Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M06_DataExtractor.ExtractDataFromFile - Extracted Date: " & Format(dateInLoop, "yyyy-mm-dd")
        End If
        ' このステップでは工程処理ループと詳細データ項目抽出は行わない
NextDayInLoop_M06:
    Next dayIdx
    ExtractDataFromFile = True ' Mark as success if loop completes (or even partially runs for some dates)

ExtractDataFromFile_Finally:
    If Not wbKoutei Is Nothing Then wbKoutei.Close SaveChanges:=False
    Set wsKoutei = Nothing
    Set wbKoutei = Nothing
    Set filterLogSht = Nothing
    If DEBUG_MODE_DETAIL Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M06_DataExtractor.ExtractDataFromFile - Closed file: '" & kouteiFilePath & "', Result: " & ExtractDataFromFile
    Exit Function

ExtractDataFromFile_Error:
    tempStr = "実行時エラー " & Err.Number & ": " & Err.Description & ", Procedure: ExtractDataFromFile, File: " & kouteiFilePath
    Call M04_LogWriter.SafeWriteErrorLog(ActiveWorkbook, config.ErrorLogSheetName, "M06_DataExtractor", "ExtractDataFromFile", tempStr, Err.Number, Err.Description)
    ExtractDataFromFile = False ' Ensure False on error
    Resume ExtractDataFromFile_Finally
End Function
