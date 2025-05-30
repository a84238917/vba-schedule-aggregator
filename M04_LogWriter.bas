' バージョン：v0.5.0
Option Explicit
' このモジュールは、エラーログおよび検索条件ログのシートへの書き込み処理を専門に担当します。

Public Sub WriteErrorLog(errorLevel As String, moduleName As String, procedureName As String, relatedInfo As String, errorNumber As Long, errorDescription As String, Optional actionTaken As String = "", Optional variableInfo As String = "")
    ' エラー情報をグローバルエラーログシート(g_errorLogWorksheet)に書き込みます。
    ' Arguments:
    '   errorLevel: エラーの重要度 ("ERROR", "WARNING", "INFO"など)
    '   moduleName: エラーが発生したモジュール名
    '   procedureName: エラーが発生したプロシージャ名
    '   relatedInfo: 関連情報（ファイル名、シート名など）
    '   errorNumber: エラー番号
    '   errorDescription: エラー内容
    '   actionTaken: (Optional) エラーに対する対処内容
    '   variableInfo: (Optional) エラー発生時の関連変数情報 (32767文字に切り詰められます)

    On Error GoTo WriteErrorLog_InternalError

    If g_errorLogWorksheet Is Nothing Then
        If DEBUG_MODE_ERROR Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - ERROR: M04_LogWriter.WriteErrorLog - g_errorLogWorksheet is Nothing. Cannot write error log."
        Exit Sub
    End If
    
    If g_configSettings.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M04_LogWriter.WriteErrorLog - Writing to: '" & g_errorLogWorksheet.Name & "'!A" & g_nextErrorLogRow & ", Level: " & errorLevel & ", Module: " & moduleName

    With g_errorLogWorksheet
        .Cells(g_nextErrorLogRow, 1).Value = errorLevel                         ' A: 重要度
        .Cells(g_nextErrorLogRow, 2).Value = Format(Now, "yyyy/mm/dd hh:nn:ss") ' B: 発生日時
        .Cells(g_nextErrorLogRow, 3).Value = moduleName                         ' C: モジュール
        .Cells(g_nextErrorLogRow, 4).Value = procedureName                       ' D: プロシージャ
        .Cells(g_nextErrorLogRow, 5).Value = relatedInfo                         ' E: 関連情報
        .Cells(g_nextErrorLogRow, 6).Value = errorNumber                         ' F: エラー番号
        .Cells(g_nextErrorLogRow, 7).Value = "'" & errorDescription              ' G: エラー内容 (先頭にアポストロフィ)
        .Cells(g_nextErrorLogRow, 8).Value = actionTaken                         ' H: 対処内容
        .Cells(g_nextErrorLogRow, 9).Value = Left(variableInfo, 32767)           ' I: 変数情報 (最大長制限)
    End With
    g_nextErrorLogRow = g_nextErrorLogRow + 1
    Exit Sub

WriteErrorLog_InternalError:
    If DEBUG_MODE_ERROR Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - CRITICAL_ERROR: M04_LogWriter.WriteErrorLog internal error - " & Err.Description
End Sub

Public Sub SafeWriteErrorLog(errorLevel As String, targetWorkbook As Workbook, errorLogSheetNameAttempt As String, moduleName As String, procedureName As String, relatedInfo As String, errorNumber As Long, errorDescription As String, Optional actionTaken As String = "", Optional variableInfo As String = "")
    ' Configシート読み込み前やグローバル変数が未初期化の段階でも使用可能な、より堅牢なエラーログ書き込み処理です。
    ' 指定されたワークブックとシート名でエラーログシートを特定または作成し、情報を書き込みます。
    ' Arguments:
    '   errorLevel: エラーの重要度 ("ERROR", "WARNING", "INFO"など)
    '   targetWorkbook: 書き込み対象のワークブック
    '   errorLogSheetNameAttempt: 試行するエラーログシート名
    '   moduleName: エラーが発生したモジュール名
    '   procedureName: エラーが発生したプロシージャ名
    '   relatedInfo: 関連情報
    '   errorNumber: エラー番号
    '   errorDescription: エラー内容
    '   actionTaken: (Optional) 対処内容
    '   variableInfo: (Optional) 変数情報

    On Error Resume Next ' このプロシージャ全体のエラーはできるだけ無視して処理を試みる

    If Trim(errorLogSheetNameAttempt) = "" Then
        If DEBUG_MODE_ERROR Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - ERROR: M04_LogWriter.SafeWriteErrorLog - errorLogSheetNameAttempt is empty. Cannot write log."
        Exit Sub
    End If

    Dim ws As Worksheet
    Dim nextRow As Long

    If targetWorkbook Is Nothing Then Exit Sub ' ワークブックが無効なら終了

    Set ws = Nothing ' 初期化
    Set ws = targetWorkbook.Sheets(errorLogSheetNameAttempt)

    If ws Is Nothing Then ' シートが存在しない場合
        Set ws = targetWorkbook.Sheets.Add(After:=targetWorkbook.Sheets(targetWorkbook.Sheets.Count))
        If ws Is Nothing Then Exit Sub ' シート追加に失敗した場合 (例: 保護されたブックなど)
        ws.Name = errorLogSheetNameAttempt
        ' ヘッダーを書き込む (New 9-column structure)
        ws.Cells(1, 1).Value = "重要度"
        ws.Cells(1, 2).Value = "発生日時"
        ws.Cells(1, 3).Value = "モジュール"
        ws.Cells(1, 4).Value = "プロシージャ"
        ws.Cells(1, 5).Value = "関連情報"
        ws.Cells(1, 6).Value = "エラー番号"
        ws.Cells(1, 7).Value = "エラー内容"
        ws.Cells(1, 8).Value = "対処内容"
        ws.Cells(1, 9).Value = "変数情報"
        nextRow = 2
    Else ' シートが存在する場合 (Revised nextRow logic)
        If ws.Cells(1, 1).Value = vbNullString Then ' Check if A1 (which should be "重要度") is empty
            ' Sheet is not new, but A1 is empty. Check if other cells in Col A have data.
            If ws.Cells(ws.Rows.Count, 1).End(xlUp).Row = 1 And ws.Cells(1,1).Value = vbNullString Then ' Added second check for A1 again for clarity
                nextRow = 1
            Else
                nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
            End If
        Else
            ' A1 has data
            nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        End If
        If nextRow > ws.Rows.Count Then nextRow = ws.Rows.Count ' Safety for very full sheet
        If nextRow <= 0 Then nextRow = 1 ' Ensure nextRow is at least 1
        
        ' If headers were expected and nextRow is 1, ensure headers are written.
        ' This part is maintained from previous logic, now applied after the new nextRow calculation.
        ' If nextRow becomes 1 for an existing sheet, it means the sheet was likely empty or headerless.
        ' Write headers if they seem missing (checking new A1 "重要度" and old A1 "発生日時" for robustness).
        If nextRow = 1 And (ws.Cells(1,1).Value = vbNullString Or ws.Cells(1,2).Value = vbNullString) Then
            ws.Cells(1, 1).Value = "重要度"
            ws.Cells(1, 2).Value = "発生日時"
            ws.Cells(1, 3).Value = "モジュール"
            ws.Cells(1, 4).Value = "プロシージャ"
            ws.Cells(1, 5).Value = "関連情報"
            ws.Cells(1, 6).Value = "エラー番号"
            ws.Cells(1, 7).Value = "エラー内容"
            ws.Cells(1, 8).Value = "対処内容"
            ws.Cells(1, 9).Value = "変数情報"
        End If
    End If

    If g_configSettings.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M04_LogWriter.SafeWriteErrorLog - Writing to: '" & ws.Name & "'!A" & nextRow & ", Level: " & errorLevel & ", Module: " & moduleName

    With ws
        .Cells(nextRow, 1).Value = errorLevel                         ' A: 重要度
        .Cells(nextRow, 2).Value = Format(Now, "yyyy/mm/dd hh:nn:ss") ' B: 発生日時
        .Cells(nextRow, 3).Value = moduleName                         ' C: モジュール
        .Cells(nextRow, 4).Value = procedureName                       ' D: プロシージャ
        .Cells(nextRow, 5).Value = relatedInfo                         ' E: 関連情報
        .Cells(nextRow, 6).Value = errorNumber                         ' F: エラー番号
        .Cells(nextRow, 7).Value = "'" & errorDescription              ' G: エラー内容
        .Cells(nextRow, 8).Value = actionTaken                         ' H: 対処内容
        .Cells(nextRow, 9).Value = Left(variableInfo, 32767)           ' I: 変数情報
    End With

    Set ws = Nothing
    Err.Clear ' このプロシージャ内で発生した可能性のあるエラーをクリア
End Sub

Public Sub WriteFilterLog(ByRef config As tConfigSettings, ByVal targetWorkbook As Workbook)
    ' 検索条件ログシートに、設定されたフィルター条件やマクロの基本情報を書き込みます。
    ' これは現在フレームワークであり、今後拡張されてより多くの情報が記録される予定です。
    ' Arguments:
    '   config: 読み込まれた設定情報 (tConfigSettings型)
    '   targetWorkbook: ログを書き込む対象のワークブック

    Dim wsLog As Worksheet
    Dim nextLogWriteRow As Long

    On Error GoTo WriteFilterLog_Error

    If targetWorkbook Is Nothing Or Len(config.SearchConditionLogSheetName) = 0 Then
        If DEBUG_MODE_WARNING Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - WARNING: M04_LogWriter.WriteFilterLog - Target workbook or SearchConditionLogSheetName is invalid."
        Exit Sub
    End If

    Set wsLog = Nothing ' 初期化
    On Error Resume Next ' シート存在確認のエラーをハンドル
    Set wsLog = targetWorkbook.Sheets(config.SearchConditionLogSheetName)
    On Error GoTo WriteFilterLog_Error ' エラーハンドラを元に戻す
    
    If wsLog Is Nothing Then
        If DEBUG_MODE_WARNING Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - WARNING: M04_LogWriter.WriteFilterLog - SearchConditionLogSheetName '" & config.SearchConditionLogSheetName & "' not found. Should be created by M03_SheetManager."
        Exit Sub ' Should be created by M03_SheetManager
    End If

    nextLogWriteRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1
    If wsLog.Cells(1, 1).Value = vbNullString And nextLogWriteRow = 2 Then ' Handle completely empty sheet
        nextLogWriteRow = 1
    End If
    If nextLogWriteRow < 1 Then nextLogWriteRow = 1 ' 念のため

    ' ヘッダー行がなければ書き込む (M03で作成されるはずだが念のため)
    ' A列は日時なのでB列(項目名)とC列(値)のヘッダを確認
    If nextLogWriteRow = 1 And (wsLog.Cells(1, 2).Value = vbNullString Or wsLog.Cells(1,3).Value = vbNullString) Then
        wsLog.Cells(1, 1).Value = "記録日時"
        wsLog.Cells(1, 2).Value = "項目名"
        wsLog.Cells(1, 3).Value = "値"
        nextLogWriteRow = 2 ' ヘッダー書いたのでデータは次から
    End If
    
    If g_configSettings.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M04_LogWriter.WriteFilterLog - Attempting to write initial logs to: '" & wsLog.Name & "'!A" & nextLogWriteRow

    Call WriteFilterLogEntry(wsLog, nextLogWriteRow, "マクロ実行", "開始: " & Format(config.StartTime, "yyyy/mm/dd hh:nn:ss"))
    Call WriteFilterLogEntry(wsLog, nextLogWriteRow, "実行ファイルパス", config.ScriptFullName)
    Call WriteFilterLogEntry(wsLog, nextLogWriteRow, "---", "---") ' 区切り線

    ' TODO: ここにconfig内の各種フィルター情報を書き出す処理を追加する
    ' 例: Call WriteFilterLogEntry(wsLog, nextLogWriteRow, "作業員フィルター論理", config.WorkerFilterLogic)
    ' 例: Call WriteFilterLogArrayEntry(wsLog, nextLogWriteRow, "作業員フィルターリスト", config.WorkerFilterList)
    ' ... 他のフィルター条件 ...

    If g_configSettings.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M04_LogWriter.WriteFilterLog - Filter log entries written up to row " & nextLogWriteRow -1
    Exit Sub

WriteFilterLog_Error:
    If DEBUG_MODE_ERROR Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - ERROR: M04_LogWriter.WriteFilterLog - Error " & Err.Number & ": " & Err.Description
    Call SafeWriteErrorLog("ERROR", targetWorkbook, config.ErrorLogSheetName, "M04_LogWriter", "WriteFilterLog", "フィルターログ書き込みエラー", Err.Number, Err.Description)
End Sub

Public Sub WriteFilterLogEntry(targetLogSheet As Worksheet, ByRef nextLogRow As Long, itemName As String, itemValue As String)
    ' 検索条件ログシートに単一の項目名と値を書き込み、次の書き込み行を更新します。
    ' Arguments:
    '   targetLogSheet: 書き込み対象のログシート
    '   nextLogRow: (ByRef) 書き込む行番号。このプロシージャ内でインクリメントされます。
    '   itemName: 書き込む項目名
    '   itemValue: 書き込む値

    If targetLogSheet Is Nothing Then Exit Sub
    If nextLogRow < 1 Then nextLogRow = 1 ' 行番号が不正な場合は1行目から

    targetLogSheet.Cells(nextLogRow, 1).Value = Format(Now, "yyyy/mm/dd hh:nn:ss") ' A: 記録日時
    targetLogSheet.Cells(nextLogRow, 2).Value = itemName                         ' B: 項目名
    targetLogSheet.Cells(nextLogRow, 3).Value = itemValue                        ' C: 値
    nextLogRow = nextLogRow + 1
End Sub

Public Sub WriteFilterLogArrayEntry(targetLogSheet As Worksheet, ByRef nextLogRow As Long, itemName As String, ByRef itemArray() As String)
    ' 検索条件ログシートに、文字列配列の内容を単一のエントリとして書き込みます。
    ' 配列が空または未初期化の場合、その状態を示す文字列が書き込まれます。
    ' Arguments:
    '   targetLogSheet: 書き込み対象のログシート
    '   nextLogRow: (ByRef) 書き込む行番号。このプロシージャ内でインクリメントされます。
    '   itemName: 書き込む項目名
    '   itemArray: (ByRef) 書き込む文字列配列

    Dim outputValue As String

    If LogWriter_IsArrayInitialized(itemArray) Then
        If UBound(itemArray) - LBound(itemArray) + 1 > 0 Then ' 配列に要素が1つ以上存在するか
            outputValue = Join(itemArray, ", ")
        Else
            outputValue = "(リスト空)" ' 例: ReDim MyArray(0 To -1) のような状態
        End If
    Else
        outputValue = "(リスト未設定)" ' 例: Dim MyArray() のみでReDimされていない状態
    End If
    Call WriteFilterLogEntry(targetLogSheet, nextLogRow, itemName, outputValue)
End Sub

Private Function LogWriter_IsArrayInitialized(arr As Variant) As Boolean
    ' 配列が有効に初期化されているか（少なくとも1つの要素を持つか）を確認します。
    ' Variant型が配列でない場合、または配列であっても要素が割り当てられていない場合（Dim arr() のみでReDimされていない状態など）はFalseを返します。
    ' Arguments:
    '   arr: 確認対象のVariant変数
    On Error GoTo NotAnArrayOrNotInitialized
    If IsArray(arr) Then
        Dim lBoundCheck As Long
        lBoundCheck = LBound(arr) ' 配列がReDimされていれば、LBoundはエラーにならない (空でも ReDim arr(0 To -1) など)
        LogWriter_IsArrayInitialized = True ' LBoundがエラーを起こさなければ、配列は有効（空でもReDimされていればOK）
        Exit Function
    End If
NotAnArrayOrNotInitialized:
    LogWriter_IsArrayInitialized = False
End Function
