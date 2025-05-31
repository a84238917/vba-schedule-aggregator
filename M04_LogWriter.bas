' バージョン：v0.5.1
Option Explicit
' このモジュールは、ログシートへの書き込み処理を担当します。
' エラー情報、処理の進捗、フィルタ条件などを指定されたシートに記録します。

Private Const MODULE_NAME As String = "M04_LogWriter"

' Public Sub: WriteErrorLog
' エラーログシートにエラー情報を書き込みます。
Public Sub WriteErrorLog(ByVal errorLevel As String, ByVal moduleN As String, ByVal procedureN As String, _
                         ByVal message As String, Optional errNumber As Long = 0, Optional errDescription As String = "")
    Dim funcName As String: funcName = "WriteErrorLog"

    On Error GoTo ErrorHandler

    If g_errorLogWorksheet Is Nothing Then
        Debug.Print Now & " WriteErrorLog FATAL: g_errorLogWorksheet is Not Set. Cannot log error."
        Debug.Print "  > Level: " & errorLevel & ", Module: " & moduleN & ", Proc: " & procedureN
        Debug.Print "  > Message: " & message
        If errNumber <> 0 Then Debug.Print "  > Err #: " & errNumber & " - " & errDescription
        Exit Sub
    End If

    If g_nextErrorLogRow <= 0 Then
        ' M03_SheetManager.PrepareSheets で設定されるはずだが、万が一のためのフォールバック
        If Application.WorksheetFunction.CountA(g_errorLogWorksheet.Rows(1)) = 0 Then
             g_nextErrorLogRow = 1
        Else
             g_nextErrorLogRow = g_errorLogWorksheet.Cells(Rows.Count, "A").End(xlUp).Row + 1
        End If
        If g_nextErrorLogRow <=0 Then g_nextErrorLogRow = 1 '最終フォールバック
    End If


    With g_errorLogWorksheet
        .Cells(g_nextErrorLogRow, 1).Value = Now() ' 日時
        .Cells(g_nextErrorLogRow, 2).Value = errorLevel
        .Cells(g_nextErrorLogRow, 3).Value = moduleN
        .Cells(g_nextErrorLogRow, 4).Value = procedureN
        .Cells(g_nextErrorLogRow, 5).Value = message
        If errNumber <> 0 Then
            .Cells(g_nextErrorLogRow, 6).Value = errNumber
            .Cells(g_nextErrorLogRow, 7).Value = errDescription
        Else
            .Cells(g_nextErrorLogRow, 6).Value = vbNullString ' 空白を明示
            .Cells(g_nextErrorLogRow, 7).Value = vbNullString ' 空白を明示
        End If
    End With

    g_nextErrorLogRow = g_nextErrorLogRow + 1
    Exit Sub

ErrorHandler:
    Debug.Print Now & " CRITICAL ERROR in M04_LogWriter.WriteErrorLog itself! Err# " & Err.Number & " - " & Err.Description
    Debug.Print "  Original Log Attempt: Level=" & errorLevel & ", Module=" & moduleN & ", Proc=" & procedureN & ", Msg=" & message
End Sub

' Public Sub: WriteFilterLog
' 検索条件ログシートに、マクロ実行時の主要なフィルタ設定などを記録します。
Public Sub WriteFilterLog(ByRef config As tConfigSettings, ByVal wb As Workbook)
    Dim funcName As String: funcName = "WriteFilterLog"
    Dim wsLog As Worksheet
    Dim nextLogWriteRow As Long
    Dim i As Long ' ループカウンタ

    On Error GoTo ErrorHandler

    If Trim(config.SearchConditionLogSheetName) = "" Then
        Call WriteErrorLog("WARNING", MODULE_NAME, funcName, "検索条件ログシート名が設定されていません。ログ記録をスキップします。")
        Exit Sub
    End If

    On Error Resume Next
    Set wsLog = wb.Sheets(config.SearchConditionLogSheetName)
    On Error GoTo ErrorHandler

    If wsLog Is Nothing Then
        Call WriteErrorLog("ERROR", MODULE_NAME, funcName, "検索条件ログシート「" & config.SearchConditionLogSheetName & "」が見つかりません。")
        Exit Sub
    End If

    If Application.WorksheetFunction.CountA(wsLog.Rows(1)) = 0 Then
         nextLogWriteRow = 1
    Else
         nextLogWriteRow = wsLog.Cells(Rows.Count, "A").End(xlUp).Row + 1
    End If
    If nextLogWriteRow <= 0 Then nextLogWriteRow = 1

    ' マクロ基本情報
    Call WriteFilterLogEntry(wsLog, nextLogWriteRow, "マクロ実行開始時刻", CStr(config.StartTime))
    Call WriteFilterLogEntry(wsLog, nextLogWriteRow, "マクロファイル", config.ScriptFullName)
    Call WriteFilterLogEntry(wsLog, nextLogWriteRow, "設定ファイルシート", config.ConfigSheetFullName)
    Call WriteFilterLogEntry(wsLog, nextLogWriteRow, "デバッグモード", CStr(config.DebugModeFlag))

    ' A. 一般設定
    Call WriteFilterLogEntry(wsLog, nextLogWriteRow, "A.デフォルトフォルダパス", config.DefaultFolderPath)
    Call WriteFilterLogEntry(wsLog, nextLogWriteRow, "A.抽出結果出力シート名", config.OutputSheetName)
    Call WriteFilterLogEntry(wsLog, nextLogWriteRow, "A.工程パターンデータ取得方法", IIf(config.GetPatternDataMethod, "数式", "VBA"))

    ' B. 工程表ファイル設定
    If General_IsArrayInitialized(config.TargetSheetNames) Then
        Call WriteFilterLogArrayEntry(wsLog, nextLogWriteRow, "B.検索対象シート名リスト", config.TargetSheetNames)
    End If
    Call WriteFilterLogEntry(wsLog, nextLogWriteRow, "B.工程表ヘッダー行数", CStr(config.HeaderRowCount))
    Call WriteFilterLogEntry(wsLog, nextLogWriteRow, "B.1日の工程数", CStr(config.ProcessesPerDay))


    ' D. フィルタ条件
    Call WriteFilterLogEntry(wsLog, nextLogWriteRow, "D.作業員フィルター検索論理", config.WorkerFilterLogic)
    If General_IsArrayInitialized(config.WorkerFilterList) Then
        Call WriteFilterLogArrayEntry(wsLog, nextLogWriteRow, "D.作業員フィルターリスト", config.WorkerFilterList)
    End If
    Call WriteFilterLogEntry(wsLog, nextLogWriteRow, "D.人数フィルター", config.NinzuFilter & IIf(config.IsNinzuFilterOriginallyEmpty, " (元々空)", ""))

    ' E. 処理対象ファイル定義
    If General_IsArrayInitialized(config.TargetFileFolderPaths) Then
         Call WriteFilterLogArrayEntry(wsLog, nextLogWriteRow, "E.処理対象ファイル/フォルダパスリスト", config.TargetFileFolderPaths)
    End If
    If General_IsArrayInitialized(config.FilePatternIdentifiers) Then
         Call WriteFilterLogArrayEntry(wsLog, nextLogWriteRow, "E.適用工程パターン識別子リスト", config.FilePatternIdentifiers)
    End If

    ' G. 出力シート設定
    Call WriteFilterLogEntry(wsLog, nextLogWriteRow, "G.出力データオプション", config.OutputDataOption)
    If General_IsArrayInitialized(config.OutputHeaderContents) Then
         Call WriteFilterLogArrayEntry(wsLog, nextLogWriteRow, "G.出力シートヘッダー内容", config.OutputHeaderContents)
    End If

    ' ログ書き込み完了メッセージ (エラーログへ)
    Call WriteErrorLog("INFORMATION", MODULE_NAME, funcName, "検索条件ログの書き込みが完了しました。")
    Exit Sub

ErrorHandler:
    Call WriteErrorLog("ERROR", MODULE_NAME, funcName, "検索条件ログ書き込み中にエラー。", Err.Number, Err.Description)
End Sub

' Private Sub: WriteFilterLogEntry
Private Sub WriteFilterLogEntry(ByVal ws As Worksheet, ByRef nextRow As Long, ByVal item As String, ByVal value As String)
    On Error Resume Next
    ws.Cells(nextRow, 1).Value = Now()
    ws.Cells(nextRow, 2).Value = item
    ws.Cells(nextRow, 3).Value = value
    If Err.Number = 0 Then
        nextRow = nextRow + 1
    Else
        Debug.Print Now & " Error writing filter log entry: " & item & " - " & value & " (Err: " & Err.Number & " - " & Err.Description & ")"
        Err.Clear
    End If
    On Error GoTo 0
End Sub

' Private Sub: WriteFilterLogArrayEntry
Private Sub WriteFilterLogArrayEntry(ByVal ws As Worksheet, ByRef nextRow As Long, ByVal itemBaseName As String, ByRef arr() As String)
    Dim i As Long
    Dim currentItemName As String

    If Not General_IsArrayInitialized(arr) Then Exit Sub

    For i = LBound(arr) To UBound(arr)
        currentItemName = itemBaseName ' 配列の場合、要素ごとにインデックスを付けないシンプルなログ形式
        If Trim(arr(i)) <> "" Then
            ' 配列の各要素を個別の行として記録。項目名は同じitemBaseNameを使う。
             Call WriteFilterLogEntry(ws, nextRow, itemBaseName, arr(i))
        End If
    Next i
End Sub

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
