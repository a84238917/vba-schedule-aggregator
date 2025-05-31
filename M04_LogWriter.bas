' バージョン：v0.5.1
Option Explicit
' このモジュールは、ログシートへの書き込み処理を担当します。
' エラー情報、処理の進捗、フィルタ条件などを指定されたシートに記録します。

Private Const MODULE_NAME As String = "M04_LogWriter"

' Public Sub: WriteErrorLog
' エラーログシートにエラー情報を書き込みます。
Public Sub WriteErrorLog(ByVal errorLevel As String, ByVal moduleN As String, ByVal procedureN As String, _
                         ByVal message As String, Optional errNumber As Long = 0, Optional errDescription As String = "")
    Dim funcName_Internal As String: funcName_Internal = "WriteErrorLog_Internal"
    Dim targetSheet As Worksheet
    Dim targetNextRow As Long
    Dim isDebugLog As Boolean
    isDebugLog = (Left(errorLevel, 6) = "DEBUG_")

    ' ▼▼▼ Extended Debug.Print statements ▼▼▼
    Debug.Print Now & " M04.WriteErrorLog: ENTER. errorLevel='" & errorLevel & "', isDebugLog=" & isDebugLog
    If Not g_configSettings Is Nothing Then ' Check if g_configSettings itself is available
        Debug.Print Now & " M04.WriteErrorLog: Config Flags: EnableSheetLogging(O5)=" & g_configSettings.EnableSheetLogging & _
                      ", EnableErrorLogSheetOutput(O7)=" & g_configSettings.EnableErrorLogSheetOutput & _
                      ", DebugDetailLevel1Enabled(O4)=" & g_configSettings.DebugDetailLevel1Enabled
    Else
        Debug.Print Now & " M04.WriteErrorLog: g_configSettings is Nothing. Cannot check flags."
    End If

    If isDebugLog Then
        Debug.Print Now & " M04.WriteErrorLog: DEBUG log detected. Target: GenericLogSheet."
        If g_genericLogWorksheet Is Nothing Then
            Debug.Print Now & " M04.WriteErrorLog: g_genericLogWorksheet is Nothing."
        Else
            Debug.Print Now & " M04.WriteErrorLog: g_genericLogWorksheet is '" & g_genericLogWorksheet.Name & "'. NextRow: " & g_nextGenericLogRow
        End If

        If Not g_configSettings.EnableSheetLogging Then ' Check O5 for DEBUG logs
            Debug.Print Now & " M04.WriteErrorLog: Skipping DEBUG log to sheet as EnableSheetLogging (O5) is False."
            Exit Sub
        End If
        Set targetSheet = g_genericLogWorksheet
        targetNextRow = g_nextGenericLogRow
    Else ' Not a DEBUG_ log
        Debug.Print Now & " M04.WriteErrorLog: NON-DEBUG log detected. Target: ErrorLogSheet."
        If g_errorLogWorksheet Is Nothing Then
            Debug.Print Now & " M04.WriteErrorLog: g_errorLogWorksheet is Nothing."
        Else
            Debug.Print Now & " M04.WriteErrorLog: g_errorLogWorksheet is '" & g_errorLogWorksheet.Name & "'. NextRow: " & g_nextErrorLogRow
        End If

        If Not g_configSettings.EnableErrorLogSheetOutput Then ' Check O7 for ERROR logs
            Debug.Print Now & " M04.WriteErrorLog: Skipping ERROR log to sheet as EnableErrorLogSheetOutput (O7) is False."
            ' Fallback to Debug.Print for non-debug logs if sheet output is off
            Debug.Print Now & " WriteErrorLog FALLBACK (ErrorLogSheet Output Disabled O7=FALSE): Level=" & errorLevel & ", Mod=" & moduleN & ", Proc=" & procedureN & ", Msg=" & message
            If errNumber <> 0 Then Debug.Print "  > Err #: " & errNumber & " - " & errDescription
            Exit Sub
        End If
        Set targetSheet = g_errorLogWorksheet
        targetNextRow = g_nextErrorLogRow
    End If
    ' ▲▲▲ End of Extended Debug.Print statements ▲▲▲

    ' Conditional Debug.Print for actual call details (only if L1 is on, and it's about to write to sheet)
    If Not targetSheet Is Nothing And g_configSettings.DebugDetailLevel1Enabled Then
        Debug.Print Now & " WriteErrorLog CALLED (will attempt sheet write): Level=" & errorLevel & ", Mod=" & moduleN & ", Proc=" & procedureN & ", Msg=" & message & ", Err#=" & errNumber & ", ErrDesc=" & errDescription
    End If

    On Error GoTo ErrorHandler_WriteErrorLog

    If targetSheet Is Nothing Then
        ' This Debug.Print is a fallback if targetSheet ended up Nothing despite checks
        Debug.Print Now & " WriteErrorLog FALLBACK (targetSheet is Nothing just before write): Level=" & errorLevel & ", Mod=" & moduleN & ", Proc=" & procedureN & ", Msg=" & message
        If errNumber <> 0 And Not isDebugLog Then Debug.Print "  > Err #: " & errNumber & " - " & errDescription
        Exit Sub
    End If

    If targetNextRow <= 0 Then
        If Application.WorksheetFunction.CountA(targetSheet.Rows(1)) = 0 Then
             targetNextRow = 1
        Else
             targetNextRow = targetSheet.Cells(Rows.Count, "A").End(xlUp).Row + 1
        End If
        If targetNextRow <= 0 Then targetNextRow = 1
    End If

    If targetNextRow > targetSheet.Rows.Count Then
        Debug.Print Now & " WriteErrorLog HALT: targetNextRow (" & targetNextRow & ") exceeds sheet rows (" & targetSheet.Rows.Count & ") for sheet " & targetSheet.Name
        Exit Sub
    End If

    If g_configSettings.DebugDetailLevel1Enabled Then
        Debug.Print Now & " WriteErrorLog: Writing to Sheet '" & targetSheet.Name & "', Row: " & targetNextRow
    End If

    With targetSheet
        .Cells(targetNextRow, 1).value = Now()
        .Cells(targetNextRow, 2).value = errorLevel
        .Cells(targetNextRow, 3).value = moduleN
        .Cells(targetNextRow, 4).value = procedureN
        .Cells(targetNextRow, 5).value = message
        If errNumber <> 0 Then
            .Cells(targetNextRow, 6).value = errNumber
            .Cells(targetNextRow, 7).value = errDescription
        Else
            .Cells(targetNextRow, 6).value = vbNullString
            .Cells(targetNextRow, 7).value = vbNullString
        End If
    End With

    If isDebugLog Then
        g_nextGenericLogRow = targetNextRow + 1
    Else
        g_nextErrorLogRow = targetNextRow + 1
    End If
    Exit Sub

ErrorHandler_WriteErrorLog:
    ' Critical error within WriteErrorLog itself. Only use Debug.Print.
    Debug.Print Now & " CRITICAL ERROR in M04_LogWriter.WriteErrorLog ITSELF! Err# " & Err.Number & " - " & Err.Description
    Debug.Print "  Attempted Log: Level=" & errorLevel & ", Mod=" & moduleN & ", Proc=" & procedureN & ", Msg=" & message
    ' Do not try to call WriteErrorLog here.
    ' Resume Next or Exit Sub can be chosen. Exit Sub is safer to prevent loops.
    ' If this handler is reached, something is very wrong with the log sheet or Excel state.
End Sub

' Public Sub: WriteFilterLog (Revised to only log D-Section: Filter Conditions)
Public Sub WriteFilterLog(ByRef config As tConfigSettings, ByVal wb As Workbook)
    Dim funcName As String: funcName = "WriteFilterLog"
    Dim wsLog As Worksheet
    Dim nextLogWriteRow As Long

    If Not config.EnableSearchConditionLogSheetOutput Then Exit Sub ' Controlled by O6

    On Error GoTo ErrorHandler_WriteFilterLog
    Set wsLog = wb.Sheets(config.SearchConditionLogSheetName)
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

    Call WriteFilterLogEntry(wsLog, nextLogWriteRow, "--- D. フィルター条件 ---", "開始")

    If Trim(config.WorkerFilterLogic) <> "" And UCase(Trim(config.WorkerFilterLogic)) <> "AND" Then ' Log if not default "AND" or if explicitly set
        Call WriteFilterLogEntry(wsLog, nextLogWriteRow, "D.作業員フィルター検索論理", config.WorkerFilterLogic)
    End If

    If General_IsArrayInitialized(config.WorkerFilterList) Then
        Call WriteFilterLogArrayEntryIfSet(wsLog, nextLogWriteRow, "D.作業員フィルターリスト", config.WorkerFilterList)
    End If
    If General_IsArrayInitialized(config.Kankatsu1FilterList) Then
        Call WriteFilterLogArrayEntryIfSet(wsLog, nextLogWriteRow, "D.管内1フィルターリスト", config.Kankatsu1FilterList)
    End If
    If General_IsArrayInitialized(config.Kankatsu2FilterList) Then
        Call WriteFilterLogArrayEntryIfSet(wsLog, nextLogWriteRow, "D.管内2フィルターリスト", config.Kankatsu2FilterList)
    End If
    If Trim(config.Bunrui1Filter) <> "" Then
        Call WriteFilterLogEntry(wsLog, nextLogWriteRow, "D.分類1フィルター", config.Bunrui1Filter)
    End If
    If Trim(config.Bunrui2Filter) <> "" Then
        Call WriteFilterLogEntry(wsLog, nextLogWriteRow, "D.分類2フィルター", config.Bunrui2Filter)
    End If
    If Trim(config.Bunrui3Filter) <> "" Then
        Call WriteFilterLogEntry(wsLog, nextLogWriteRow, "D.分類3フィルター", config.Bunrui3Filter)
    End If
    If General_IsArrayInitialized(config.KoujiShuruiFilterList) Then
        Call WriteFilterLogArrayEntryIfSet(wsLog, nextLogWriteRow, "D.工事種類フィルターリスト", config.KoujiShuruiFilterList)
    End If
    If General_IsArrayInitialized(config.KoubanFilterList) Then
        Call WriteFilterLogArrayEntryIfSet(wsLog, nextLogWriteRow, "D.工番フィルターリスト", config.KoubanFilterList)
    End If
    If General_IsArrayInitialized(config.SagyoushuruiFilterList) Then
        Call WriteFilterLogArrayEntryIfSet(wsLog, nextLogWriteRow, "D.作業種類フィルターリスト", config.SagyoushuruiFilterList)
    End If
    If General_IsArrayInitialized(config.TantouFilterList) Then
        Call WriteFilterLogArrayEntryIfSet(wsLog, nextLogWriteRow, "D.担当の名前フィルターリスト", config.TantouFilterList)
    End If
    If Not config.IsNinzuFilterOriginallyEmpty Then ' Log if it was not originally empty (i.e. user set something, even if it's "0")
        Call WriteFilterLogEntry(wsLog, nextLogWriteRow, "D.人数フィルター", config.NinzuFilter)
    End If
    If Trim(config.SagyouKashoKindFilter) <> "" Then
        Call WriteFilterLogEntry(wsLog, nextLogWriteRow, "D.作業箇所の種類フィルター", config.SagyouKashoKindFilter)
    End If
    If General_IsArrayInitialized(config.SagyouKashoFilterList) Then
        Call WriteFilterLogArrayEntryIfSet(wsLog, nextLogWriteRow, "D.作業箇所フィルターリスト", config.SagyouKashoFilterList)
    End If
    Call WriteFilterLogEntry(wsLog, nextLogWriteRow, "--- D. フィルター条件 ---", "終了")

    Call WriteErrorLog("INFORMATION", MODULE_NAME, funcName, "フィルター条件ログの書き込みが完了しました。")
    Exit Sub
ErrorHandler_WriteFilterLog: ' Ensure this label exists
    Call WriteErrorLog("ERROR", MODULE_NAME, funcName, "フィルター条件ログ書き込み中にエラー。", Err.Number, Err.Description)
End Sub

' Public Sub: WriteOperationLog (New procedure for general operational logs)
Public Sub WriteOperationLog(ByRef config As tConfigSettings, ByVal wb As Workbook, Optional eventName As String = "", Optional eventDetail As String = "")
    Dim funcName As String: funcName = "WriteOperationLog"
    Dim wsLog As Worksheet
    Dim nextLogWriteRow As Long
    Dim i As Long ' Loop counter

    If Not config.EnableSheetLogging Then Exit Sub ' Controlled by O5

    On Error GoTo ErrorHandler_WriteOperationLog
    Set wsLog = wb.Sheets(config.LogSheetName) ' Output to Generic Log sheet (O42)
    If wsLog Is Nothing Then
        Call WriteErrorLog("ERROR", MODULE_NAME, funcName, "汎用ログシート「" & config.LogSheetName & "」が見つかりません。")
        Exit Sub
    End If

    If Application.WorksheetFunction.CountA(wsLog.Rows(1)) = 0 Then
         nextLogWriteRow = 1
    Else
         nextLogWriteRow = wsLog.Cells(Rows.Count, "A").End(xlUp).Row + 1
    End If
    If nextLogWriteRow <= 0 Then nextLogWriteRow = 1

    ' If eventName is provided, log it as a specific event
    If eventName <> "" Then
        Call WriteFilterLogEntry(wsLog, nextLogWriteRow, eventName, eventDetail) ' Using WriteFilterLogEntry for simplicity, adapt if needed
        Exit Sub ' For specific events, we don't log all settings again
    End If

    ' Initial logging of settings (Sections A, B, C, E, G)
    Call WriteFilterLogEntry(wsLog, nextLogWriteRow, "--- マクロ基本情報 ---", "開始")
    Call WriteFilterLogEntry(wsLog, nextLogWriteRow, "マクロ実行開始時刻", CStr(config.startTime))
    Call WriteFilterLogEntry(wsLog, nextLogWriteRow, "マクロファイル", config.ScriptFullName)
    Call WriteFilterLogEntry(wsLog, nextLogWriteRow, "設定ファイルシート", config.ConfigSheetFullName)
    Call WriteFilterLogEntry(wsLog, nextLogWriteRow, "デバッグモードフラグ(O3)", CStr(config.DebugModeFlag))
    Call WriteFilterLogEntry(wsLog, nextLogWriteRow, "デバッグ詳細出力レベル1(O4)", CStr(config.DebugDetailLevel1Enabled)) ' Verified: This line was already correct from previous subtask.
    Call WriteFilterLogEntry(wsLog, nextLogWriteRow, "--- マクロ基本情報 ---", "終了")

    Call WriteFilterLogEntry(wsLog, nextLogWriteRow, "--- A. 一般設定 (抜粋) ---", "開始")
    Call WriteFilterLogEntry(wsLog, nextLogWriteRow, "A.デフォルトフォルダパス(O12)", config.DefaultFolderPath)
    Call WriteFilterLogEntry(wsLog, nextLogWriteRow, "A.抽出結果出力シート名(O43)", config.OutputSheetName)
    Call WriteFilterLogEntry(wsLog, nextLogWriteRow, "A.工程パターンデータ取得方法(O122)", IIf(config.GetPatternDataMethod, "数式", "VBA"))
    Call WriteFilterLogEntry(wsLog, nextLogWriteRow, "--- A. 一般設定 (抜粋) ---", "終了")

    Call WriteFilterLogEntry(wsLog, nextLogWriteRow, "--- B. 工程表ファイル内 設定 ---", "開始")
    If General_IsArrayInitialized(config.TargetSheetNames) Then
        Call WriteFilterLogArrayEntry(wsLog, nextLogWriteRow, "B.検索対象シート名リスト(O66-O75)", config.TargetSheetNames)
    End If
    Call WriteFilterLogEntry(wsLog, nextLogWriteRow, "B.工程表ヘッダー行数(O87)", CStr(config.HeaderRowCount))
    Call WriteFilterLogEntry(wsLog, nextLogWriteRow, "B.1日の工程数(O114)", CStr(config.ProcessesPerDay))
    Call WriteFilterLogEntry(wsLog, nextLogWriteRow, "--- B. 工程表ファイル内 設定 ---", "終了")

    Call WriteFilterLogEntry(wsLog, nextLogWriteRow, "--- C. 工程パターン定義 (抜粋) ---", "開始")
    Call WriteFilterLogEntry(wsLog, nextLogWriteRow, "C.現在処理中パターン識別子(O126)", config.CurrentPatternIdentifier)
     If General_IsArrayInitialized(config.ProcessColCountSheetHeaders) Then
        Call WriteFilterLogArrayEntry(wsLog, nextLogWriteRow, "C.工程列数定義シートヘッダー(O128-X128)", config.ProcessColCountSheetHeaders)
    End If
    Call WriteFilterLogEntry(wsLog, nextLogWriteRow, "--- C. 工程パターン定義 (抜粋) ---", "終了")

    Call WriteFilterLogEntry(wsLog, nextLogWriteRow, "--- E. 処理対象ファイル定義 ---", "開始")
    If General_IsArrayInitialized(config.TargetFileFolderPaths) Then
         Call WriteFilterLogArrayEntry(wsLog, nextLogWriteRow, "E.処理対象ファイル/フォルダパスリスト(P557-P756)", config.TargetFileFolderPaths)
    End If
    If General_IsArrayInitialized(config.FilePatternIdentifiers) Then
         Call WriteFilterLogArrayEntry(wsLog, nextLogWriteRow, "E.適用工程パターン識別子リスト(Q557-Q756)", config.FilePatternIdentifiers)
    End If
    Call WriteFilterLogEntry(wsLog, nextLogWriteRow, "--- E. 処理対象ファイル定義 ---", "終了")

    Call WriteFilterLogEntry(wsLog, nextLogWriteRow, "--- G. 出力シート設定 ---", "開始")
    Call WriteFilterLogEntry(wsLog, nextLogWriteRow, "G.出力シートヘッダー行数(O811)", CStr(config.OutputHeaderRowCount))
    If General_IsArrayInitialized(config.OutputHeaderContents) Then
         Call WriteFilterLogArrayEntry(wsLog, nextLogWriteRow, "G.出力シートヘッダー内容(O812-O821)", config.OutputHeaderContents)
    End If
    Call WriteFilterLogEntry(wsLog, nextLogWriteRow, "G.出力データオプション(O1124)", config.OutputDataOption)
    Call WriteFilterLogEntry(wsLog, nextLogWriteRow, "--- G. 出力シート設定 ---", "終了")

    Call WriteErrorLog("INFORMATION", MODULE_NAME, funcName, "汎用動作ログの初期書き込みが完了しました。")
    Exit Sub
ErrorHandler_WriteOperationLog:
    Call WriteErrorLog("ERROR", MODULE_NAME, funcName, "汎用動作ログ書き込み中にエラー。", Err.Number, Err.Description)
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

    If Not General_IsArrayInitialized(arr) Then Exit Sub ' 配列でない場合は終了

    ' 要素が存在するかどうかを安全に確認
    Dim hasElements As Boolean
    hasElements = False ' Default to false
    On Error Resume Next ' LBound/UBoundでエラーが発生するケースを考慮
    If LBound(arr) <= UBound(arr) Then
        hasElements = True
    End If
    If Err.Number <> 0 Then
        hasElements = False ' LBound/UBoundでエラーなら要素なしとみなす
        Err.Clear
    End If
    On Error GoTo 0

    If Not hasElements Then Exit Sub ' 要素がなければ何もせず終了

    For i = LBound(arr) To UBound(arr)
        currentItemName = itemBaseName ' 配列の場合、要素ごとにインデックスを付けないシンプルなログ形式
        If Trim(arr(i)) <> "" Then
            ' 配列の各要素を個別の行として記録。項目名は同じitemBaseNameを使う。
             Call WriteFilterLogEntry(ws, nextRow, itemBaseName, arr(i))
        End If
    Next i
End Sub

Private Sub WriteFilterLogArrayEntryIfSet(ByVal ws As Worksheet, ByRef nextRow As Long, ByVal itemBaseName As String, ByRef arr() As String)
    Dim i As Long
    Dim hasActualValue As Boolean
    hasActualValue = False

    If Not General_IsArrayInitialized(arr) Then Exit Sub

    ' Check if there's any non-empty value in the array first
    For i = LBound(arr) To UBound(arr)
        If Trim(arr(i)) <> "" Then
            hasActualValue = True
            Exit For
        End If
    Next i

    If Not hasActualValue Then Exit Sub ' No actual values to log

    ' Log the base name once if there are values
    ' Call WriteFilterLogEntry(ws, nextRow, itemBaseName, "(以下リスト)") ' Optional: indicate list follows

    For i = LBound(arr) To UBound(arr)
        If Trim(arr(i)) <> "" Then
             Call WriteFilterLogEntry(ws, nextRow, itemBaseName & " (" & i & ")", arr(i)) ' Log with index
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
