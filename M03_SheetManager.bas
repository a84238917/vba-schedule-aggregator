' バージョン：v0.5.1
Option Explicit
' このモジュールは、ワークシートの管理（作成、クリア、存在確認など）を担当します。
' 特に、出力シートやログシートの準備に関連する機能を提供します。

Private Const MODULE_NAME As String = "M03_SheetManager"

' Public Sub: PrepareErrorLogSheet
' エラーログシートのみを準備し、グローバル変数 g_errorLogWorksheet を設定します。
Public Sub PrepareErrorLogSheet(ByRef config As tConfigSettings, ByVal wb As Workbook)
    Dim funcName As String: funcName = "PrepareErrorLogSheet"
    Dim wasCreated As Boolean

    If Not config.EnableErrorLogSheetOutput Then
        Set g_errorLogWorksheet = Nothing
        g_nextErrorLogRow = 1 ' Reset anyway
        Exit Sub
    End If

    On Error GoTo ErrorHandler_PrepareErrorLogSheet

    If Trim(config.ErrorLogSheetName) = "" Then
        Debug.Print Now & " CRITICAL: " & MODULE_NAME & "." & funcName & " - ErrorLogSheetName (Config O45) is empty."
        Set g_errorLogWorksheet = Nothing
        Exit Sub
    End If

    Set g_errorLogWorksheet = EnsureSheetExists(config.ErrorLogSheetName, wb, wasCreated)

    If g_errorLogWorksheet Is Nothing Then
        Debug.Print Now & " CRITICAL: " & MODULE_NAME & "." & funcName & " - ErrorLogシート「" & config.ErrorLogSheetName & "」の準備(EnsureSheetExists呼び出し)に失敗しました。EnsureSheetExists内のDebug.Printも確認してください。"
        Exit Sub
    End If

    Dim firstCellEmpty As Boolean
    On Error Resume Next
    firstCellEmpty = IsEmpty(g_errorLogWorksheet.Cells(1, 1).value)
    On Error GoTo ErrorHandler_PrepareErrorLogSheet

    If wasCreated Or firstCellEmpty Then
        Call WriteSheetHeaders(g_errorLogWorksheet, "ErrorLog", config)
    End If

    If Application.WorksheetFunction.CountA(g_errorLogWorksheet.Rows(1)) = 0 Then
         g_nextErrorLogRow = 1
    Else
         g_nextErrorLogRow = g_errorLogWorksheet.Cells(Rows.Count, "A").End(xlUp).Row + 1
    End If
    If g_nextErrorLogRow <= 0 Then g_nextErrorLogRow = 1

    Exit Sub

ErrorHandler_PrepareErrorLogSheet:
    Debug.Print Now & " CRITICAL ERROR in " & MODULE_NAME & "." & funcName & " (シート名: " & config.ErrorLogSheetName & "): Err# " & Err.Number & " - " & Err.Description
    Set g_errorLogWorksheet = Nothing
End Sub

' Public Sub: PrepareRemainingLogSheets
' エラーログ以外のログシート（検索条件ログ、汎用ログ）を準備します。
Public Sub PrepareRemainingLogSheets(ByRef config As tConfigSettings, ByVal wb As Workbook)
    Dim funcName As String: funcName = "PrepareRemainingLogSheets"
    Dim ws As Worksheet
    Dim wasCreated As Boolean
    Dim firstCellEmpty As Boolean
    On Error GoTo ErrorHandler_PrepareRemainingLogSheets ' Specific error handler

    ' 検索条件ログシートの準備
    If config.EnableSearchConditionLogSheetOutput Then ' ★追加
        If Trim(config.SearchConditionLogSheetName) <> "" Then
            Set ws = EnsureSheetExists(config.SearchConditionLogSheetName, wb, wasCreated)
            If Not ws Is Nothing Then
                On Error Resume Next
                firstCellEmpty = IsEmpty(ws.Cells(1, 1).Value)
                On Error GoTo ErrorHandler_PrepareRemainingLogSheets
                If wasCreated Or firstCellEmpty Then
                    Call WriteSheetHeaders(ws, "SearchLog", config)
                End If
            Else
                Call M04_LogWriter.WriteErrorLog("ERROR", MODULE_NAME, funcName, "検索条件ログシート「" & config.SearchConditionLogSheetName & "」の準備に失敗しました。")
            End If
        Else
             Call M04_LogWriter.WriteErrorLog("WARNING", MODULE_NAME, funcName, "検索条件ログシート名が設定されていませんが、出力は有効です(O6)。")
        End If
    End If ' ★追加
    Set ws = Nothing ' Reset for next sheet

    ' 汎用ログシートの準備 (Config O42: LogSheetName)
    Call M04_LogWriter.WriteErrorLog("DEBUG_M03", MODULE_NAME, funcName, "Attempting to prepare Generic Log Sheet. EnableSheetLogging (O5): " & config.EnableSheetLogging & ", LogSheetName (O42): " & config.LogSheetName)
    Debug.Print Now & " M03.PrepareRemainingLogSheets: Attempting Generic Log. EnableSheetLogging=" & config.EnableSheetLogging & ", SheetName='" & config.LogSheetName & "'"

    If config.EnableSheetLogging And Trim(config.LogSheetName) <> "" Then
        Set ws = EnsureSheetExists(config.LogSheetName, wb, wasCreated)
        Set g_genericLogWorksheet = ws

        If g_genericLogWorksheet Is Nothing Then
            Call M04_LogWriter.WriteErrorLog("ERROR", MODULE_NAME, funcName, "汎用ログシートオブジェクト(g_genericLogWorksheet)がNothingです。EnsureSheetExistsが失敗。シート名: " & config.LogSheetName)
            Debug.Print Now & " M03.PrepareRemainingLogSheets: g_genericLogWorksheet is Nothing after EnsureSheetExists for: " & config.LogSheetName
            g_nextGenericLogRow = 0
        Else
            Call M04_LogWriter.WriteErrorLog("INFORMATION", MODULE_NAME, funcName, "汎用ログシートオブジェクト(g_genericLogWorksheet)セット完了。シート名: " & g_genericLogWorksheet.Name)
            Debug.Print Now & " M03.PrepareRemainingLogSheets: g_genericLogWorksheet was SET to: '" & g_genericLogWorksheet.Name & "'"

            Dim firstCellEmpty_GL As Boolean
            On Error Resume Next
            firstCellEmpty_GL = IsEmpty(g_genericLogWorksheet.Cells(1, 1).value)
            On Error GoTo ErrorHandler_PrepareRemainingLogSheets

            If wasCreated Or firstCellEmpty_GL Then
                Debug.Print Now & " M03.PrepareRemainingLogSheets: Generic Log Sheet ('" & g_genericLogWorksheet.Name & "'), wasCreated=" & wasCreated & ", firstCellEmpty=" & firstCellEmpty_GL & ". Writing headers."
                Call WriteSheetHeaders(g_genericLogWorksheet, "GenericLog", config)
            Else
                Debug.Print Now & " M03.PrepareRemainingLogSheets: Generic Log Sheet ('" & g_genericLogWorksheet.Name & "'), wasCreated=" & wasCreated & ", firstCellEmpty=" & firstCellEmpty_GL & ". Headers NOT written."
            End If

            If Application.WorksheetFunction.CountA(g_genericLogWorksheet.Rows(1)) = 0 Then
                 g_nextGenericLogRow = 1
            Else
                 g_nextGenericLogRow = g_genericLogWorksheet.Cells(Rows.Count, "A").End(xlUp).Row + 1
            End If
            If g_nextGenericLogRow <= 0 Then g_nextGenericLogRow = 1
            Debug.Print Now & " M03.PrepareRemainingLogSheets: g_nextGenericLogRow for '" & g_genericLogWorksheet.Name & "' initialized to: " & g_nextGenericLogRow
            Call M04_LogWriter.WriteErrorLog("DEBUG_M03", MODULE_NAME, funcName, "g_nextGenericLogRow for '" & g_genericLogWorksheet.Name & "' initialized to: " & g_nextGenericLogRow)
        End If
    ElseIf config.EnableSheetLogging And Trim(config.LogSheetName) = "" Then
         Call M04_LogWriter.WriteErrorLog("WARNING", MODULE_NAME, funcName, "汎用ログシート名(O42)が設定されていませんが、シートログ(O5)は有効です。")
         Debug.Print Now & " M03.PrepareRemainingLogSheets: Generic Log Sheet name is empty but EnableSheetLogging is True."
         Set g_genericLogWorksheet = Nothing
         g_nextGenericLogRow = 0
    Else ' EnableSheetLogging is False
         Call M04_LogWriter.WriteErrorLog("DEBUG_M03", MODULE_NAME, funcName, "汎用ログシート(O42)への出力は無効です (O5がFALSE)。")
         Debug.Print Now & " M03.PrepareRemainingLogSheets: Generic Log Sheet output is disabled (EnableSheetLogging is False)."
         Set g_genericLogWorksheet = Nothing
         g_nextGenericLogRow = 0
    End If
    Set ws = Nothing

    Exit Sub
ErrorHandler_PrepareRemainingLogSheets:
     Call M04_LogWriter.WriteErrorLog("CRITICAL", MODULE_NAME, funcName, "残りのログシート準備中にエラー。", Err.Number, Err.Description)
End Sub

' Public Sub: PrepareOutputSheet
' 指定された出力シートを準備します。
Public Sub PrepareOutputSheet(ByRef config As tConfigSettings, ByVal wb As Workbook, ByRef nextRow As Long)
    Dim funcName As String: funcName = "PrepareOutputSheet"
    Dim wsOutput As Worksheet
    Dim wasCreated As Boolean

    On Error GoTo ErrorHandler_PrepareOutputSheet ' Specific error handler for this sub
    On Error GoTo ErrorHandler_PrepareOutputSheet ' Specific error handler for this sub

    If Trim(config.OutputSheetName) = "" Then
        Call M04_LogWriter.WriteErrorLog("CRITICAL", MODULE_NAME, funcName, "出力シート名が設定されていません。処理を続行できません。")
        nextRow = 1 ' 安全なフォールバック
        Exit Sub
    End If

    Set wsOutput = EnsureSheetExists(config.OutputSheetName, wb, wasCreated)
    If wsOutput Is Nothing Then
        Call M04_LogWriter.WriteErrorLog("CRITICAL", MODULE_NAME, funcName, "出力シート「" & config.OutputSheetName & "」の準備に失敗しました。")
        nextRow = 1 ' 安全なフォールバック
        Exit Sub
    End If

    If wasCreated Or UCase(Trim(config.OutputDataOption)) = "リセット" Then
        Dim clearStartRow As Long
        clearStartRow = 1
        If config.OutputHeaderRowCount > 0 Then ' ヘッダーがある場合
            clearStartRow = config.OutputHeaderRowCount + 1
        End If

        ' ヘッダー行より下の行をクリア
        If clearStartRow <= wsOutput.Rows.Count Then
             wsOutput.Rows(clearStartRow & ":" & wsOutput.Rows.Count).ClearContents
        End If

        Call WriteSheetHeaders(wsOutput, "Output", config) ' ヘッダーを書き込む
        nextRow = config.OutputHeaderRowCount + 1
    Else ' "引継ぎ" またはその他の場合 (wasCreated = False で "リセット" でない)
        If config.OutputHeaderRowCount > 0 Then
            nextRow = wsOutput.Cells(Rows.Count, 1).End(xlUp).Row + 1
            If nextRow <= config.OutputHeaderRowCount Then
                nextRow = config.OutputHeaderRowCount + 1
            End If
        Else
            nextRow = wsOutput.Cells(Rows.Count, 1).End(xlUp).Row
            ' データが全くないシート(A1も空)の場合、End(xlUp)は1を返すことがあるので、
            ' A1が空ならnextRowは1、そうでなければ+1する
            If wsOutput.Cells(1,1).Value = "" And nextRow = 1 Then
                ' nextRow is already 1
            ElseIf nextRow = 1 And wsOutput.Cells(1,1).Value <> "" Then ' A1にデータあり
                nextRow = 2
            ElseIf nextRow > 1 Then
                 nextRow = nextRow + 1
            End If
        End If
    End If

    Set wsOutput = Nothing
    Exit Sub

ErrorHandler_PrepareOutputSheet: ' Specific error handler for this sub
    Call M04_LogWriter.WriteErrorLog("CRITICAL", MODULE_NAME, funcName, "出力シート準備中にエラーが発生しました。", Err.Number, Err.Description)
    nextRow = 1 ' フォールバック値
End Sub


' Private Function: EnsureSheetExists (remains mostly the same, ensure M04_LogWriter calls are safe)
Private Function EnsureSheetExists(ByVal sheetName As String, ByVal wb As Workbook, ByRef wasCreated As Boolean) As Worksheet
    Dim funcName As String: funcName = "EnsureSheetExists"
    Dim ws As Worksheet
    wasCreated = False

    On Error Resume Next ' For the initial Set ws attempt
    Set ws = wb.Sheets(sheetName)
    On Error GoTo 0 ' Reset error handling before specific error checks

    If ws Is Nothing Then
        If wb.ReadOnly Then
            ' Use Debug.Print if g_errorLogWorksheet might not be ready
            Debug.Print Now & " ERROR: " & MODULE_NAME & "." & funcName & " - ブック「" & wb.Name & "」は読み取り専用のため、シート「" & sheetName & "」を作成できません。"
            Set EnsureSheetExists = Nothing
            Exit Function
        End If

        On Error Resume Next ' For Add/Name operations
        Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        If Err.Number <> 0 Then
             Debug.Print Now & " ERROR: " & MODULE_NAME & "." & funcName & " - シート追加(Add)中にエラー。 Err# " & Err.Number & " - " & Err.Description
             Set EnsureSheetExists = Nothing
             Exit Function
        End If

        ws.Name = sheetName
        If Err.Number <> 0 Then
             Debug.Print Now & " ERROR: " & MODULE_NAME & "." & funcName & " - シート名変更(Name to '" & sheetName & "')中にエラー。 Err# " & Err.Number & " - " & Err.Description
             ' Attempt to delete the added sheet if renaming fails
             Application.DisplayAlerts = False
             ws.Delete
             Application.DisplayAlerts = True
             Set EnsureSheetExists = Nothing
             Exit Function
        End If
        On Error GoTo 0 ' Clear error handling
        wasCreated = True
        Set EnsureSheetExists = ws
    Else
        Set EnsureSheetExists = ws
    End If
End Function

' Private Sub: WriteSheetHeaders (Modify for "GenericLog" and header writing logic)
Private Sub WriteSheetHeaders(ByVal ws As Worksheet, ByVal sheetType As String, ByRef config As tConfigSettings)
    Dim funcName As String: funcName = "WriteSheetHeaders"
    On Error GoTo ErrorHandler_WriteSheetHeaders ' Specific error handler

    ' Clear the first row before writing headers to prevent partial overwrites if called on existing sheet
    ' However, this should ideally be called only if wasCreated or first row is truly empty.
    ' The current logic in PrepareErrorLogSheet and PrepareRemainingLogSheets handles this decision.
    ' So, we assume when this sub is called, it's appropriate to write headers.
    ' For safety, one could add 'ws.Rows(1).ClearContents' here if there's doubt.

    Select Case sheetType
        Case "ErrorLog"
            ws.Range("A1").value = "日時"
            ws.Range("B1").value = "レベル"
            ws.Range("C1").value = "モジュール"
            ws.Range("D1").value = "プロシージャ"
            ws.Range("E1").value = "メッセージ"
            ws.Range("F1").value = "エラー番号"
            ws.Range("G1").value = "エラー詳細"
        Case "SearchLog"
            ws.Range("A1").value = "実行日時"
            ws.Range("B1").value = "設定項目"
            ws.Range("C1").value = "設定値"
        Case "GenericLog" ' ★追加: 汎用ログシート用ヘッダー
            ws.Range("A1").value = "日時"
            ws.Range("B1").value = "レベル"
            ws.Range("C1").value = "モジュール"
            ws.Range("D1").value = "プロシージャ"
            ws.Range("E1").value = "メッセージ"
        Case "Output"
            Dim r As Long, c As Long ' ★ Ensure this declaration is here
            Dim singleRowHeaders() As String

            ' The existing conditional structure is kept, only the Dim statement is ensured.
            ' The prompt's target structure for If conditions is slightly different but not part of this specific fix.
            If General_IsArrayInitialized(config.OutputHeaderContents) And config.OutputHeaderRowCount > 0 Then
                For r = 1 To config.OutputHeaderRowCount
                    If r >= LBound(config.OutputHeaderContents) And r <= UBound(config.OutputHeaderContents) Then
                        singleRowHeaders = Split(config.OutputHeaderContents(r), vbTab) ' タブ区切りを想定
                        For c = 0 To UBound(singleRowHeaders)
                           ws.Cells(r, c + 1).value = Trim(singleRowHeaders(c))
                        Next c
                    Else
                        Call M04_LogWriter.WriteErrorLog("WARNING", MODULE_NAME, funcName, "OutputHeaderContentsの要素数(" & UBound(config.OutputHeaderContents) & ")が行数指定(" & r & ")と一致しません。")
                    End If
                Next r
            Else
                Call M04_LogWriter.WriteErrorLog("WARNING", MODULE_NAME, funcName, "出力シートのヘッダー内容が未設定またはヘッダー行数が0です。")
            End If
        Case Else
            Call M04_LogWriter.WriteErrorLog("WARNING", MODULE_NAME, funcName, "不明なシートタイプ「" & sheetType & "」。ヘッダー書き込み不可。")
    End Select
    Exit Sub

ErrorHandler_WriteSheetHeaders:
    Call M04_LogWriter.WriteErrorLog("ERROR", MODULE_NAME, funcName, sheetType & "シートのヘッダー書き込み中にエラー。", Err.Number, Err.Description)
End Sub

' Public Sub: PrepareOutputSheet (Ensure it uses updated WriteSheetHeaders logic if applicable)
' This sub seems to manage its own header writing via WriteSheetHeaders("Output",...)
' The condition for clearing (wasCreated Or UCase(Trim(config.OutputDataOption)) = "リセット")
' and then calling WriteSheetHeaders seems okay.

' Comment out or delete original PrepareSheets
' Public Sub PrepareSheets(ByRef config As tConfigSettings, ByVal wb As Workbook)
' ...
' End Sub

' General_IsArrayInitialized function remains the same
Public Function General_IsArrayInitialized(arr As Variant) As Boolean
    If Not IsArray(arr) Then Exit Function
    On Error Resume Next
    Dim lBoundCheck As Long: lBoundCheck = LBound(arr)
    If Err.Number = 0 Then General_IsArrayInitialized = True
    On Error GoTo 0
End Function
