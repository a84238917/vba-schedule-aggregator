' バージョン：v0.5.1
Option Explicit
' このモジュールは、ワークシートの管理（作成、クリア、存在確認など）を担当します。
' 特に、出力シートやログシートの準備に関連する機能を提供します。

Private Const MODULE_NAME As String = "M03_SheetManager"

' Public Sub: PrepareSheets
' マクロ実行に必要な基本ログシート（エラーログ、検索条件ログ）を準備します。
Public Sub PrepareSheets(ByRef config As tConfigSettings, ByVal wb As Workbook)
    Dim funcName As String: funcName = "PrepareSheets"
    Dim wasCreated As Boolean

    On Error GoTo ErrorHandler

    ' エラーログシートの準備
    If Trim(config.ErrorLogSheetName) <> "" Then
        Set g_errorLogWorksheet = EnsureSheetExists(config.ErrorLogSheetName, wb, wasCreated)
        If Not g_errorLogWorksheet Is Nothing Then
            If wasCreated Then
                Call WriteSheetHeaders(g_errorLogWorksheet, "ErrorLog", config)
            End If
            ' g_nextErrorLogRow の初期化 (M01_MainControlで行うか、ここで行うか設計による)
            ' ここで初期化する場合:
            If Application.WorksheetFunction.CountA(g_errorLogWorksheet.Rows(1)) = 0 Then ' 1行目が空なら
                 g_nextErrorLogRow = 1
            Else
                 g_nextErrorLogRow = g_errorLogWorksheet.Cells(Rows.Count, "A").End(xlUp).Row + 1
            End If

        Else
             Call M04_LogWriter.WriteErrorLog("CRITICAL", MODULE_NAME, funcName, "エラーログシート「" & config.ErrorLogSheetName & "」の準備に失敗しました。")
        End If
    Else
        Call M04_LogWriter.WriteErrorLog("WARNING", MODULE_NAME, funcName, "エラーログシート名が設定されていません。エラーログは記録されません。")
    End If

    ' 検索条件ログシートの準備
    Dim searchLogSht As Worksheet
    If Trim(config.SearchConditionLogSheetName) <> "" Then
        Set searchLogSht = EnsureSheetExists(config.SearchConditionLogSheetName, wb, wasCreated)
        If Not searchLogSht Is Nothing And wasCreated Then
            Call WriteSheetHeaders(searchLogSht, "SearchLog", config)
        ElseIf searchLogSht Is Nothing Then
            Call M04_LogWriter.WriteErrorLog("CRITICAL", MODULE_NAME, funcName, "検索条件ログシート「" & config.SearchConditionLogSheetName & "」の準備に失敗しました。")
        End If
    Else
        Call M04_LogWriter.WriteErrorLog("WARNING", MODULE_NAME, funcName, "検索条件ログシート名が設定されていません。検索条件ログは記録されません。")
    End If

    Set searchLogSht = Nothing
    Exit Sub

ErrorHandler:
    Call M04_LogWriter.WriteErrorLog("CRITICAL", MODULE_NAME, funcName, "ログシート準備中にエラーが発生しました。", Err.Number, Err.Description)
    Set g_errorLogWorksheet = Nothing
End Sub

' Public Sub: PrepareOutputSheet
' 指定された出力シートを準備します。
Public Sub PrepareOutputSheet(ByRef config As tConfigSettings, ByVal wb As Workbook, ByRef nextRow As Long)
    Dim funcName As String: funcName = "PrepareOutputSheet"
    Dim wsOutput As Worksheet
    Dim wasCreated As Boolean

    On Error GoTo ErrorHandler

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

ErrorHandler:
    Call M04_LogWriter.WriteErrorLog("CRITICAL", MODULE_NAME, funcName, "出力シート準備中にエラーが発生しました。", Err.Number, Err.Description)
    nextRow = 1 ' フォールバック値
End Sub


' Private Function: EnsureSheetExists
' 指定された名前のシートが存在することを確認し、なければ作成します。
Private Function EnsureSheetExists(ByVal sheetName As String, ByVal wb As Workbook, ByRef wasCreated As Boolean) As Worksheet
    Dim funcName As String: funcName = "EnsureSheetExists"
    Dim ws As Worksheet
    wasCreated = False

    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    On Error GoTo ErrorHandler

    If ws Is Nothing Then
        If wb.ReadOnly Then ' 読み取り専用ブックにはシートを追加できない
            Call M04_LogWriter.WriteErrorLog("ERROR", MODULE_NAME, funcName, "ブック「" & wb.Name & "」は読み取り専用のため、シート「" & sheetName & "」を作成できません。", 0, "Read-only Workbook")
            Set EnsureSheetExists = Nothing
            Exit Function
        End If
        On Error GoTo CreateErrorHandler ' AddSheetでエラーが発生した場合のため
        Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.Name = sheetName
        wasCreated = True
        Set EnsureSheetExists = ws
    Else
        Set EnsureSheetExists = ws
    End If
    Exit Function

CreateErrorHandler:
    Call M04_LogWriter.WriteErrorLog("ERROR", MODULE_NAME, funcName, "シート「" & sheetName & "」の新規作成(Add/Name)中にエラー。", Err.Number, Err.Description)
    Set EnsureSheetExists = Nothing
    Exit Function

ErrorHandler:
    Call M04_LogWriter.WriteErrorLog("ERROR", MODULE_NAME, funcName, "シート「" & sheetName & "」の確認中に予期せぬエラー。", Err.Number, Err.Description)
    Set EnsureSheetExists = Nothing
End Function

' Private Sub: WriteSheetHeaders
' シートにヘッダーを書き込みます。
Private Sub WriteSheetHeaders(ByVal ws As Worksheet, ByVal sheetType As String, ByRef config As tConfigSettings)
    Dim funcName As String: funcName = "WriteSheetHeaders"
    On Error GoTo ErrorHandler

    Select Case sheetType
        Case "ErrorLog"
            ws.Range("A1").Value = "日時"
            ws.Range("B1").Value = "レベル"
            ws.Range("C1").Value = "モジュール"
            ws.Range("D1").Value = "プロシージャ"
            ws.Range("E1").Value = "メッセージ"
            ws.Range("F1").Value = "エラー番号"
            ws.Range("G1").Value = "エラー詳細"
        Case "SearchLog"
            ws.Range("A1").Value = "実行日時"
            ws.Range("B1").Value = "設定項目"
            ws.Range("C1").Value = "設定値"
        Case "Output"
            If General_IsArrayInitialized(config.OutputHeaderContents) And config.OutputHeaderRowCount > 0 Then
                Dim r As Long, c As Long
                Dim singleRowHeaders() As String
                For r = 1 To config.OutputHeaderRowCount
                    If r >= LBound(config.OutputHeaderContents) And r <= UBound(config.OutputHeaderContents) Then
                        singleRowHeaders = Split(config.OutputHeaderContents(r), vbTab) ' タブ区切りを想定
                        For c = 0 To UBound(singleRowHeaders)
                           ws.Cells(r, c + 1).Value = Trim(singleRowHeaders(c))
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

ErrorHandler:
    Call M04_LogWriter.WriteErrorLog("ERROR", MODULE_NAME, funcName, sheetType & "シートのヘッダー書き込み中にエラー。", Err.Number, Err.Description)
End Sub

Public Function General_IsArrayInitialized(arr As Variant) As Boolean
    If Not IsArray(arr) Then Exit Function
    On Error Resume Next
    Dim lBoundCheck As Long: lBoundCheck = LBound(arr)
    If Err.Number = 0 Then General_IsArrayInitialized = True
    On Error GoTo 0
End Function
