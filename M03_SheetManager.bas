' バージョン：v0.5.0
Option Explicit
' このモジュールは、マクロが使用する各種ワークシート（出力シート、ログシートなど）の準備、検証、および管理を担当します。

Private Function SheetManager_IsArrayInitialized(arr As Variant) As Boolean
    ' 配列が有効に初期化されているか（少なくとも1つの要素を持つか）を確認します。
    ' Variant型が配列でない場合、または配列であっても要素が割り当てられていない場合（Dim arr() のみでReDimされていない状態など）はFalseを返します。
    On Error GoTo NotAnArrayOrNotInitialized_SM
    If IsArray(arr) Then
        Dim lBoundCheck As Long
        lBoundCheck = LBound(arr) ' 配列がReDimされていれば、LBoundはエラーにならない (空でも ReDim arr(0 To -1) など)
        SheetManager_IsArrayInitialized = True ' LBoundがエラーを起こさなければ、配列は有効（空でもReDimされていればOK）
        Exit Function
    End If
NotAnArrayOrNotInitialized_SM:
    SheetManager_IsArrayInitialized = False
End Function

Private Function GetHeaderRowCountForSheet(targetSheet As Worksheet, ByRef config As tConfigSettings, ByVal mainWorkbook As Workbook) As Long
    ' 指定されたシートのヘッダー行数をConfig設定に基づいて取得します。主に出力シート用です。
    ' Arguments:
    '   targetSheet: 対象のワークシート
    '   config: 設定情報
    '   mainWorkbook: ログ出力用のメインワークブック

    If targetSheet Is Nothing Then
        GetHeaderRowCountForSheet = 0
        If DEBUG_MODE_ERROR Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - ERROR: M03_SheetManager.GetHeaderRowCountForSheet - targetSheet is Nothing."
        Exit Function
    End If

    ' Check if essential config members are populated before use
    If Len(Trim(CStr(config.OutputSheetName))) = 0 Then
        If DEBUG_MODE_ERROR Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - ERROR: M03_SheetManager.GetHeaderRowCountForSheet - config.OutputSheetName is empty."
        ' Attempt to log this problem. ErrorLogSheetName might also be empty if config loading failed badly.
        Call M04_LogWriter.SafeWriteErrorLog("ERROR", mainWorkbook, config.ErrorLogSheetName, "M03_SheetManager", "GetHeaderRowCountForSheet", "Configメンバー OutputSheetName が未設定または空です。ヘッダー行数を正しく特定できません。", 0, "")
        GetHeaderRowCountForSheet = 0 ' Default to 0 rows if essential config is missing
        Exit Function
    End If

    ' config.OutputHeaderRowCount is a Long and will be 0 if not set.
    ' M02_ConfigReader validates it to be within 0-10 (or 1-10 if headers are mandatory).
    ' If it's 0, it means no headers are configured, which is a valid state.

    If targetSheet.Name = config.OutputSheetName Then
        GetHeaderRowCountForSheet = config.OutputHeaderRowCount
    Else
        GetHeaderRowCountForSheet = 0 ' Not the specifically configured output sheet
    End If

    ' Ensure non-negative, though M02 should have validated this.
    If GetHeaderRowCountForSheet < 0 Then GetHeaderRowCountForSheet = 0

    If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE: M03_SheetManager.GetHeaderRowCountForSheet - Sheet: " & targetSheet.Name & ", Configured OutputSheetName: " & config.OutputSheetName & ", Returned HeaderRows: " & GetHeaderRowCountForSheet
End Function

Private Function EnsureSheetExists(targetWorkbook As Workbook, sheetNameToEnsure As String, ByRef config As tConfigSettings, callerFuncName As String, createHeaders As Boolean) As Worksheet
    ' 指定されたワークブック内に特定の名前のシートが存在するか確認し、存在しない場合は作成します。
    ' ヘッダー作成が要求された場合、シート名に応じて適切なヘッダーを書き込みます。
    ' Arguments:
    '   targetWorkbook: 対象のワークブック
    '   sheetNameToEnsure: 存在を確認または作成するシートの名前
    '   config: (ByRef) 設定情報を保持するtConfigSettings型の変数
    '   callerFuncName: この関数を呼び出した関数名 (エラーログ用)
    '   createHeaders: (Boolean) Trueの場合、新規作成時にヘッダーを書き込む
    ' Returns:
    '   Worksheetオブジェクト (成功時)、Nothing (失敗時)

    If Trim(sheetNameToEnsure) = "" Then
        If DEBUG_MODE_ERROR Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - ERROR: M03_SheetManager.EnsureSheetExists - sheetNameToEnsure is empty. Called from: " & callerFuncName
        Set EnsureSheetExists = Nothing
        Exit Function
    End If

    Dim ws As Worksheet

    On Error Resume Next ' シートの存在確認に関するエラーを一旦無視
    Set ws = targetWorkbook.Sheets(sheetNameToEnsure)
    On Error GoTo EnsureSheetExists_Error ' 通常のエラーハンドリングに戻す

    If ws Is Nothing Then ' シートが存在しない場合
        If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M03_SheetManager.EnsureSheetExists - Sheet '" & sheetNameToEnsure & "' not found. Creating new sheet."
        On Error GoTo CreateSheet_Error ' シート作成に特化したエラーハンドリング
        Set ws = targetWorkbook.Sheets.Add(After:=targetWorkbook.Sheets(targetWorkbook.Sheets.Count))
        ws.Name = sheetNameToEnsure
        On Error GoTo EnsureSheetExists_Error ' 通常のエラーハンドリングに戻す

        If createHeaders Then
            If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M03_SheetManager.EnsureSheetExists - Creating headers for new sheet: " & sheetNameToEnsure
            If sheetNameToEnsure = config.ErrorLogSheetName Then
                ws.Cells(1, 1).Value = "重要度"       ' New Column A
                ws.Cells(1, 2).Value = "発生日時"     ' Old A -> New B
                ws.Cells(1, 3).Value = "モジュール"   ' Old B -> New C
                ws.Cells(1, 4).Value = "プロシージャ" ' Old C -> New D
                ws.Cells(1, 5).Value = "関連情報"     ' Old D -> New E
                ws.Cells(1, 6).Value = "エラー番号"   ' Old E -> New F
                ws.Cells(1, 7).Value = "エラー内容"   ' Old F -> New G
                ws.Cells(1, 8).Value = "対処内容"     ' Old G -> New H
                ws.Cells(1, 9).Value = "変数情報"     ' Old H -> New I
            ElseIf sheetNameToEnsure = config.SearchConditionLogSheetName Then
                ws.Cells(1, 1).Value = "実行日時"
                ws.Cells(1, 2).Value = "フィルター項目"
                ws.Cells(1, 3).Value = "条件"
                ws.Cells(1, 4).Value = "備考"
            ElseIf sheetNameToEnsure = config.OutputSheetName Then
                ' G-2. 出力シートヘッダー内容 (config.OutputHeaderContents)
                If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE: M03_SheetManager.EnsureSheetExists - Creating headers for Output Sheet: '" & sheetNameToEnsure & "'"
                Dim r As Long, c As Long
                Dim headerParts() As String

                If Not (config.OutputHeaderRowCount > 0 And SheetManager_IsArrayInitialized(config.OutputHeaderContents)) Then
                    If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE: M03_SheetManager.EnsureSheetExists - OutputHeaderRowCount is " & config.OutputHeaderRowCount & " or OutputHeaderContents not initialized. No headers written for " & sheetNameToEnsure
                    Call M04_LogWriter.SafeWriteErrorLog("WARNING", targetWorkbook, config.ErrorLogSheetName, "M03_SheetManager", "EnsureSheetExists", "出力シート「" & sheetNameToEnsure & "」のヘッダー情報 (OutputHeaderRowCount/OutputHeaderContents) がConfigに正しく設定されていません。ヘッダーは作成されません。", 0, "")
                    ' Allow sheet to be created empty, do not Exit Function here if ws was newly created.
                Else
                    ' Existing header writing loop
                    For r = 1 To config.OutputHeaderRowCount
                        If r <= UBound(config.OutputHeaderContents) And r >= LBound(config.OutputHeaderContents) Then
                            If Len(config.OutputHeaderContents(r)) > 0 Then
                                headerParts = Split(config.OutputHeaderContents(r), vbTab)
                                For c = 0 To UBound(headerParts)
                                    ws.Cells(r, c + 1).Value = headerParts(c)
                                Next c
                            Else ' Empty string in OutputHeaderContents(r) - write single empty cell to make row used
                                ws.Cells(r, 1).Value = ""
                            End If
                        End If
                    Next r
                End If
            End If
        End If
    Else
        If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M03_SheetManager.EnsureSheetExists - Sheet '" & sheetNameToEnsure & "' already exists."
    End If

    Set EnsureSheetExists = ws
    Exit Function

CreateSheet_Error:
    If DEBUG_MODE_ERROR Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - ERROR: M03_SheetManager.EnsureSheetExists (CreateSheet_Error) - Error " & Err.Number & ": " & Err.Description & " while trying to create/name sheet '" & sheetNameToEnsure & "'"
    Call M04_LogWriter.SafeWriteErrorLog("ERROR", targetWorkbook, config.ErrorLogSheetName, "M03_SheetManager", "EnsureSheetExists (CreateSheet_Error)", "シート '" & sheetNameToEnsure & "' の作成または命名に失敗しました。", Err.Number, Err.Description)
    Set EnsureSheetExists = Nothing
    Exit Function

EnsureSheetExists_Error:
    If DEBUG_MODE_ERROR Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - ERROR: M03_SheetManager.EnsureSheetExists - Error " & Err.Number & ": " & Err.Description & " (Caller: " & callerFuncName & ")"
    Call M04_LogWriter.SafeWriteErrorLog("ERROR", targetWorkbook, config.ErrorLogSheetName, "M03_SheetManager", "EnsureSheetExists", "シート '" & sheetNameToEnsure & "' の確認/作成中に予期せぬエラー (呼び出し元: " & callerFuncName & ")", Err.Number, Err.Description)
    Set EnsureSheetExists = Nothing
End Function

Public Function PrepareSheets(ByRef config As tConfigSettings, ByVal targetWorkbook As Workbook) As Boolean
    ' エラーログシートおよび検索条件ログシートを準備（存在確認または新規作成）し、
    ' グローバルエラーログ関連変数 (g_errorLogWorksheet, g_nextErrorLogRow) を設定します。
    ' Arguments:
    '   config: (ByRef) 設定情報を保持するtConfigSettings型の変数
    '   targetWorkbook: 対象のワークブック
    ' Returns:
    '   Boolean: True (成功時), False (失敗時)

    Dim wsErr As Worksheet
    Dim wsFilter As Worksheet

    PrepareSheets = False ' Default to failure
    On Error GoTo PrepareSheets_Error

    If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M03_SheetManager.PrepareSheets - Starting sheet preparation."

    ' エラーログシートの準備
    If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M03_SheetManager.PrepareSheets - Preparing Error Log Sheet: '" & config.ErrorLogSheetName & "'"
    Set wsErr = EnsureSheetExists(targetWorkbook, config.ErrorLogSheetName, config, "PrepareSheets", True)

    If Not wsErr Is Nothing Then ' Ensure wsErr is valid before using it
        Set g_errorLogWorksheet = wsErr
        ' Calculate g_nextErrorLogRow based on content of column A in g_errorLogWorksheet
        If Application.WorksheetFunction.CountA(g_errorLogWorksheet.Columns(1)) = 0 Then
            ' Column A is completely empty
            g_nextErrorLogRow = 1
        Else
            ' Column A has some data, find the last cell with data and add 1
            g_nextErrorLogRow = g_errorLogWorksheet.Cells(g_errorLogWorksheet.Rows.Count, 1).End(xlUp).Row + 1
        End If
        ' If headers were just written by EnsureSheetExists, and CountA(Columns(1)) found only the header,
        ' End(xlUp).Row would be 1, so g_nextErrorLogRow becomes 2, which is correct.
        ' If the sheet was truly empty and headers were just written, CountA(Columns(1)) is 1.
        ' If the sheet was truly empty and no headers (e.g. createHeaders = False), CountA is 0, next row is 1.
        If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M03_SheetManager.PrepareSheets - g_nextErrorLogRow determined as: " & g_nextErrorLogRow
    Else
        ' This case should ideally be caught by "If wsErr Is Nothing Then" block earlier,
        ' but as a safeguard:
        If DEBUG_MODE_ERROR Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - ERROR: M03_SheetManager.PrepareSheets - wsErr object was Nothing when trying to set g_nextErrorLogRow."
        ' SafeWriteErrorLog is called by EnsureSheetExists on failure to create, or here if EnsureSheetExists returns Nothing for other reasons.
        Call M04_LogWriter.SafeWriteErrorLog("CRITICAL", targetWorkbook, config.ErrorLogSheetName, "M03_SheetManager", "PrepareSheets", "エラーログシートの準備に失敗しました(wsErr is Nothing)。", 0, "EnsureSheetExistsがNothingを返しました")
        PrepareSheets = False ' Explicitly set to false as a critical part failed
        Exit Function
    End If

    ' 検索条件ログシートの準備
    If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M03_SheetManager.PrepareSheets - Preparing Filter Log Sheet: '" & config.SearchConditionLogSheetName & "'"
    Set wsFilter = EnsureSheetExists(targetWorkbook, config.SearchConditionLogSheetName, config, "PrepareSheets", True)
    If wsFilter Is Nothing Then
        If DEBUG_MODE_ERROR Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - ERROR: M03_SheetManager.PrepareSheets - Failed to prepare Search Condition Log Sheet."
        Call M04_LogWriter.SafeWriteErrorLog("ERROR", targetWorkbook, config.ErrorLogSheetName, "M03_SheetManager", "PrepareSheets", "検索条件ログシートの準備に失敗しました。", 0, "EnsureSheetExistsがNothingを返しました")
        Exit Function ' Returns False
    End If
    ' If wsFilter is not Nothing, it implies success for this part.
    If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M03_SheetManager.PrepareSheets - Search Condition Log Sheet ready."


    PrepareSheets = True ' All successful
    If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M03_SheetManager.PrepareSheets - Sheet preparation finished successfully."
    Exit Function

PrepareSheets_Error:
    If DEBUG_MODE_ERROR Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - ERROR: M03_SheetManager.PrepareSheets - Error " & Err.Number & ": " & Err.Description
    Call M04_LogWriter.SafeWriteErrorLog("ERROR", targetWorkbook, config.ErrorLogSheetName, "M03_SheetManager", "PrepareSheets", "PrepareSheets内で予期せぬエラー", Err.Number, Err.Description)
    PrepareSheets = False
    ' g_errorLogWorksheet might not be set, so can't use WriteErrorLog here.
End Function

Public Sub PrepareOutputSheet(ByRef config As tConfigSettings, ByVal mainWorkbook As Workbook, ByRef outOutputStartRow As Long)
    ' 出力シートを準備します。必要に応じて既存データをクリアし、データの書き込み開始行を設定します。
    ' Arguments:
    '   config: (I) tConfigSettings型。設定情報を保持します。
    '   mainWorkbook: (I) Workbook型。マクロ本体のワークブックオブジェクト。
    '   outOutputStartRow: (O) Long型。出力シートへの書き込み開始行番号。

    Dim wsOutput As Worksheet
    Dim headerActualRowCount As Long

    On Error GoTo PrepareOutputSheet_Error
    outOutputStartRow = 1 ' Default if something fails

    If mainWorkbook Is Nothing Then ' Simplified check, config cannot be Nothing as UDT
         Call M04_LogWriter.SafeWriteErrorLog("ERROR", mainWorkbook, "ErrorLog_M03_Fallback", "M03_SheetManager", "PrepareOutputSheet", "mainWorkbookがNothingです。", 0, "")
         Exit Sub
    End If
    If Len(config.OutputSheetName) = 0 Then
         Call M04_LogWriter.SafeWriteErrorLog("ERROR", mainWorkbook, config.ErrorLogSheetName, "M03_SheetManager", "PrepareOutputSheet", "config.OutputSheetNameが空です。", 0, "")
         Exit Sub
    End If

    On Error Resume Next
    Set wsOutput = mainWorkbook.Worksheets(config.OutputSheetName)
    On Error GoTo PrepareOutputSheet_Error

    If wsOutput Is Nothing Then
        ' EnsureSheetExists (called by PrepareSheets) should have created it. If not, something is wrong.
        Call M04_LogWriter.SafeWriteErrorLog("ERROR", mainWorkbook, config.ErrorLogSheetName, "M03_SheetManager", "PrepareOutputSheet", "出力シート「" & config.OutputSheetName & "」が見つかりません。", 0, "")
        Exit Sub
    End If

    headerActualRowCount = GetHeaderRowCountForSheet(wsOutput, config, mainWorkbook) ' Added mainWorkbook
    If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE: M03_SheetManager.PrepareOutputSheet - Determined header rows: " & headerActualRowCount & " for sheet " & wsOutput.Name

    If UCase(config.OutputDataOption) = "リセット" Then
        If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE: M03_SheetManager.PrepareOutputSheet - Clearing data from row " & (headerActualRowCount + 1) & " onwards in sheet " & wsOutput.Name
        If wsOutput.ProtectContents Then
             Call M04_LogWriter.SafeWriteErrorLog("WARNING", mainWorkbook, config.ErrorLogSheetName, "M03_SheetManager", "PrepareOutputSheet", "シート「" & wsOutput.Name & "」が保護されているためデータクリアをスキップしました。", 0, "")
        ElseIf headerActualRowCount < wsOutput.Rows.Count Then ' Ensure there are rows below header to clear
            wsOutput.Rows(headerActualRowCount + 1 & ":" & wsOutput.Rows.Count).ClearContents
        End If
    End If

    outOutputStartRow = headerActualRowCount + 1
    If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE: M03_SheetManager.PrepareOutputSheet - Output start row set to: " & outOutputStartRow
    Exit Sub
PrepareOutputSheet_Error:
    Call M04_LogWriter.SafeWriteErrorLog("ERROR", mainWorkbook, config.ErrorLogSheetName, "M03_SheetManager", "PrepareOutputSheet", "実行時エラー " & Err.Number & ": " & Err.Description, Err.Number, Err.Description)
    outOutputStartRow = 1 ' Fallback
End Sub
