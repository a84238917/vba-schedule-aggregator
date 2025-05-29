Option Explicit
' このモジュールは、マクロが使用する各種ワークシート（出力シート、ログシートなど）の準備、検証、および管理を担当します。

Private Function IsArrayInitialized(arr As Variant) As Boolean
    ' 配列が有効に初期化されているか（少なくとも1つの要素を持つか）を確認します。
    ' Variant型が配列でない場合、または配列であっても要素が割り当てられていない場合（Dim arr() のみでReDimされていない状態など）はFalseを返します。
    On Error GoTo NotAnArrayOrNotInitialized
    If IsArray(arr) Then
        Dim lBoundCheck As Long
        lBoundCheck = LBound(arr) ' 配列がReDimされていれば、LBoundはエラーにならない (空でも ReDim arr(0 To -1) など)
        IsArrayInitialized = True ' LBoundがエラーを起こさなければ、配列は有効（空でもReDimされていればOK）
        Exit Function
    End If
NotAnArrayOrNotInitialized:
    IsArrayInitialized = False
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

    Dim ws As Worksheet

    On Error Resume Next ' シートの存在確認に関するエラーを一旦無視
    Set ws = targetWorkbook.Sheets(sheetNameToEnsure)
    On Error GoTo EnsureSheetExists_Error ' 通常のエラーハンドリングに戻す

    If ws Is Nothing Then ' シートが存在しない場合
        If DEBUG_MODE_DETAIL Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M03_SheetManager.EnsureSheetExists - Sheet '" & sheetNameToEnsure & "' not found. Creating new sheet."
        On Error GoTo CreateSheet_Error ' シート作成に特化したエラーハンドリング
        Set ws = targetWorkbook.Sheets.Add(After:=targetWorkbook.Sheets(targetWorkbook.Sheets.Count))
        ws.Name = sheetNameToEnsure
        On Error GoTo EnsureSheetExists_Error ' 通常のエラーハンドリングに戻す

        If createHeaders Then
            If DEBUG_MODE_DETAIL Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M03_SheetManager.EnsureSheetExists - Creating headers for new sheet: " & sheetNameToEnsure
            If sheetNameToEnsure = config.ErrorLogSheetName Then
                ws.Cells(1, 1).Value = "発生日時"
                ws.Cells(1, 2).Value = "モジュール"
                ws.Cells(1, 3).Value = "プロシージャ"
                ws.Cells(1, 4).Value = "関連情報"
                ws.Cells(1, 5).Value = "エラー番号"
                ws.Cells(1, 6).Value = "エラー内容"
                ws.Cells(1, 7).Value = "対処内容"
                ws.Cells(1, 8).Value = "変数情報"
            ElseIf sheetNameToEnsure = config.SearchConditionLogSheetName Then
                ws.Cells(1, 1).Value = "実行日時"
                ws.Cells(1, 2).Value = "フィルター項目"
                ws.Cells(1, 3).Value = "条件"
                ws.Cells(1, 4).Value = "備考"
            ElseIf sheetNameToEnsure = config.OutputSheetName Then
                ' G-2. 出力シートヘッダー内容 (config.OutputHeaderContents) に基づくヘッダー作成ロジックはステップ5で実装
                ' For now, can leave a placeholder comment or a single dummy header for testing if needed.
                ' ws.Cells(1, 1).Value = "（出力シートヘッダー仮）"
                If DEBUG_MODE_DETAIL Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M03_SheetManager.EnsureSheetExists - Placeholder for OutputSheetName headers. Full implementation in Step 5."
            End If
        End If
    Else
        If DEBUG_MODE_DETAIL Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M03_SheetManager.EnsureSheetExists - Sheet '" & sheetNameToEnsure & "' already exists."
    End If

    Set EnsureSheetExists = ws
    Exit Function

CreateSheet_Error:
    If DEBUG_MODE_ERROR Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - ERROR: M03_SheetManager.EnsureSheetExists (CreateSheet_Error) - Error " & Err.Number & ": " & Err.Description & " while trying to create/name sheet '" & sheetNameToEnsure & "'"
    Call SafeWriteErrorLog(targetWorkbook, config.ErrorLogSheetName, "M03_SheetManager", "EnsureSheetExists (Create)", "シート '" & sheetNameToEnsure & "' の作成または命名に失敗しました。", Err.Number, Err.Description)
    Set EnsureSheetExists = Nothing
    Exit Function

EnsureSheetExists_Error:
    If DEBUG_MODE_ERROR Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - ERROR: M03_SheetManager.EnsureSheetExists - Error " & Err.Number & ": " & Err.Description & " (Caller: " & callerFuncName & ")"
    Call SafeWriteErrorLog(targetWorkbook, config.ErrorLogSheetName, "M03_SheetManager", "EnsureSheetExists", "シート '" & sheetNameToEnsure & "' の確認/作成中に予期せぬエラー (呼び出し元: " & callerFuncName & ")", Err.Number, Err.Description)
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

    If DEBUG_MODE_DETAIL Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M03_SheetManager.PrepareSheets - Starting sheet preparation."

    ' エラーログシートの準備
    Set wsErr = EnsureSheetExists(targetWorkbook, config.ErrorLogSheetName, config, "PrepareSheets", True)
    If wsErr Is Nothing Then
        If DEBUG_MODE_ERROR Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - ERROR: M03_SheetManager.PrepareSheets - Failed to prepare Error Log Sheet."
        ' SafeWriteErrorLog is called by EnsureSheetExists on failure to create, or here if EnsureSheetExists returns Nothing for other reasons.
        ' However, the prompt explicitly asks for this call here too.
        Call SafeWriteErrorLog(targetWorkbook, config.ErrorLogSheetName, "M03_SheetManager", "PrepareSheets", "エラーログシートの準備に失敗しました。", 0, "EnsureSheetExistsがNothingを返しました")
        Exit Function ' Returns False
    Else
        Set g_errorLogWorksheet = wsErr
        ' 次のエラーログ書き込み行を決定
        If wsErr.Cells(1, 1).Value = vbNullString Then ' Check if sheet is completely empty or A1 is blank
            g_nextErrorLogRow = IIf(Application.WorksheetFunction.CountA(wsErr.Rows(1)) > 0, 2, 1) ' Start from row 2 if headers (any cell in row 1) exist, else 1
        Else ' A1 has data (likely headers or existing log)
            g_nextErrorLogRow = wsErr.Cells(wsErr.Rows.Count, 1).End(xlUp).Row + 1
        End If
        
        ' if first row has data (e.g. header), and g_nextErrorLogRow is 1 (e.g. from an empty sheet where End(xlUp) was row 1), start from second
        If g_nextErrorLogRow = 1 And wsErr.Cells(1,1).Value <> vbNullString Then g_nextErrorLogRow = 2

        If g_nextErrorLogRow < 1 Then g_nextErrorLogRow = 1 ' Ensure it's at least 1
        
        If DEBUG_MODE_DETAIL Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M03_SheetManager.PrepareSheets - Error Log Sheet ready. Next error log row: " & g_nextErrorLogRow
    End If

    ' 検索条件ログシートの準備
    Set wsFilter = EnsureSheetExists(targetWorkbook, config.SearchConditionLogSheetName, config, "PrepareSheets", True)
    If wsFilter Is Nothing Then
        If DEBUG_MODE_ERROR Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - ERROR: M03_SheetManager.PrepareSheets - Failed to prepare Search Condition Log Sheet."
        Call SafeWriteErrorLog(targetWorkbook, config.ErrorLogSheetName, "M03_SheetManager", "PrepareSheets", "検索条件ログシートの準備に失敗しました。", 0, "EnsureSheetExistsがNothingを返しました")
        Exit Function ' Returns False
    Else
        If DEBUG_MODE_DETAIL Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M03_SheetManager.PrepareSheets - Search Condition Log Sheet ready."
    End If

    PrepareSheets = True ' All successful
    If DEBUG_MODE_DETAIL Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M03_SheetManager.PrepareSheets - Sheet preparation completed successfully."
    Exit Function

PrepareSheets_Error:
    If DEBUG_MODE_ERROR Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - ERROR: M03_SheetManager.PrepareSheets - Error " & Err.Number & ": " & Err.Description
    Call SafeWriteErrorLog(targetWorkbook, config.ErrorLogSheetName, "M03_SheetManager", "PrepareSheets", "PrepareSheets内で予期せぬエラー", Err.Number, Err.Description)
    PrepareSheets = False
    ' g_errorLogWorksheet might not be set, so can't use WriteErrorLog here.
End Function
