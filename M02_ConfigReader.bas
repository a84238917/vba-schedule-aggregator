' バージョン：v0.5.1
Option Explicit
' このモジュールは、設定シートから情報を読み取り、g_configSettings グローバル変数を設定する役割を担います。
' 主に LoadConfiguration 関数を通じて、M00_GlobalDeclarationsで定義された tConfigSettings 型の変数に値を設定します。

Private Const MODULE_NAME As String = "M02_ConfigReader"

' --- Public Functions ---

' 設定情報をConfigシートから読み込み、提供されたtConfigSettings構造体に格納します。
' @param configStruct 読み込まれた設定情報を格納するためのtConfigSettings型の変数（参照渡し）。
' @param mainWorkbook マクロが実行されているメインのワークブックオブジェクト。
' @param configSheetName 設定情報が記載されているシートの名前。
' @return Boolean 読み込みが成功した場合はTrue、失敗した場合はFalse。
Public Function LoadConfiguration(ByRef configStruct As tConfigSettings, ByVal mainWorkbook As Workbook, ByVal configSheetName As String) As Boolean
    Dim wsConfig As Worksheet
    Dim funcName As String
    funcName = "LoadConfiguration"

    ' Declarations for loop counters as per subtask instructions
    Dim fSectionConfigReadIndex As Long   ' For F-Section reading loop
    Dim gSectionHeaderReadIndex As Long   ' For G-Section OutputHeaderContents reading loop
    Dim dbgFSectionPrintIndex As Long     ' For Debug Print of F-Section arrays
    Dim dbgGSectionPrintIndex As Long     ' For Debug Print of G-Section OutputHeaderContents array

    ' Variables that might be used in F-Section reading if it were inline
    Dim itemName As String
    Dim offsetStr As String
    Dim tempOffset As tOffset
    Dim actualOffsetCount As Long ' Added based on common F-section logic
    Dim currentFatalErrorState As Boolean ' Added based on common F-section logic

    ' Variables for G-Section (if inline)
    Dim headerCellAddress As String
    Dim headerVal As String
    Dim rawHeaderCellVal As Variant
    Dim outputOpt As String


    On Error GoTo ErrorHandler

    ' 設定シートの存在確認と取得
    On Error Resume Next
    Set wsConfig = mainWorkbook.Sheets(configSheetName)
    On Error GoTo ErrorHandler ' Resume Next の影響範囲を最小限に

    If wsConfig Is Nothing Then
        ' M04_LogWriter.WriteErrorLog がまだ利用できない可能性を考慮し、呼び出し先の整備が前提
        ' この段階では MsgBox で十分かもしれないが、設計に合わせてM04を利用
        Call M04_LogWriter.WriteErrorLog("CRITICAL", MODULE_NAME, funcName, "設定シート「" & configSheetName & "」が見つかりません。処理を続行できません。", 0, "File Not Found")
        LoadConfiguration = False
        Exit Function
    End If

    ' 各セクションの読み込み
    Call LoadGeneralSettings(configStruct, wsConfig)                 ' A. 一般設定
    Call LoadScheduleFileSettings(configStruct, wsConfig)           ' B. 工程表ファイル設定
    Call LoadProcessPatternDefinition(configStruct, wsConfig)       ' C. 工程パターン定義
    Call LoadFilterConditions(configStruct, wsConfig)               ' D. フィルタ条件
    Call LoadTargetFileDefinition(configStruct, wsConfig)           ' E. 処理対象ファイル定義
    Call LoadExtractionDataOffsetDefinition(configStruct, wsConfig) ' F. 抽出データオフセット定義
    Call LoadOutputSheetSettings(configStruct, wsConfig)            ' G. 出力シート設定

    ' 追加メンバーの設定
    configStruct.ConfigSheetFullName = mainWorkbook.FullName & "\" & wsConfig.Name ' wsConfig.Parent.FullName は mainWorkbook.FullName と同じはず
    ' configStruct.StartTime は MainControl で設定済み
    ' configStruct.ScriptFullName は MainControl で設定済み
    ' configStruct.MainWorkbookObject は MainControl で設定済み

    LoadConfiguration = True
    Exit Function

ErrorHandler:
    ' 同上、M04_LogWriter の利用は整備が前提
    Call M04_LogWriter.WriteErrorLog("CRITICAL", MODULE_NAME, funcName, "設定読み込み中に予期せぬエラーが発生しました。", Err.Number, Err.Description)
    LoadConfiguration = False
End Function


' --- Private Helper Subroutines ---

' A. 一般設定 (O列)
Private Sub LoadGeneralSettings(ByRef config As tConfigSettings, ByVal ws As Worksheet)
    Dim funcName As String: funcName = "LoadGeneralSettings"
    On Error Resume Next ' 特定のセルアクセスエラーをハンドルするため

    config.DebugModeFlag = ReadBoolCell(ws, "O3", MODULE_NAME, funcName, "デバッグモードフラグ")
    config.DefaultFolderPath = ReadStringCell(ws, "O12", MODULE_NAME, funcName, "デフォルトフォルダパス")
    config.OutputSheetName = ReadStringCell(ws, "O43", MODULE_NAME, funcName, "抽出結果出力シート名", "抽出結果")
    config.SearchConditionLogSheetName = ReadStringCell(ws, "O44", MODULE_NAME, funcName, "検索条件ログシート名", "検索条件ログ")
    config.ErrorLogSheetName = ReadStringCell(ws, "O45", MODULE_NAME, funcName, "エラーログシート名", "エラーログ")
    config.ConfigSheetName = ReadStringCell(ws, "O46", MODULE_NAME, funcName, "設定ファイルシート名", CONFIG_SHEET_DEFAULT_NAME)
    config.GetPatternDataMethod = ReadBoolCell(ws, "O122", MODULE_NAME, funcName, "工程パターンデータ取得方法")

    If Err.Number <> 0 Then Call M04_LogWriter.WriteErrorLog("WARNING", MODULE_NAME, funcName, "一般設定の読み込み中にエラーが発生しました。", Err.Number, Err.Description)
    On Error GoTo 0
End Sub

' B. 工程表ファイル設定 (O列)
Private Sub LoadScheduleFileSettings(ByRef config As tConfigSettings, ByVal ws As Worksheet)
    Dim funcName As String: funcName = "LoadScheduleFileSettings"
    On Error Resume Next

    config.TargetSheetNames = ReadRangeToArray(ws, "O66:O75", MODULE_NAME, funcName, "工程表内 検索対象シート名リスト")
    config.HeaderRowCount = ReadLongCell(ws, "O87", MODULE_NAME, funcName, "工程表ヘッダー行数")
    config.HeaderColCount = ReadLongCell(ws, "O88", MODULE_NAME, funcName, "工程表ヘッダー列数")
    config.RowsPerDay = ReadLongCell(ws, "O89", MODULE_NAME, funcName, "1日のデータが占める行数")
    config.MaxDaysPerSheet = ReadLongCell(ws, "O90", MODULE_NAME, funcName, "1シート内の最大日数")
    config.YearCellAddress = ReadStringCell(ws, "O101", MODULE_NAME, funcName, "「年」のセルアドレス")
    config.MonthCellAddress = ReadStringCell(ws, "O102", MODULE_NAME, funcName, "「月」のセルアドレス")
    config.DayColumnLetter = ReadStringCell(ws, "O103", MODULE_NAME, funcName, "「日」の値がある列文字")
    config.DayRowOffset = ReadLongCell(ws, "O104", MODULE_NAME, funcName, "「日」の値の行オフセット")
    config.ProcessesPerDay = ReadLongCell(ws, "O114", MODULE_NAME, funcName, "1日の工程数")

    If Err.Number <> 0 Then Call M04_LogWriter.WriteErrorLog("WARNING", MODULE_NAME, funcName, "工程表ファイル設定の読み込み中にエラー。", Err.Number, Err.Description)
    On Error GoTo 0
End Sub

' C. 工程パターン定義 (I,J,K,L,M,N列, O-X列)
Private Sub LoadProcessPatternDefinition(ByRef config As tConfigSettings, ByVal ws As Worksheet)
    Dim funcName As String: funcName = "LoadProcessPatternDefinition"
    Dim procPtn_i As Long ' Renamed from i
    Dim numProcesses As Long
    On Error Resume Next

    config.CurrentPatternIdentifier = ReadStringCell(ws, "O126", MODULE_NAME, funcName, "現在処理中ファイル適用工程パターン識別子")
    numProcesses = config.ProcessesPerDay
    If numProcesses <= 0 Then numProcesses = 10 ' デフォルト値、またはエラーログ

    config.ProcessKeys = ReadRangeToArray(ws, "I129:I" & (128 + numProcesses), MODULE_NAME, funcName, "工程キーリスト")
    config.Kankatsu1List = ReadRangeToArray(ws, "J129:J" & (128 + numProcesses), MODULE_NAME, funcName, "管内1リスト")
    config.Kankatsu2List = ReadRangeToArray(ws, "K129:K" & (128 + numProcesses), MODULE_NAME, funcName, "管内2リスト")
    config.Bunrui1List = ReadRangeToArray(ws, "L129:L" & (128 + numProcesses), MODULE_NAME, funcName, "分類1リスト")
    config.Bunrui2List = ReadRangeToArray(ws, "M129:M" & (128 + numProcesses), MODULE_NAME, funcName, "分類2リスト")
    config.Bunrui3List = ReadRangeToArray(ws, "N129:N" & (128 + numProcesses), MODULE_NAME, funcName, "分類3リスト")

    Dim headerData As Variant
    headerData = ws.Range("O128:X128").Value
    If IsArray(headerData) Then
        ReDim config.ProcessColCountSheetHeaders(1 To UBound(headerData, 2))
        For procPtn_i = 1 To UBound(headerData, 2) ' Renamed i to procPtn_i
            config.ProcessColCountSheetHeaders(procPtn_i) = Trim(CStr(headerData(1, procPtn_i))) ' Renamed i to procPtn_i
        Next procPtn_i
    End If

    config.ProcessColCounts = ws.Range("O129:X" & (128 + numProcesses)).Value

    If General_IsArrayInitialized(config.Kankatsu1List) And General_IsArrayInitialized(config.Kankatsu2List) Then
        Dim k1Count As Long, k2Count As Long, maxCount As Long
        On Error Resume Next ' UBoundでエラーになる場合（配列が初期化されていないなど）
        k1Count = UBound(config.Kankatsu1List)
        k2Count = UBound(config.Kankatsu2List)
        If Err.Number <> 0 Then Err.Clear Else maxCount = IIf(k1Count > k2Count, k1Count, k2Count)
        On Error GoTo 0

        If maxCount > 0 Then
            ReDim config.ProcessDetails(1 To maxCount)
            For procPtn_i = 1 To maxCount ' Renamed i to procPtn_i
                If procPtn_i <= k1Count Then config.ProcessDetails(procPtn_i).Kankatsu1 = config.Kankatsu1List(procPtn_i)
                If procPtn_i <= k2Count Then config.ProcessDetails(procPtn_i).Kankatsu2 = config.Kankatsu2List(procPtn_i)
            Next procPtn_i
        End If
    End If

    If Err.Number <> 0 Then Call M04_LogWriter.WriteErrorLog("WARNING", MODULE_NAME, funcName, "工程パターン定義の読み込み中にエラー。", Err.Number, Err.Description)
    On Error GoTo 0
End Sub

' D. フィルタ条件 (O列)
Private Sub LoadFilterConditions(ByRef config As tConfigSettings, ByVal ws As Worksheet)
    Dim funcName As String: funcName = "LoadFilterConditions"
    On Error Resume Next

    config.WorkerFilterLogic = ReadStringCell(ws, "O242", MODULE_NAME, funcName, "作業員フィルター検索論理", "AND")
    config.WorkerFilterList = ReadRangeToArray(ws, "O243:O262", MODULE_NAME, funcName, "作業員フィルターリスト")
    config.Kankatsu1FilterList = ReadRangeToArray(ws, "O275:O294", MODULE_NAME, funcName, "管内1フィルターリスト")
    config.Kankatsu2FilterList = ReadRangeToArray(ws, "O305:O334", MODULE_NAME, funcName, "管内2フィルターリスト")
    config.Bunrui1Filter = ReadStringCell(ws, "O346", MODULE_NAME, funcName, "分類1フィルター")
    config.Bunrui2Filter = ReadStringCell(ws, "O367", MODULE_NAME, funcName, "分類2フィルター")
    config.Bunrui3Filter = ReadStringCell(ws, "O388", MODULE_NAME, funcName, "分類3フィルター")
    config.KoujiShuruiFilterList = ReadRangeToArray(ws, "O409:O418", MODULE_NAME, funcName, "工事種類フィルターリスト")
    config.KoubanFilterList = ReadRangeToArray(ws, "O431:O440", MODULE_NAME, funcName, "工番フィルターリスト")
    config.SagyoushuruiFilterList = ReadRangeToArray(ws, "O451:O470", MODULE_NAME, funcName, "作業種類フィルターリスト")
    config.TantouFilterList = ReadRangeToArray(ws, "O481:O490", MODULE_NAME, funcName, "担当の名前フィルターリスト")
    config.NinzuFilter = ReadStringCell(ws, "O503", MODULE_NAME, funcName, "人数フィルター")
    config.IsNinzuFilterOriginallyEmpty = (Trim(config.NinzuFilter) = "")
    config.SagyouKashoKindFilter = ReadStringCell(ws, "O514", MODULE_NAME, funcName, "作業箇所の種類フィルター")
    config.SagyouKashoFilterList = ReadRangeToArray(ws, "O525:O544", MODULE_NAME, funcName, "作業箇所フィルターリスト")

    If Err.Number <> 0 Then Call M04_LogWriter.WriteErrorLog("WARNING", MODULE_NAME, funcName, "フィルタ条件の読み込み中にエラー。", Err.Number, Err.Description)
    On Error GoTo 0
End Sub

' E. 処理対象ファイル定義 (P, Q列)
Private Sub LoadTargetFileDefinition(ByRef config As tConfigSettings, ByVal ws As Worksheet)
    Dim funcName As String: funcName = "LoadTargetFileDefinition"
    On Error Resume Next

    config.TargetFileFolderPaths = ReadRangeToArray(ws, "P557:P756", MODULE_NAME, funcName, "処理対象ファイル/フォルダパスリスト")
    config.FilePatternIdentifiers = ReadRangeToArray(ws, "Q557:Q756", MODULE_NAME, funcName, "各処理対象ファイル適用工程パターン識別子")

    If Err.Number <> 0 Then Call M04_LogWriter.WriteErrorLog("WARNING", MODULE_NAME, funcName, "処理対象ファイル定義の読み込み中にエラー。", Err.Number, Err.Description)
    On Error GoTo 0
End Sub

' F. 抽出データオフセット定義 (N, O列)
Private Sub LoadExtractionDataOffsetDefinition(ByRef config As tConfigSettings, ByVal ws As Worksheet)
    Dim funcName As String: funcName = "LoadExtractionDataOffsetDefinition"
    Dim offsetDef_i As Long, actualCount As Long ' Renamed i to offsetDef_i
    Dim itemsRaw As Variant, offsetStringsRaw As Variant
    Dim parsedItems() As String, parsedOffsets() As tOffset, parsedRawStrings() As String
    Dim tempOffset As tOffset
    On Error Resume Next

    itemsRaw = ws.Range("N778:N792").Value
    offsetStringsRaw = ws.Range("O778:O792").Value
    If Err.Number <> 0 Then
        Call M04_LogWriter.WriteErrorLog("ERROR", MODULE_NAME, funcName, "オフセット定義範囲(N778:O792)の読み込みに失敗しました。", Err.Number, Err.Description)
        Exit Sub
    End If
    On Error GoTo 0 ' Resume Next 解除

    actualCount = 0
    If IsArray(itemsRaw) And IsArray(offsetStringsRaw) Then
        For offsetDef_i = 1 To UBound(itemsRaw, 1) ' Renamed i to offsetDef_i
            If Trim(CStr(itemsRaw(offsetDef_i, 1))) <> "" And Not IsEmpty(itemsRaw(offsetDef_i,1)) Then '項目名があればカウント ' Renamed i to offsetDef_i
                actualCount = actualCount + 1
            End If
        Next offsetDef_i

        If actualCount > 0 Then
            ReDim parsedItems(1 To actualCount)
            ReDim parsedOffsets(1 To actualCount)
            ReDim parsedRawStrings(1 To actualCount)

            Dim offsetDef_currentIdx As Long: offsetDef_currentIdx = 0 ' Renamed currentIdx
            For offsetDef_i = 1 To UBound(itemsRaw, 1) ' Renamed i to offsetDef_i
                If Trim(CStr(itemsRaw(offsetDef_i, 1))) <> "" And Not IsEmpty(itemsRaw(offsetDef_i,1)) Then ' Renamed i to offsetDef_i
                    offsetDef_currentIdx = offsetDef_currentIdx + 1 ' Renamed currentIdx
                    parsedItems(offsetDef_currentIdx) = Trim(CStr(itemsRaw(offsetDef_i, 1))) ' Renamed currentIdx, i
                    parsedRawStrings(offsetDef_currentIdx) = Trim(CStr(offsetStringsRaw(offsetDef_i, 1))) ' Renamed currentIdx, i

                    tempOffset = GetSpecificOffsetFromString(parsedRawStrings(offsetDef_currentIdx), parsedItems(offsetDef_currentIdx), funcName) ' Renamed currentIdx
                    parsedOffsets(offsetDef_currentIdx) = tempOffset ' Renamed currentIdx

                    ' 特定の項目に対するIs...OriginallyEmptyフラグの設定
                    Select Case parsedItems(offsetDef_currentIdx) ' Renamed currentIdx
                        Case "工番": config.IsOffsetKoubanOriginallyEmpty = (parsedRawStrings(offsetDef_currentIdx) = "")
                        Case "変電所": config.IsOffsetHensendenjoOriginallyEmpty = (parsedRawStrings(offsetDef_currentIdx) = "")
                        Case "作業名1": config.IsOffsetSagyomei1OriginallyEmpty = (parsedRawStrings(offsetDef_currentIdx) = "")
                        Case "作業名2": config.IsOffsetSagyomei2OriginallyEmpty = (parsedRawStrings(offsetDef_currentIdx) = "")
                        Case "担当の名前": config.IsOffsetTantouOriginallyEmpty = (parsedRawStrings(offsetDef_currentIdx) = "")
                        Case "工事種類": config.IsOffsetKoujiShuruiOriginallyEmpty = (parsedRawStrings(offsetDef_currentIdx) = "")
                        Case "人数": config.IsOffsetNinzuOriginallyEmpty = (parsedRawStrings(offsetDef_currentIdx) = "")
                        Case "作業員": config.IsOffsetSagyoinOriginallyEmpty = (parsedRawStrings(offsetDef_currentIdx) = "")
                        Case "旧その他": config.IsOffsetSonotaOriginallyEmpty = (parsedRawStrings(offsetDef_currentIdx) = "")
                        Case "終了時間": config.IsOffsetShuuryoJikanOriginallyEmpty = (parsedRawStrings(offsetDef_currentIdx) = "")
                        Case "分類1抽出元": config.IsOffsetBunrui1ExtSrcOriginallyEmpty = (parsedRawStrings(offsetDef_currentIdx) = "")
                    End Select
                End If
            Next offsetDef_i ' Renamed i to offsetDef_i
            config.OffsetItemNames = parsedItems
            config.Offsets = parsedOffsets
            config.OffsetValuesRaw = parsedRawStrings ' 生の文字列も保存
        Else
            Call M04_LogWriter.WriteErrorLog("INFORMATION", MODULE_NAME, funcName, "有効なオフセット定義項目名(N列)が見つかりませんでした。")
        End If
    Else
         Call M04_LogWriter.WriteErrorLog("WARNING", MODULE_NAME, funcName, "オフセット項目名または値の範囲が配列として読み込めませんでした。")
    End If

    If Err.Number <> 0 Then Call M04_LogWriter.WriteErrorLog("WARNING", MODULE_NAME, funcName, "抽出データオフセット定義の読み込み中にエラー。", Err.Number, Err.Description)
    On Error GoTo 0
End Sub

' G. 出力シート設定 (O列)
Private Sub LoadOutputSheetSettings(ByRef config As tConfigSettings, ByVal ws As Worksheet)
    Dim funcName As String: funcName = "LoadOutputSheetSettings"
    On Error Resume Next

    config.OutputHeaderRowCount = ReadLongCell(ws, "O811", MODULE_NAME, funcName, "出力シートヘッダー行数", 1)
    config.OutputHeaderContents = ReadRangeToArray(ws, "O812:O821", MODULE_NAME, funcName, "出力シートヘッダー内容")
    config.OutputDataOption = ReadStringCell(ws, "O1124", MODULE_NAME, funcName, "出力データオプション", "上書き")
    config.HideSheetMethod = ReadStringCell(ws, "O1126", MODULE_NAME, funcName, "非表示方式", "非表示")
    config.HideSheetNames = ReadRangeToArray(ws, "O1127:O1146", MODULE_NAME, funcName, "マクロ実行後非表示シートリスト")

    If Err.Number <> 0 Then Call M04_LogWriter.WriteErrorLog("WARNING", MODULE_NAME, funcName, "出力シート設定の読み込み中にエラー。", Err.Number, Err.Description)
    On Error GoTo 0
End Sub


' --- Private Helper Functions for Reading Cell Values ---

Private Function ReadStringCell(ws As Worksheet, addr As String, moduleN As String, funcN As String, itemName As String, Optional defaultValue As String = vbNullString) As String
    Dim val As Variant
    On Error Resume Next
    val = ws.Range(addr).Value
    If Err.Number <> 0 Then
        ReadStringCell = defaultValue
        Call M04_LogWriter.WriteErrorLog("WARNING", moduleN, funcN, itemName & " (" & addr & ") 読み取り失敗。デフォルト「" & defaultValue & "」使用。", Err.Number, Err.Description)
    Else
        If IsEmpty(val) Or Trim(CStr(val)) = "" Then
            ReadStringCell = defaultValue
        Else
            ReadStringCell = Trim(CStr(val))
        End If
    End If
    On Error GoTo 0
End Function

Private Function ReadLongCell(ws As Worksheet, addr As String, moduleN As String, funcN As String, itemName As String, Optional defaultValue As Long = 0) As Long
    Dim val As Variant
    On Error Resume Next
    val = ws.Range(addr).Value
    If Err.Number <> 0 Then
        ReadLongCell = defaultValue
        Call M04_LogWriter.WriteErrorLog("WARNING", moduleN, funcN, itemName & " (" & addr & ") 読み取り失敗。デフォルト「" & defaultValue & "」使用。", Err.Number, Err.Description)
    Else
        If IsEmpty(val) Or Not IsNumeric(val) Then
            ReadLongCell = defaultValue
            If Not IsEmpty(val) Then Call M04_LogWriter.WriteErrorLog("WARNING", moduleN, funcN, itemName & " (" & addr & ") が数値でない。デフォルト「" & defaultValue & "」使用。")
        Else
            ReadLongCell = CLng(val)
        End If
    End If
    On Error GoTo 0
End Function

Private Function ReadBoolCell(ws As Worksheet, addr As String, moduleN As String, funcN As String, itemName As String, Optional defaultValue As Boolean = False) As Boolean
    Dim val As Variant
    On Error Resume Next
    val = ws.Range(addr).Value
    If Err.Number <> 0 Then
        ReadBoolCell = defaultValue
        Call M04_LogWriter.WriteErrorLog("WARNING", moduleN, funcN, itemName & " (" & addr & ") 読み取り失敗。デフォルト「" & defaultValue & "」使用。", Err.Number, Err.Description)
    Else
        If IsEmpty(val) Then
            ReadBoolCell = defaultValue
        Else
            ReadBoolCell = (UCase(Trim(CStr(val))) = "TRUE")
        End If
    End If
    On Error GoTo 0
End Function

Private Function ReadRangeToArray(ws As Worksheet, rangeAddress As String, moduleN As String, funcN As String, itemName As String) As String()
    Dim data As Variant, result() As String, arrRead_i As Long, nonEmptyCount As Long ' Renamed i to arrRead_i
    On Error Resume Next
    data = ws.Range(rangeAddress).Value
    If Err.Number <> 0 Then
        Call M04_LogWriter.WriteErrorLog("WARNING", moduleN, funcN, itemName & " (" & rangeAddress & ") 範囲読み取り失敗。", Err.Number, Err.Description)
        Exit Function ' Returns uninitialized array
    End If
    On Error GoTo 0

    If IsArray(data) Then
        ReDim result(1 To UBound(data, 1))
        For arrRead_i = 1 To UBound(data, 1) ' Renamed i to arrRead_i
            If Not IsEmpty(data(arrRead_i, 1)) And Trim(CStr(data(arrRead_i, 1))) <> "" Then ' Renamed i to arrRead_i
                result(arrRead_i) = Trim(CStr(data(arrRead_i, 1))) ' Renamed i to arrRead_i
                nonEmptyCount = nonEmptyCount + 1
            Else
                result(arrRead_i) = vbNullString ' Renamed i to arrRead_i
            End If
        Next arrRead_i
        If nonEmptyCount = 0 Then Erase result ' No valid data, return uninitialized
    Else ' Single cell
        If Not IsEmpty(data) And Trim(CStr(data)) <> "" Then
            ReDim result(1 To 1): result(1) = Trim(CStr(data))
        End If ' Else returns uninitialized
    End If
    ReadRangeToArray = result
End Function

' 文字列からオフセットを解析。エラー時は(0,0)を返しログ記録。
Private Function GetSpecificOffsetFromString(val As String, itemName As String, funcN As String) As tOffset
    Dim parts() As String
    Dim tempOffset As tOffset
    tempOffset.Row = 0: tempOffset.Col = 0 ' Initialize

    If val <> vbNullString Then
        parts = Split(val, ",")
        If UBound(parts) = 1 Then
            If IsNumeric(Trim(parts(0))) And IsNumeric(Trim(parts(1))) Then
                tempOffset.Row = CLng(Trim(parts(0)))
                tempOffset.Col = CLng(Trim(parts(1)))
            Else
                Call M04_LogWriter.WriteErrorLog("WARNING", MODULE_NAME, funcN, itemName & " のオフセット値「" & val & "」に数値でない部分が含まれます。デフォルト(0,0)使用。")
            End If
        Else
            Call M04_LogWriter.WriteErrorLog("WARNING", MODULE_NAME, funcN, itemName & " のオフセット値「" & val & "」の形式が不正。デフォルト(0,0)使用。")
        End If
    End If
    GetSpecificOffsetFromString = tempOffset
End Function

Public Function General_IsArrayInitialized(arr As Variant) As Boolean
    If Not IsArray(arr) Then Exit Function
    On Error Resume Next
    Dim lBoundCheck As Long: lBoundCheck = LBound(arr)
    If Err.Number = 0 Then General_IsArrayInitialized = True
    On Error GoTo 0
End Function
