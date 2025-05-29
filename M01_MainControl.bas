Option Explicit
' このモジュールは、マクロ全体の実行エントリーポイントを提供し、主要な処理フェーズの呼び出しフローを定義し、基本的な初期化処理と包括的なエラーハンドリングを行います。

' Main Procedure: ExtractDataMain() As Sub
Public Sub ExtractDataMain()
    ' マクロの実行開始点です。
    Dim wbThis As Workbook ' このマクロが記述されているワークブック
    Dim startTime As Double ' 処理開始時刻
    Dim endTime As Double   ' 処理終了時刻
    Dim errNum As Long, errDesc As String, errSource As String ' GlobalErrorHandler_M01 で使用

    On Error GoTo GlobalErrorHandler_M01
    Application.ScreenUpdating = False
    startTime = Timer
    Set wbThis = ThisWorkbook

    Call InitializeConfigStructure(g_configSettings) ' グローバル設定構造体を初期化
    g_configSettings.StartTime = Now()
    g_configSettings.ScriptFullName = wbThis.FullName

    If DEBUG_MODE_ERROR Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG: M01_MainControl.ExtractDataMain - マクロ実行開始。初期化処理完了。"

    ' --- 1. Configシート読み込みフェーズ ---
    ' Call M02_ConfigReader.LoadConfiguration(g_configSettings, wbThis.Worksheets(CONFIG_SHEET_DEFAULT_NAME))

    ' --- 2. 各種シート準備フェーズ ---
    ' M03_SheetManager.PrepareSheets はグローバル変数 g_errorLogWorksheet と g_nextErrorLogRow を設定します。
    If Not M03_SheetManager.PrepareSheets(g_configSettings, wbThis) Then
        ' SafeWriteErrorLogの呼び出しには、既に存在するg_configSettingsからシート名を取得する
        ' ただし、g_configSettings自体が未初期化の可能性は極めて低い（InitializeConfigStructureが先のため）
        ' ここではconfigが最低限読み込めている前提でエラーログシート名をg_configSettingsから取得
        Dim tempErrorLogName As String
        If Not g_configSettings.ErrorLogSheetName = "" Then
            tempErrorLogName = g_configSettings.ErrorLogSheetName
        Else
            tempErrorLogName = "エラーログ(デフォルト)" ' Fallback name
        End If
        Call SafeWriteErrorLog(wbThis, tempErrorLogName, "M01_MainControl", "ExtractDataMain", "ログシート準備失敗", 0, "M03_SheetManager.PrepareSheetsがFalseを返しました")
        MsgBox "ログシートの準備に失敗しました。処理を中断します。", vbCritical, "初期化エラー"
        GoTo FinalizeMacro_M01
    End If

    ' --- 3. 処理対象ファイル特定フェーズ ---
    ' Call M05_FileProcessor.GetTargetFiles(g_configSettings)

    ' --- 4. 出力/ログ準備フェーズ ---
    ' Call M03_SheetManager.PrepareOutputSheet(wbThis, g_configSettings)

    ' --- 5. 検索条件ログ出力フェーズ ---
    ' ログシートが正常に準備された後に、検索条件ログを書き込みます。
    Call M04_LogWriter.WriteFilterLog(g_configSettings, wbThis)

    ' --- 6. メインループフェーズ (ファイルごとのデータ抽出処理) ---
    ' Dim targetFile As Variant
    ' For Each targetFile In g_configSettings.TargetFileFolderPaths
    '     If LogMain_IsArrayInitialized(g_configSettings.TargetFileFolderPaths) Then '念のため実行前に確認
    '         Call M06_DataExtractor.ExtractDataFromFile(CStr(targetFile), g_configSettings, wbThis.Worksheets(g_configSettings.OutputSheetName))
    '     End If
    ' Next targetFile

FinalizeMacro_M01:
    On Error Resume Next ' 終了処理中のエラーは無視
    Application.ScreenUpdating = True
    endTime = Timer
    ' MsgBox "処理完了 (仮) 処理時間: " & Format(endTime - startTime, "0.00") & "秒"
    Set wbThis = Nothing
    If DEBUG_MODE_ERROR Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG: M01_MainControl.ExtractDataMain - マクロ実行正常終了。処理時間: " & Format(endTime - startTime, "0.00") & "秒"
    Exit Sub

GlobalErrorHandler_M01:
    errNum = Err.Number
    errDesc = Err.Description
    errSource = Err.Source
    If DEBUG_MODE_ERROR Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - ERROR: M01_MainControl.ExtractDataMain (GlobalErrorHandler_M01) - Error " & errNum & ": " & errDesc & " (Source: " & errSource & ")"
    
    ' エラー情報をログに記録
    If g_errorLogWorksheet Is Nothing Then
        ' g_errorLogWorksheetが未設定の場合 (PrepareSheetsより前、または失敗時) はSafeWriteErrorLogを試みる
        Dim errorSheetNameAttempt As String
        ' g_configSettings は InitializeConfigStructure で初期化されているため Nothing にはならない。
        ' ErrorLogSheetName が空文字列の可能性はある。
        If Not g_configSettings.ErrorLogSheetName = "" Then
            errorSheetNameAttempt = g_configSettings.ErrorLogSheetName
        Else
            errorSheetNameAttempt = "エラーログ(M01エラーハンドラ)" ' より明確なフォールバック名
        End If
        Call SafeWriteErrorLog(wbThis, errorSheetNameAttempt, "M01_MainControl", "ExtractDataMain (GlobalErrorHandler_M01)", "エラー発生 (エラーログシート準備前または失敗): " & errSource, errNum, errDesc)
    Else
        ' g_errorLogWorksheetが設定されていれば通常のWriteErrorLogを使用
        Call WriteErrorLog("M01_MainControl", "ExtractDataMain (GlobalErrorHandler_M01)", errSource, errNum, errDesc, "処理中断")
    End If
    
    MsgBox "エラーが発生しました。" & vbCrLf & _
           "エラー番号: " & errNum & vbCrLf & _
           "内容: " & errDesc & vbCrLf & _
           "発生元: " & errSource & vbCrLf & _
           "処理を中断します。", vbCritical, "実行時エラー"
    Resume FinalizeMacro_M01
End Sub

' Helper Procedure: InitializeConfigStructure
Private Sub InitializeConfigStructure(ByRef configStruct As tConfigSettings)
    ' 引数で受け取ったtConfigSettings型の構造体の全メンバー（特に動的配列）を初期化（Erase）します。
    If DEBUG_MODE_DETAIL Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M01_MainControl.InitializeConfigStructure - 初期化開始"

    Erase configStruct.TargetSheetNames
    Erase configStruct.ProcessKeys
    Erase configStruct.Kankatsu1List
    Erase configStruct.Kankatsu2List
    Erase configStruct.Bunrui1List
    Erase configStruct.Bunrui2List
    Erase configStruct.Bunrui3List
    Erase configStruct.ProcessColCountSheetHeaders
    Erase configStruct.ProcessColCounts
    Erase configStruct.ProcessDetails
    Erase configStruct.ProcessPatternColNumbers
    Erase configStruct.WorkerFilterList
    Erase configStruct.Kankatsu1FilterList
    Erase configStruct.Kankatsu2FilterList
    Erase configStruct.KoujiShuruiFilterList
    Erase configStruct.KoubanFilterList
    Erase configStruct.SagyoushuruiFilterList
    Erase configStruct.TantouFilterList
    Erase configStruct.SagyouKashoFilterList
    Erase configStruct.TargetFileFolderPaths
    Erase configStruct.FilePatternIdentifiers
    Erase configStruct.OffsetItemNames
    Erase configStruct.OffsetValuesRaw
    Erase configStruct.Offsets
    Erase configStruct.OutputHeaderContents
    Erase configStruct.HideSheetNames

    If DEBUG_MODE_DETAIL Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M01_MainControl.InitializeConfigStructure - 初期化完了"
End Sub

' Helper Function: LogMain_IsArrayInitialized
Private Function LogMain_IsArrayInitialized(arr As Variant) As Boolean
    ' 配列が初期化されているか（少なくとも1つの要素を持つか）を確認します。
    ' Variant型が配列でない場合、または配列であっても要素が割り当てられていない場合（Dim arr() のみでReDimされていない状態など）はFalseを返します。
    On Error GoTo NotAnArrayOrNotInitialized
    If IsArray(arr) Then
        Dim lBoundCheck As Long
        lBoundCheck = LBound(arr)
        LogMain_IsArrayInitialized = True ' LBoundがエラーを起こさなければ、配列は有効（空でもReDimされていればOK）
        Exit Function
    End If
NotAnArrayOrNotInitialized:
    LogMain_IsArrayInitialized = False
End Function
