' バージョン：v0.5.0
Option Explicit
' このモジュールは、「Config」シートから全ての設定情報を読み込み、検証し、g_configSettingsグローバル変数に格納する役割を担います。

' --- Private Helper Functions ---

Private Function GetCellValue(targetSheet As Worksheet, cellAddressString As String, callerProcName As String, ByRef errorFlag As Boolean, itemDescription As String, _
                            targetWorkbookForLog As Workbook, errorLogSheetNameForLog As String, _
                            Optional isRequiredField As Boolean = False, Optional validationType As String = "", _
                            Optional validationMinValue As Variant, Optional validationMaxValue As Variant) As Variant
    ' 指定されたセルから値を読み込み、必要に応じて検証・型変換を行います。
    ' Arguments:
    '   targetSheet: 値を読み込むワークシート
    '   cellAddressString: 値を読み込むセルアドレス
    '   callerProcName: 呼び出し元のプロシージャ名 (エラーレポート用)
    '   errorFlag: (ByRef) エラーが発生した場合にTrueに設定されるフラグ
    '   itemDescription: 設定項目の説明 (エラーレポート用)
    '   targetWorkbookForLog: エラーログ書き込み用のワークブック
    '   errorLogSheetNameForLog: エラーログシート名
    '   isRequiredField: (Optional) Trueの場合、必須項目として検証
    '   validationType: (Optional) "String", "Long", "Boolean", "CellAddress" など
    '   validationMinValue: (Optional) 数値型の場合の最小許容値
    '   validationMaxValue: (Optional) 数値型の場合の最大許容値
    ' Returns:
    '   Variant: 読み込んだ値 (型変換後)。エラー時はEmptyまたはデフォルト値。

    Dim rawValue As Variant
    Dim tempStr As String
    Dim tempLong As Long

    On Error Resume Next ' Cell read can error if sheet is protected, etc.
    rawValue = targetSheet.Range(cellAddressString).Value
    If Err.Number <> 0 Then
        Call ReportConfigError(errorFlag, callerProcName, cellAddressString, itemDescription & " の読み込み中にエラー発生: " & Err.Description, targetWorkbookForLog, errorLogSheetNameForLog)
        GetCellValue = Empty
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0 ' Restore default error handling

    ' Check for required field
    If isRequiredField Then
        If IsError(rawValue) Or IsEmpty(rawValue) Or Len(Trim(CStr(rawValue))) = 0 Then
            Call ReportConfigError(errorFlag, callerProcName, cellAddressString, itemDescription & " は必須項目ですが、値が空またはエラーです。", targetWorkbookForLog, errorLogSheetNameForLog)
            GetCellValue = Empty
            Exit Function
        End If
    ElseIf IsEmpty(rawValue) Then ' Optional field and genuinely empty
        GetCellValue = Empty
        Exit Function
    ElseIf Len(Trim(CStr(rawValue))) = 0 And validationType <> "Boolean" Then ' Optional string-like field that is blank after trim
        GetCellValue = Empty
        Exit Function
    End If


    ' Type validation and conversion
    Select Case validationType
        Case "String"
            GetCellValue = Trim(CStr(rawValue))
        Case "Long"
            If IsNumeric(rawValue) Then
                On Error Resume Next ' CLng can error for very large numbers that IsNumeric considers valid
                tempLong = CLng(rawValue)
                If Err.Number <> 0 Then
                    Call ReportConfigError(errorFlag, callerProcName, cellAddressString, itemDescription & " の値「" & rawValue & "」をLong型に変換できませんでした。", targetWorkbookForLog, errorLogSheetNameForLog)
                    GetCellValue = Empty
                    Err.Clear
                    Exit Function
                End If
                On Error GoTo 0

                If Not IsMissing(validationMinValue) Then
                    If tempLong < CLng(validationMinValue) Then
                        Call ReportConfigError(errorFlag, callerProcName, cellAddressString, itemDescription & " の値 (" & tempLong & ") が最小許容値 (" & validationMinValue & ") 未満です。", targetWorkbookForLog, errorLogSheetNameForLog)
                        GetCellValue = Empty
                        Exit Function
                    End If
                End If
                If Not IsMissing(validationMaxValue) Then
                    If tempLong > CLng(validationMaxValue) Then
                        Call ReportConfigError(errorFlag, callerProcName, cellAddressString, itemDescription & " の値 (" & tempLong & ") が最大許容値 (" & validationMaxValue & ") を超えています。", targetWorkbookForLog, errorLogSheetNameForLog)
                        GetCellValue = Empty
                        Exit Function
                    End If
                End If
                GetCellValue = tempLong
            Else
                Call ReportConfigError(errorFlag, callerProcName, cellAddressString, itemDescription & " の値「" & rawValue & "」は有効な数値ではありません。", targetWorkbookForLog, errorLogSheetNameForLog)
                GetCellValue = Empty
                Exit Function
            End If
        Case "Boolean"
            tempStr = UCase(Trim(CStr(rawValue)))
            If tempStr = "TRUE" Or tempStr = "-1" Or tempStr = "1" Then ' Common representations of True
                GetCellValue = True
            ElseIf tempStr = "FALSE" Or tempStr = "0" Then ' Common representations of False
                GetCellValue = False
            Else
                ' Invalid boolean string - return Empty. Caller handles default & warning.
                GetCellValue = Empty
            End If
        Case "CellAddress"
            tempStr = Trim(CStr(rawValue))
            If IsValidCellAddress(tempStr) Then
                GetCellValue = tempStr
            Else
                Call ReportConfigError(errorFlag, callerProcName, cellAddressString, itemDescription & " の値「" & rawValue & "」は有効なセルアドレスではありません。", targetWorkbookForLog, errorLogSheetNameForLog)
                GetCellValue = Empty
                Exit Function
            End If
        Case Else ' No validation type or unknown type
            GetCellValue = rawValue ' Return as is
    End Select
End Function

Private Sub LoadStringList(ByRef targetStringArray() As String, sourceSheet As Worksheet, columnLetter As String, firstRow As Long, lastRow As Long, _
                            callerProcName As String, listDescription As String, ByRef overallErrorFlag As Boolean, _
                            targetWorkbookForLog As Workbook, errorLogSheetForLog As String, Optional isRequired As Boolean = False)
    ' 指定された列範囲から文字列リストを読み込み、配列に格納します。
    Dim tempCollection As Collection
    Dim i As Long
    Dim cellValue As Variant
    Dim arrIndex As Long

    Set tempCollection = New Collection
    If g_configSettings.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M02_ConfigReader.LoadStringList - Loading " & listDescription & " from " & columnLetter & firstRow & ":" & columnLetter & lastRow

    For i = firstRow To lastRow
        cellValue = sourceSheet.Cells(i, columnLetter).Value
        If Not IsError(cellValue) And Not IsEmpty(cellValue) And Len(Trim(CStr(cellValue))) > 0 Then
            tempCollection.Add Trim(CStr(cellValue))
        End If
    Next i

    If tempCollection.Count > 0 Then
        ReDim targetStringArray(1 To tempCollection.Count)
        For arrIndex = 1 To tempCollection.Count
            targetStringArray(arrIndex) = tempCollection(arrIndex)
        Next arrIndex
    Else
        Erase targetStringArray ' Ensure array is uninitialized if empty
        If isRequired Then
            Call ReportConfigError(overallErrorFlag, callerProcName, columnLetter & firstRow & "-" & lastRow, listDescription & " は必須項目ですが、リストが空です。", targetWorkbookForLog, errorLogSheetForLog)
        End If
    End If
    If g_configSettings.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M02_ConfigReader.LoadStringList - " & listDescription & " loaded with " & tempCollection.Count & " items."
End Sub

Private Sub ReportConfigError(ByRef overallErrorFlag As Boolean, errorSourceContext As String, errorSourceCell As String, errorMessageText As String, _
                                ByVal wbForLog As Workbook, ByVal errorLogSheetNameForLog As String, Optional isFatal As Boolean = True, Optional errorLevel As String = "")
    ' 設定読み込みエラーを報告し、必要に応じて全体エラーフラグを立てます。
    If isFatal Then overallErrorFlag = True

    Dim logSourceModule As String
    logSourceModule = "M02_ConfigReader." & errorSourceContext ' errorSourceContext will be like "LoadConfiguration (A-1)"

    Dim levelToLog As String
    If Len(errorLevel) > 0 Then
        levelToLog = errorLevel
    Else
        If isFatal Then
            levelToLog = "ERROR" ' Or "CONFIG_ERROR_FATAL"
        Else
            levelToLog = "WARNING" ' Or "CONFIG_WARNING"
        End If
    End If

    If DEBUG_MODE_ERROR Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - CONFIG_REPORT: Level='" & levelToLog & "', Module='" & logSourceModule & "', Cell='" & errorSourceCell & "', Message='" & errorMessageText & "'"

    Call M04_LogWriter.SafeWriteErrorLog(levelToLog, wbForLog, errorLogSheetNameForLog, logSourceModule, errorSourceCell, errorMessageText, 0, "")
End Sub

Private Function IsValidCellAddress(cellAddressString As String) As Boolean
    ' セルアドレス文字列が有効かどうかを検証します。
    Dim tempVal As Variant
    If Trim(cellAddressString) = "" Then
        IsValidCellAddress = False
        Exit Function
    End If
    On Error Resume Next
    ' ISREF returns True for valid references, False for #VALUE! (e.g. "A"), Error for others (e.g. "1A")
    tempVal = Application.Evaluate("ISREF(" & cellAddressString & ")")
    If Err.Number <> 0 Then
        IsValidCellAddress = False
    Else
        If IsError(tempVal) Then ' e.g. ISREF(1A) gives #NAME? error, Evaluate returns Error 2029
            IsValidCellAddress = False
        Else
            IsValidCellAddress = CBool(tempVal) ' tempVal should be True or False
        End If
    End If
    On Error GoTo 0
End Function

Private Function ConfigReader_IsArrayInitialized(arr As Variant) As Boolean
    ' 配列が有効に初期化されているか（少なくとも1つの要素を持つか）を確認します。
    ' Variant型が配列でない場合、または配列であっても要素が割り当てられていない場合（Dim arr() のみでReDimされていない状態など）はFalseを返します。
    On Error GoTo NotAnArrayOrNotInitialized
    If IsArray(arr) Then
        Dim lBoundCheck As Long
        lBoundCheck = LBound(arr)
        ConfigReader_IsArrayInitialized = True
        Exit Function
    End If
NotAnArrayOrNotInitialized:
    ConfigReader_IsArrayInitialized = False
End Function

Private Sub LoadProcessDetailsLimited(ByRef configS As tConfigSettings, srcSheet As Worksheet, ByRef errFlag As Boolean, wbLog As Workbook, errLogName As String)
    ' 「Config」シートのJ列(管内1)、K列(管内2)からパターン"1"に対応する情報を読み込みます。(ステップ4限定)
    ' デバッグモード時は、L列(分類1)、M列(分類2)、N列(分類3)も読み込んでログ出力します。
    ' 引数:
    '   configS: (I/O) tConfigSettings型。読み込んだ設定が格納されます。
    '   srcSheet: (I) Worksheet型。読み込み元のConfigシートオブジェクト。
    '   errFlag: (I/O) Boolean型。エラー発生時にTrueに設定されます。
    '   wbLog: (I) Workbook型。エラーログ書き込み用のワークブック。
    '   errLogName: (I) String型。エラーログシート名。

    Dim i As Long
    Dim valJ As Variant, valK As Variant
    Dim valL As Variant, valM As Variant, valN As Variant ' For debug logging only

    If configS.ProcessesPerDay <= 0 Then Exit Sub ' Should not happen if validation in LoadConfiguration is correct

    If configS.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M02_ConfigReader.LoadProcessDetailsLimited - Reading J, K" & IIf(configS.DebugModeFlag, ", L, M, N", "") & " cols from row 129 for " & configS.ProcessesPerDay & " processes."

    For i = 0 To configS.ProcessesPerDay - 1
        ' Kankatsu1 (J column)
        valJ = GetCellValue(srcSheet, "J" & (129 + i), "LoadProcessDetailsLimited", errFlag, "管内1 (J" & (129 + i) & ")", wbLog, errLogName, False, "String")
        If errFlag Then Exit For ' Stop if GetCellValue reported a fatal error
        configS.ProcessDetails(i).Kankatsu1 = CStr(valJ)

        ' Kankatsu2 (K column)
        valK = GetCellValue(srcSheet, "K" & (129 + i), "LoadProcessDetailsLimited", errFlag, "管内2 (K" & (129 + i) & ")", wbLog, errLogName, False, "String")
        If errFlag Then Exit For
        configS.ProcessDetails(i).Kankatsu2 = CStr(valK)

        If configS.DebugModeFlag Then ' Only read and log Bunrui if DebugMode is ON
            valL = GetCellValue(srcSheet, "L" & (129 + i), "LoadProcessDetailsLimited", errFlag, "分類1 (L" & (129 + i) & ") for debug", wbLog, errLogName, False, "String")
            If errFlag Then Exit For
            valM = GetCellValue(srcSheet, "M" & (129 + i), "LoadProcessDetailsLimited", errFlag, "分類2 (M" & (129 + i) & ") for debug", wbLog, errLogName, False, "String")
            If errFlag Then Exit For
            valN = GetCellValue(srcSheet, "N" & (129 + i), "LoadProcessDetailsLimited", errFlag, "分類3 (N" & (129 + i) & ") for debug", wbLog, errLogName, False, "String")
            If errFlag Then Exit For
            Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_CONFIG_DETAIL:   Process " & i & ": Kankatsu1='" & CStr(valJ) & "', Kankatsu2='" & CStr(valK) & "', Bunrui1='" & CStr(valL) & "', Bunrui2='" & CStr(valM) & "', Bunrui3='" & CStr(valN) & "'"
        End If
    Next i
End Sub

Private Sub LoadProcessPatternColNumbersLimited(ByRef configS As tConfigSettings, srcSheet As Worksheet, ByRef errFlag As Boolean, wbLog As Workbook, errLogName As String)
    ' 「Config」シートのO列からパターン"1"に対応する工程列数を読み込みます。(ステップ4限定)
    ' Arguments:
    '   configS: (I/O) tConfigSettings型。読み込んだ設定が格納されます。
    '   srcSheet: (I) Worksheet型。読み込み元のConfigシートオブジェクト。
    '   errFlag: (I/O) Boolean型。エラー発生時にTrueに設定されます。
    '   wbLog: (I) Workbook型。エラーログ書き込み用のワークブック。
    '   errLogName: (I) String型。エラーログシート名。

    Dim i As Long
    Dim colCountVal As Variant

    If configS.ProcessesPerDay <= 0 Then Exit Sub

    If configS.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M02_ConfigReader.LoadProcessPatternColNumbersLimited - Reading O col from row 129 for " & configS.ProcessesPerDay & " processes for pattern 1."

    For i = 0 To configS.ProcessesPerDay - 1
        colCountVal = GetCellValue(srcSheet, "O" & (129 + i), "LoadProcessPatternColNumbersLimited", errFlag, "工程列数 (O" & (129 + i) & ")", wbLog, errLogName, True, "Long", 0) ' isRequired=True, MinValue=0

        If errFlag Then Exit For ' Stop if GetCellValue reported a fatal error (e.g. non-numeric for required Long)

        If Not IsEmpty(colCountVal) Then
            configS.ProcessPatternColNumbers(1)(i) = CLng(colCountVal)
        Else
            ' This case should ideally not be reached if isRequiredField is True in GetCellValue,
            ' as GetCellValue would set errFlag=True and exit.
            ' However, if it somehow passes (e.g. isRequired=False), default to 0.
            configS.ProcessPatternColNumbers(1)(i) = 0 ' Default to 0 if not specified or invalid
            If DEBUG_MODE_WARNING Then Call ReportConfigError(errFlag, "LoadProcessPatternColNumbersLimited", "O" & (129 + i), "工程列数値が不正または空のため0を適用 (非致命的扱い)", wbLog, errLogName, False)
        End If
    Next i
End Sub

' --- Stubs for future implementation ---
Private Function ParseOffset(offsetString As String, ByRef resultOffset As tOffset, ByRef errorFlag As Boolean, callerProcName As String, itemDesc As String, ByVal wbForLog As Workbook, ByVal errorLogSheetNameForLog As String) As Boolean
    ' オフセット文字列("行,列")を解析し、tOffset型に格納します。書式不正の場合はエラーを報告しFalseを返します。
    ' Empty string is considered a valid parse (resulting in 0,0 offset), indicating no offset.
    ' The caller (GetCellValue or LoadConfiguration) should decide if an empty/zero offset is permissible for a required field.

    ParseOffset = False ' Default to failure unless explicitly successful
    resultOffset.Row = 0
    resultOffset.Col = 0

    Dim parts() As String
    Dim strVal As String

    strVal = Trim(offsetString)

    If Len(strVal) = 0 Then
        ParseOffset = True ' Successfully "parsed" an empty string; offset remains 0,0
        Exit Function
    End If

    parts = Split(strVal, ",")

    If UBound(parts) <> 1 Then
        Call ReportConfigError(errorFlag, callerProcName, itemDesc, "オフセット書式不正 (カンマ区切り2要素でない): '" & offsetString & "'", wbForLog, errorLogSheetNameForLog, True, "ERROR_CONFIG_PARSE")
        Exit Function ' Returns False
    End If

    If Not IsNumeric(Trim(parts(0))) Or Not IsNumeric(Trim(parts(1))) Then
        Call ReportConfigError(errorFlag, callerProcName, itemDesc, "オフセット値が数値でない: '" & offsetString & "'", wbForLog, errorLogSheetNameForLog, True, "ERROR_CONFIG_PARSE")
        Exit Function ' Returns False
    End If

    resultOffset.Row = CLng(Trim(parts(0)))
    resultOffset.Col = CLng(Trim(parts(1)))
    ParseOffset = True ' Successfully parsed
End Function

Private Sub LoadProcessPatternColNumbers(ByRef configStruct As tConfigSettings, sourceSheet As Worksheet, callerProcName As String, ByRef errorFlag As Boolean, targetWorkbookForLog As Workbook, errorLogSheetForLog As String)
    ' TODO: Implement loading of process pattern column numbers
    If DEBUG_MODE_WARNING Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - WARNING: M02_ConfigReader.LoadProcessPatternColNumbers - STUB called."
    Exit Sub
End Sub

Private Sub LoadProcessDetails(ByRef configStruct As tConfigSettings, sourceSheet As Worksheet, callerProcName As String, ByRef errorFlag As Boolean, targetWorkbookForLog As Workbook, errorLogSheetForLog As String)
    ' TODO: Implement loading of process details
    If DEBUG_MODE_WARNING Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - WARNING: M02_ConfigReader.LoadProcessDetails - STUB called."
    Exit Sub
End Sub


' --- Main Public Function ---

Public Function LoadConfiguration(ByRef configStruct As tConfigSettings, ByVal targetWorkbook As Workbook, ByVal configSheetName As String) As Boolean
    ' 「Config」シートから設定情報を読み込み、引数のconfigStruct構造体に格納します。
    ' Aセクション「全般設定」およびBセクション「工程表ファイル内 設定」の項目を読み込み、検証します。
    ' 引数:
    '   configStruct: (I/O) tConfigSettings型。読み込んだ設定が格納されます。
    '   targetWorkbook: (I) Workbook型。Configシートが含まれるワークブック。
    '   configSheetName: (I) String型。読み込み元のConfigシート名。
    ' 戻り値:
    '   Boolean: 全ての設定項目の読み込みと検証に成功した場合はTrue、それ以外はFalse。

    Dim wsConfig As Worksheet
    Dim m_errorOccurred As Boolean ' Local error flag for this loading process
    Dim tempVal As Variant

    m_errorOccurred = False
    LoadConfiguration = False ' Default to failure
    On Error GoTo LoadConfiguration_Error_MainHandler

    ' --- Configシートオブジェクト取得 ---
    ' Note: TraceDebugEnabled is not yet available from configStruct here, using g_configSettings as a fallback (though it's also not set yet)
    ' This initial log might better use DEBUG_MODE_ERROR or be moved after TraceDebugEnabled is read. For now, using g_configSettings.
    If g_configSettings.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M02_ConfigReader.LoadConfiguration - Attempting to get Config sheet: '" & configSheetName & "'"
    On Error Resume Next ' Temporarily change error handling
    Set wsConfig = targetWorkbook.Worksheets(configSheetName)
    On Error GoTo LoadConfiguration_Error_MainHandler ' Restore main error handler

    If wsConfig Is Nothing Then
        ' ErrorLogSheetName is not yet available from configStruct, SafeWriteErrorLog will use its own fallback.
        Call ReportConfigError(m_errorOccurred, "LoadConfiguration", configSheetName, "Configシートが見つかりません。処理を続行できません。", targetWorkbook, "") ' Pass empty for sheet name, SafeWrite handles it
        MsgBox "Configシート「" & configSheetName & "」が見つかりません。処理を中断します。", vbCritical, "設定エラー"
        Exit Function ' Returns False
    End If
    configStruct.ConfigSheetFullName = targetWorkbook.FullName & " | " & wsConfig.Name ' Use a clear separator
    ' TraceDebugEnabled not yet read, use g_configSettings for this specific log line.
    If g_configSettings.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M02_ConfigReader.LoadConfiguration - Config sheet found: '" & configStruct.ConfigSheetFullName & "'"

    ' --- A. 全般設定 ---
    ' TraceDebugEnabled not yet read for this section's title print.
    If g_configSettings.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M02_ConfigReader.LoadConfiguration - Reading Section A: General Settings"

    ' A-1: デバッグモードフラグ (O3)
    tempVal = GetCellValue(wsConfig, "O3", "LoadConfiguration (A-1)", m_errorOccurred, "デバッグモードフラグ", targetWorkbook, configStruct.ErrorLogSheetName, False, "Boolean")
    If Not m_errorOccurred Then ' Check if GetCellValue itself caused a fatal error
        If Not IsEmpty(tempVal) Then
            configStruct.DebugModeFlag = tempVal
        Else ' Value was empty or not a valid boolean string (but not a read error from GetCellValue)
            configStruct.DebugModeFlag = False ' Default value
            If DEBUG_MODE_WARNING Then Call ReportConfigError(m_errorOccurred, "LoadConfiguration (A-1)", "O3", "デバッグモードフラグ(O3)が不正または空のためFalseを適用", targetWorkbook, configStruct.ErrorLogSheetName, False, "WARNING_CONFIG_DEFAULT")
        End If
    End If
    ' Note: m_errorOccurred could be true if GetCellValue had a fatal read error.

    ' A-Trace: 詳細デバッグ出力フラグ (O4)
    tempVal = GetCellValue(wsConfig, "O4", "LoadConfiguration (A-Trace)", m_errorOccurred, "詳細デバッグ出力フラグ (O4)", targetWorkbook, configStruct.ErrorLogSheetName, False, "Boolean")
    If Not m_errorOccurred Then
        If Not IsEmpty(tempVal) Then
            configStruct.TraceDebugEnabled = tempVal
        Else
            configStruct.TraceDebugEnabled = False ' Default to False
            If DEBUG_MODE_WARNING Then Call ReportConfigError(m_errorOccurred, "LoadConfiguration (A-Trace)", "O4", "詳細デバッグ出力フラグ(O4)が不正または未設定のためFalseを適用", targetWorkbook, configStruct.ErrorLogSheetName, False, "WARNING_CONFIG_DEFAULT")
        End If
    End If
    ' Now TraceDebugEnabled is set in configStruct and can be used by subsequent DEBUG_DETAIL prints in this Sub

    ' A-EnableSheetLogging: 汎用ログシートへの出力有効フラグ (O5)
    tempVal = GetCellValue(wsConfig, "O5", "LoadConfiguration (A-EnableSheetLogging)", m_errorOccurred, "汎用ログシート出力有効フラグ (O5)", targetWorkbook, configStruct.ErrorLogSheetName, False, "Boolean")
    If Not m_errorOccurred Then
        If Not IsEmpty(tempVal) Then
            configStruct.EnableSheetLogging = tempVal
        Else
            configStruct.EnableSheetLogging = False ' Default value
            If DEBUG_MODE_WARNING Then Call ReportConfigError(m_errorOccurred, "LoadConfiguration (A-EnableSheetLogging)", "O5", "汎用ログシート出力有効フラグ(O5)が不正または空のためFalseを適用", targetWorkbook, configStruct.ErrorLogSheetName, False, "WARNING_CONFIG_DEFAULT")
        End If
    ElseIf IsEmpty(configStruct.EnableSheetLogging) Then ' If GetCellValue had a more serious error and didn't set it
        configStruct.EnableSheetLogging = False ' Ensure default
    End If

    ' A-LogSheetName: 汎用ログシート名 (O42)
    configStruct.LogSheetName = GetCellValue(wsConfig, "O42", "LoadConfiguration (A-LogSheetName)", m_errorOccurred, "汎用ログシート名 (O42)", targetWorkbook, configStruct.ErrorLogSheetName, False, "String")
    ' If sheet logging is enabled but no log sheet name is provided, log a warning.
    If Not m_errorOccurred And configStruct.EnableSheetLogging And Len(configStruct.LogSheetName) = 0 Then
        Call ReportConfigError(m_errorOccurred, "LoadConfiguration (A-LogSheetNameCheck)", "O42", "汎用ログシート出力が有効(O5=TRUE)ですが、ログシート名(O42)が未指定です。シートログは出力されません。", targetWorkbook, configStruct.ErrorLogSheetName, False, "WARNING_CONFIG_LOGGING_DISABLED")
        ' configStruct.EnableSheetLogging = False ' Optionally force it False here
    End If

    ' A-2: デフォルトフォルダパス (O12)
    configStruct.DefaultFolderPath = GetCellValue(wsConfig, "O12", "LoadConfiguration (A-2)", m_errorOccurred, "デフォルトフォルダパス", targetWorkbook, configStruct.ErrorLogSheetName, False, "String")

    ' A-3: 抽出結果出力シート名 (O43) - Required
    configStruct.OutputSheetName = GetCellValue(wsConfig, "O43", "LoadConfiguration (A-3)", m_errorOccurred, "抽出結果出力シート名", targetWorkbook, configStruct.ErrorLogSheetName, True, "String")

    ' A-4: 検索条件ログシート名 (O44) - Required
    configStruct.SearchConditionLogSheetName = GetCellValue(wsConfig, "O44", "LoadConfiguration (A-4)", m_errorOccurred, "検索条件ログシート名", targetWorkbook, configStruct.ErrorLogSheetName, True, "String")

    ' A-5: エラーログシート名 (O45) - Required. This is crucial.
    ' Note: errorLogSheetNameForLog is passed as configStruct.ErrorLogSheetName, which is empty at this point. SafeWriteErrorLog in ReportConfigError will use its fallback.
    configStruct.ErrorLogSheetName = GetCellValue(wsConfig, "O45", "LoadConfiguration (A-5)", m_errorOccurred, "エラーログシート名", targetWorkbook, configStruct.ErrorLogSheetName, True, "String")
    If m_errorOccurred And Len(configStruct.ErrorLogSheetName) = 0 Then
         MsgBox "エラーログシート名(O45)の読み込みに失敗しました。ログ機能が利用できません。処理を中断します。", vbCritical, "致命的な設定エラー"
         Exit Function ' Returns False
    End If
    ' From now on, configStruct.ErrorLogSheetName can be used by ReportConfigError if it was successfully read.

    ' A-6: 設定ファイルシート名 (O46)
    configStruct.ConfigSheetName = GetCellValue(wsConfig, "O46", "LoadConfiguration (A-6)", m_errorOccurred, "設定ファイルシート名", targetWorkbook, configStruct.ErrorLogSheetName, False, "String")
    If Not m_errorOccurred And Len(configStruct.ConfigSheetName) = 0 Then configStruct.ConfigSheetName = configSheetName ' Default to passed name if cell is empty

    ' A-7: 工程パターンデータ取得方法 (O122)
    tempVal = GetCellValue(wsConfig, "O122", "LoadConfiguration (A-7)", m_errorOccurred, "工程パターンデータ取得方法", targetWorkbook, configStruct.ErrorLogSheetName, False, "Boolean")
    If Not m_errorOccurred Then
        If Not IsEmpty(tempVal) Then
            configStruct.GetPatternDataMethod = tempVal
        Else
            configStruct.GetPatternDataMethod = False ' Default value (VBA method)
            If DEBUG_MODE_WARNING Then Call ReportConfigError(m_errorOccurred, "LoadConfiguration (A-7)", "O122", "工程パターンデータ取得方法(O122)が不正または空のためFalse(VBA方式)を適用", targetWorkbook, configStruct.ErrorLogSheetName, False, "WARNING_CONFIG_DEFAULT")
        End If
    End If

    ' A-ExtractIfEmpty: 作業員空でも抽出フラグ (O241) - Technically a filter condition, but read with general flags
    tempVal = GetCellValue(wsConfig, "O241", "LoadConfiguration (A-ExtractIfEmpty)", m_errorOccurred, "作業員空でも抽出フラグ (O241)", targetWorkbook, configStruct.ErrorLogSheetName, False, "Boolean")
    If Not m_errorOccurred Then ' Check if GetCellValue itself caused a fatal error
        If Not IsEmpty(tempVal) Then
            configStruct.ExtractIfWorkersEmpty = tempVal
        Else ' Value was empty or not a valid boolean string
            configStruct.ExtractIfWorkersEmpty = False ' Default value
            If DEBUG_MODE_WARNING Then Call ReportConfigError(m_errorOccurred, "LoadConfiguration (A-ExtractIfEmpty)", "O241", "作業員空でも抽出フラグ(O241)が不正または空のためFalseを適用", targetWorkbook, configStruct.ErrorLogSheetName, False, "WARNING_CONFIG_DEFAULT")
        End If
    End If

    ' --- B. 工程表ファイル内 設定 ---
    If configStruct.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M02_ConfigReader.LoadConfiguration - Reading Section B: Schedule File Settings"

    ' B-1: 工程表内 検索対象シート名リスト (O66-O75) - Required List
    Call LoadStringList(configStruct.TargetSheetNames, wsConfig, "O", 66, 75, "LoadConfiguration (B-1)", "検索対象シート名リスト", m_errorOccurred, targetWorkbook, configStruct.ErrorLogSheetName, True)

    ' B-2: 工程表ヘッダー行数 (O87) - Required, Min 0
    configStruct.HeaderRowCount = GetCellValue(wsConfig, "O87", "LoadConfiguration (B-2)", m_errorOccurred, "工程表ヘッダー行数", targetWorkbook, configStruct.ErrorLogSheetName, True, "Long", 0)

    ' B-3: 工程表ヘッダー列数 (O88) - Required, Min 0
    configStruct.HeaderColCount = GetCellValue(wsConfig, "O88", "LoadConfiguration (B-3)", m_errorOccurred, "工程表ヘッダー列数", targetWorkbook, configStruct.ErrorLogSheetName, True, "Long", 0)

    ' B-4: 1日のデータが占める行数 (O89) - Required, Min 1
    configStruct.RowsPerDay = GetCellValue(wsConfig, "O89", "LoadConfiguration (B-4)", m_errorOccurred, "1日のデータが占める行数", targetWorkbook, configStruct.ErrorLogSheetName, True, "Long", 1)

    ' B-5: 1シート内の最大日数 (O90) - Required, Min 1
    configStruct.MaxDaysPerSheet = GetCellValue(wsConfig, "O90", "LoadConfiguration (B-5)", m_errorOccurred, "1シート内の最大日数", targetWorkbook, configStruct.ErrorLogSheetName, True, "Long", 1)

    ' B-6: 「年」のセルアドレス (O101) - Required, CellAddress
    configStruct.YearCellAddress = GetCellValue(wsConfig, "O101", "LoadConfiguration (B-6)", m_errorOccurred, "「年」のセルアドレス", targetWorkbook, configStruct.ErrorLogSheetName, True, "CellAddress")

    ' B-7: 「月」のセルアドレス (O102) - Required, CellAddress
    configStruct.MonthCellAddress = GetCellValue(wsConfig, "O102", "LoadConfiguration (B-7)", m_errorOccurred, "「月」のセルアドレス", targetWorkbook, configStruct.ErrorLogSheetName, True, "CellAddress")

    ' B-8: 「日」の値がある列文字 (O103) - Required, String, Valid Column Letter(s)
    tempVal = GetCellValue(wsConfig, "O103", "LoadConfiguration (B-8)", m_errorOccurred, "「日」の値がある列文字", targetWorkbook, configStruct.ErrorLogSheetName, True, "String")
    If Not m_errorOccurred Then ' Only proceed if GetCellValue didn't set an error
        If Len(CStr(tempVal)) > 0 And Len(CStr(tempVal)) <= 3 And CStr(tempVal) Like Application.WorksheetFunction.Rept("[A-Za-z]", Len(CStr(tempVal))) Then
            configStruct.DayColumnLetter = UCase(CStr(tempVal))
        Else
            Call ReportConfigError(m_errorOccurred, "LoadConfiguration (B-8)", "O103", "「日」の列文字(O103)「" & tempVal & "」がExcelの列文字として不正です。", targetWorkbook, configStruct.ErrorLogSheetName)
            configStruct.DayColumnLetter = "A" ' Fallback, m_errorOccurred is now True
        End If
    ElseIf Len(configStruct.DayColumnLetter) = 0 Then ' If error occurred and DayColumnLetter is still not set
         configStruct.DayColumnLetter = "A" ' Fallback
    End If

    ' B-9: 「日」の値の行オフセット (O104) - Required, Long, Min 0
    configStruct.DayRowOffset = GetCellValue(wsConfig, "O104", "LoadConfiguration (B-9)", m_errorOccurred, "「日」の値の行オフセット", targetWorkbook, configStruct.ErrorLogSheetName, True, "Long", 0)

    ' B-10: 1日の工程数 (O114) - Required, Long, Min 1
    configStruct.ProcessesPerDay = GetCellValue(wsConfig, "O114", "LoadConfiguration (B-10)", m_errorOccurred, "1日の工程数", targetWorkbook, configStruct.ErrorLogSheetName, True, "Long", 1)

    ' --- C. 工程パターン定義 (ステップ4限定読み込み) ---
    If configStruct.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M02_ConfigReader.LoadConfiguration - Reading Section C (Limited for Step 4)"
    configStruct.CurrentPatternIdentifier = "1" ' 固定でパターン"1"を使用 (仕様変更: FileProcessPatternToUse -> CurrentPatternIdentifier)

    If configStruct.ProcessesPerDay > 0 Then ' 配列をReDimする前に要素数を確認
        ReDim configStruct.ProcessDetails(0 To configStruct.ProcessesPerDay - 1) As tProcessDetail
        ' ProcessPatternColNumbers is a jagged array: (PatternIndex)(ProcessIndex)
        ' For step 4, we only care about pattern "1". PatternIndex will be 1.
        ReDim configStruct.ProcessPatternColNumbers(1 To 1) ' This outer ReDim for the Variant array is fine

        ' Correctly ReDim the inner Long array for pattern 1
        Dim tempPattern1Cols() As Long ' Declare a temporary dynamic array of Long
        If configStruct.ProcessesPerDay > 0 Then
            ReDim tempPattern1Cols(0 To configStruct.ProcessesPerDay - 1) As Long
        Else
            ' Create an empty but initialized array if ProcessesPerDay is 0 or less.
            ' This prevents errors if other code tries to access LBound/UBound later,
            ' though logic should ideally check ProcessesPerDay before looping.
            ReDim tempPattern1Cols(0 To -1) As Long ' Standard way to make an empty initialized array
        End If
        configStruct.ProcessPatternColNumbers(1) = tempPattern1Cols ' Assign this Long array to the Variant slot

        Call LoadProcessDetailsLimited(configStruct, wsConfig, m_errorOccurred, targetWorkbook, configStruct.ErrorLogSheetName)
        If m_errorOccurred Then GoTo FinalConfigCheck ' Stop further processing in this section if error occurred

        Call LoadProcessPatternColNumbersLimited(configStruct, wsConfig, m_errorOccurred, targetWorkbook, configStruct.ErrorLogSheetName)
        If m_errorOccurred Then GoTo FinalConfigCheck
    Else
        If DEBUG_MODE_WARNING Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - WARNING: M02_ConfigReader.LoadConfiguration - ProcessesPerDay is 0 or less (" & configStruct.ProcessesPerDay & "). Skipping C section array ReDims and limited loading."
        ' ProcessesPerDayが0以下の場合、関連配列のReDimは行わない
    End If

    ' --- E. 処理対象ファイル定義 (ステップ4限定読み込み) ---
    If configStruct.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M02_ConfigReader.LoadConfiguration - Reading Section E (Limited for Step 4)"
    ReDim configStruct.TargetFileFolderPaths(0 To 0) As String ' Only one file for now

    Dim filePathP557Val As Variant
    filePathP557Val = GetCellValue(wsConfig, "P557", "LoadConfiguration (E-1)", m_errorOccurred, "処理対象ファイルパス(P557)", targetWorkbook, configStruct.ErrorLogSheetName, True, "String") ' isRequired = True

    If Not m_errorOccurred Then ' Check if GetCellValue itself caused a fatal error
        If Not IsEmpty(filePathP557Val) And Len(CStr(filePathP557Val)) > 0 Then
            configStruct.TargetFileFolderPaths(0) = CStr(filePathP557Val)
            If configStruct.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M02_ConfigReader.LoadConfiguration - TargetFileFolderPaths(0) (P557): '" & configStruct.TargetFileFolderPaths(0) & "'"
        Else ' This specific case (IsEmpty or empty string for a required field) should be caught by GetCellValue
             ' However, if isRequiredField was accidentally False, this would be a fallback.
             ' For safety, ensure error is flagged if required field is empty.
            Call ReportConfigError(m_errorOccurred, "LoadConfiguration (E-1)", "P557", "処理対象ファイルパス(P557)が必須ですが空です。", targetWorkbook, configStruct.ErrorLogSheetName, True)
        End If
    End If
    ' If m_errorOccurred is True due to GetCellValue failing, a value in TargetFileFolderPaths(0) might be meaningless.

    ' --- F. 抽出データオフセット定義 ---
    If configStruct.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M02_ConfigReader.LoadConfiguration - Reading Section F: Extraction Data Offset Definition (Array Method)"
    Dim itemName As String, offsetStr As String
    Dim tempOffset As tOffset
    Dim fIdx As Long ' Loop counter for F-section items
    Dim actualOffsetCount As Long: actualOffsetCount = 0

    ' Erase and prepare arrays for fresh load
    Erase configStruct.OffsetItemMasterNames
    Erase configStruct.OffsetDefinitions
    Erase configStruct.IsOffsetOriginallyEmptyFlags
    ' Initial ReDim to 0 elements for the first ReDim Preserve to work if needed, or handle first element separately.
    ' A common practice is to ReDim to a known small size if reading a fixed number of items,
    ' or use a counter and ReDim Preserve from the start.
    ' Reading a fixed potential of 11 items (N778-N788)

    For fIdx = 0 To 10 ' For 11 items, from row 778 to 788
        itemName = Trim(CStr(wsConfig.Range("N" & (778 + fIdx)).Value))
        offsetStr = Trim(CStr(wsConfig.Range("O" & (778 + fIdx)).Value))

        If Len(itemName) > 0 Then ' Only process if an item name is defined in N column
            actualOffsetCount = actualOffsetCount + 1
            ReDim Preserve configStruct.OffsetItemMasterNames(1 To actualOffsetCount)
            ReDim Preserve configStruct.OffsetDefinitions(1 To actualOffsetCount)
            ReDim Preserve configStruct.IsOffsetOriginallyEmptyFlags(1 To actualOffsetCount)

            configStruct.OffsetItemMasterNames(actualOffsetCount) = itemName
            configStruct.IsOffsetOriginallyEmptyFlags(actualOffsetCount) = (Len(offsetStr) = 0)

            If Not configStruct.IsOffsetOriginallyEmptyFlags(actualOffsetCount) Then
                ' Pass m_errorOccurred ByRef. If ParseOffset sets it to True, it's a fatal error for LoadConfiguration.
                If ParseOffset(offsetStr, tempOffset, m_errorOccurred, "LoadConfiguration (F-" & fIdx + 1 & ")", itemName & " オフセット(O" & (778 + fIdx) & ")", targetWorkbook, configStruct.ErrorLogSheetName) Then
                    configStruct.OffsetDefinitions(actualOffsetCount) = tempOffset
                Else
                    ' ParseOffset returned False. m_errorOccurred should be True if it was a real parse error.
                    configStruct.OffsetDefinitions(actualOffsetCount).Row = 0 ' Default to 0,0
                    configStruct.OffsetDefinitions(actualOffsetCount).Col = 0
                    ' If ParseOffset didn't set m_errorOccurred (e.g. some unexpected case), flag it.
                    If Not m_errorOccurred Then
                         Call ReportConfigError(m_errorOccurred, "LoadConfiguration (F-" & fIdx + 1 & ")", itemName & " オフセット(O" & (778 + fIdx) & ")", "オフセット値の解析に予期せず失敗(ParseOffsetがFalseだがエラー未セット): '" & offsetStr & "'", targetWorkbook, configStruct.ErrorLogSheetName, True, "ERROR_CONFIG_UNEXPECTED_PARSE")
                    End If
                End If
            Else
                configStruct.OffsetDefinitions(actualOffsetCount).Row = 0 ' Default for originally empty string
                configStruct.OffsetDefinitions(actualOffsetCount).Col = 0
            End If
            If configStruct.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_CONFIG_DETAIL:   F. Offset Item " & actualOffsetCount & " (" & itemName & ", N" & (778 + fIdx) & "): '" & offsetStr & "' -> R:" & configStruct.OffsetDefinitions(actualOffsetCount).Row & ", C:" & configStruct.OffsetDefinitions(actualOffsetCount).Col & ", IsEmptyOrig: " & configStruct.IsOffsetOriginallyEmptyFlags(actualOffsetCount)
            If m_errorOccurred Then GoTo FinalConfigCheck ' Propagate error to exit LoadConfiguration
        Else
            ' Optional: Log if N-column is empty but O-column has value (likely misconfiguration)
            If Len(offsetStr) > 0 And configStruct.TraceDebugEnabled Then
                Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE: M02_ConfigReader.LoadConfiguration - Offset string '" & offsetStr & "' found in O" & (778 + fIdx) & " but no item name in N" & (778 + fIdx) & ". Skipping."
            End If
        End If
    Next fIdx

    ' After loop, if actualOffsetCount is 0, it means N778:N788 were all empty.
    ' This might be an error if offsets are considered mandatory overall.
    If actualOffsetCount = 0 Then
        If configStruct.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_TRACE: M02_ConfigReader.LoadConfiguration - No offset items defined in N778:N788. Initializing offset arrays as empty."
        ' Explicitly ReDim to an empty, initialized state.
        ' This correctly handles the case where the ReDim Preserve loop was never entered.
        ReDim configStruct.OffsetItemMasterNames(1 To 0)
        ReDim configStruct.OffsetDefinitions(1 To 0)
        ReDim configStruct.IsOffsetOriginallyEmptyFlags(1 To 0)
    End If

    ' --- G. 出力シート設定 ---
    Dim i As Long ' Declare i for G section loop
    If configStruct.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M02_ConfigReader.LoadConfiguration - Reading Section G: Output Sheet Settings"
    configStruct.OutputHeaderRowCount = GetCellValue(wsConfig, "O811", "LoadConfiguration (G-1)", m_errorOccurred, "出力シートヘッダー行数", targetWorkbook, configStruct.ErrorLogSheetName, True, "Long", 1, 10)

    If Not m_errorOccurred And configStruct.OutputHeaderRowCount > 0 Then
        ReDim configStruct.OutputHeaderContents(1 To configStruct.OutputHeaderRowCount) As String
        For i = 1 To configStruct.OutputHeaderRowCount ' Loop variable i is fine here
            Dim headerCellAddress As String: headerCellAddress = "O" & (811 + i)
            Dim headerVal As String
            headerVal = Trim(CStr(GetCellValue(wsConfig, headerCellAddress, "LoadConfiguration (G-2)", m_errorOccurred, "出力シートヘッダー内容 " & i & "行目 (" & headerCellAddress & ")", targetWorkbook, configStruct.ErrorLogSheetName, False, "String")))

            If m_errorOccurred And Len(headerVal) = 0 Then
                m_errorOccurred = False ' Clear error if GetCellValue flagged it for an empty optional header line
            End If
            configStruct.OutputHeaderContents(i) = headerVal
        Next i
    End If

    Dim outputOpt As String
    outputOpt = UCase(Trim(CStr(GetCellValue(wsConfig, "O1124", "LoadConfiguration (G-3)", m_errorOccurred, "出力データオプション", targetWorkbook, configStruct.ErrorLogSheetName, False, "String"))))
    If Not m_errorOccurred Then ' Only validate if GetCellValue didn't cause a fatal error
        If outputOpt = "リセット" Or outputOpt = "引継ぎ" Then
            configStruct.OutputDataOption = outputOpt
        Else
            configStruct.OutputDataOption = "リセット" ' Default
            If Len(outputOpt) > 0 Then ' Log warning only if there was some (invalid) value
                Call ReportConfigError(m_errorOccurred, "LoadConfiguration (G-3)", "O1124", "出力データオプションが不正なため「リセット」を使用: '" & outputOpt & "'", targetWorkbook, configStruct.ErrorLogSheetName, False, "WARNING_CONFIG_DEFAULT")
                 m_errorOccurred = False ' Ensure this specific warning is not fatal for LoadConfiguration
            End If
        End If
    ElseIf Len(configStruct.OutputDataOption) = 0 Then ' If error occurred during GetCellValue and option is not set
        configStruct.OutputDataOption = "リセット" ' Default
    End If


FinalConfigCheck: ' Label for GoTo statements if errors occur in C or E
    ' --- Final Check ---
    If m_errorOccurred Then
        If DEBUG_MODE_ERROR Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - ERROR: M02_ConfigReader.LoadConfiguration - One or more configuration errors occurred. See logs."
        MsgBox "Configシートの読み込み中に1つ以上のエラーが発生しました。詳細はエラーログを確認してください。", vbExclamation, "設定エラー"
        Exit Function ' Returns False
    End If

    ' --- Add Debug Print for Loaded Config Values ---
    If configStruct.DebugModeFlag Then
        Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_CONFIG: --- Loaded Configuration Settings (M02_ConfigReader) ---"
        ' A. 全般設定
        Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_CONFIG:   A-1. DebugModeFlag (O3): " & configStruct.DebugModeFlag
        Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_CONFIG:   A-Trace. TraceDebugEnabled (O4): " & configStruct.TraceDebugEnabled
        Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_CONFIG:   A-EnableSheetLogging (O5): " & configStruct.EnableSheetLogging
        Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_CONFIG:   A-LogSheetName (O42): '" & configStruct.LogSheetName & "'"
        Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_CONFIG:   A-2. DefaultFolderPath (O12): '" & configStruct.DefaultFolderPath & "'"
        Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_CONFIG:   A-3. OutputSheetName (O43): '" & configStruct.OutputSheetName & "'"
        Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_CONFIG:   A-4. SearchConditionLogSheetName (O44): '" & configStruct.SearchConditionLogSheetName & "'"
        Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_CONFIG:   A-5. ErrorLogSheetName (O45): '" & configStruct.ErrorLogSheetName & "'"
        Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_CONFIG:   A-6. ConfigSheetName (O46): '" & configStruct.ConfigSheetName & "'"
        Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_CONFIG:   A-7. GetPatternDataMethod (O122): " & configStruct.GetPatternDataMethod
        Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_CONFIG:   A-ExtractIfEmpty (O241): " & configStruct.ExtractIfWorkersEmpty
        Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_CONFIG:        ConfigSheetFullName: '" & configStruct.ConfigSheetFullName & "'"

        ' B. 工程表ファイル内 設定
        If ConfigReader_IsArrayInitialized(configStruct.TargetSheetNames) Then
            If UBound(configStruct.TargetSheetNames) >= LBound(configStruct.TargetSheetNames) Then ' Check if array actually has elements
                Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_CONFIG:   B-1. TargetSheetNames (O66-O75): " & Join(configStruct.TargetSheetNames, ", ")
            Else
                Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_CONFIG:   B-1. TargetSheetNames (O66-O75): (Array Initialized but Empty)"
            End If
        Else
            Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_CONFIG:   B-1. TargetSheetNames (O66-O75): (Not Initialized or Empty)"
        End If
        Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_CONFIG:   B-2. HeaderRowCount (O87): " & configStruct.HeaderRowCount
        Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_CONFIG:   B-3. HeaderColCount (O88): " & configStruct.HeaderColCount
        Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_CONFIG:   B-4. RowsPerDay (O89): " & configStruct.RowsPerDay
        Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_CONFIG:   B-5. MaxDaysPerSheet (O90): " & configStruct.MaxDaysPerSheet
        Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_CONFIG:   B-6. YearCellAddress (O101): '" & configStruct.YearCellAddress & "'"
        Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_CONFIG:   B-7. MonthCellAddress (O102): '" & configStruct.MonthCellAddress & "'"
        Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_CONFIG:   B-8. DayColumnLetter (O103): '" & configStruct.DayColumnLetter & "'"
        Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_CONFIG:   B-9. DayRowOffset (O104): " & configStruct.DayRowOffset
        Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_CONFIG:   B-10. ProcessesPerDay (O114): " & configStruct.ProcessesPerDay

        ' F. 抽出データオフセット定義
        Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_CONFIG:   F. Extraction Data Offsets (Loaded " & IIf(ConfigReader_IsArrayInitialized(configStruct.OffsetItemMasterNames), UBound(configStruct.OffsetItemMasterNames), 0) & " items):"
        If ConfigReader_IsArrayInitialized(configStruct.OffsetItemMasterNames) Then
            For fIdx = LBound(configStruct.OffsetItemMasterNames) To UBound(configStruct.OffsetItemMasterNames)
                Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_CONFIG:     " & fIdx & ". Name: '" & configStruct.OffsetItemMasterNames(fIdx) & _
                                  "', Offset: R=" & configStruct.OffsetDefinitions(fIdx).Row & ", C=" & configStruct.OffsetDefinitions(fIdx).Col & _
                                  ", IsEmptyOrig: " & configStruct.IsOffsetOriginallyEmptyFlags(fIdx)
            Next fIdx
        Else
            Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_CONFIG:     (No offset items loaded or array not initialized)"
        End If

        ' G. 出力シート設定
        Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_CONFIG:   G-1. OutputHeaderRowCount (O811): " & configStruct.OutputHeaderRowCount
        If ConfigReader_IsArrayInitialized(configStruct.OutputHeaderContents) Then
             Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_CONFIG:   G-2. OutputHeaderContents (O812-O821): " & Join(configStruct.OutputHeaderContents, " | ")
        Else
             Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_CONFIG:   G-2. OutputHeaderContents (O812-O821): (Not Initialized or Empty)"
        End If
        Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_CONFIG:   G-3. OutputDataOption (O1124): '" & configStruct.OutputDataOption & "'"

        Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_CONFIG: --- End of Loaded Configuration Settings ---"
    End If

    LoadConfiguration = True
    If configStruct.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M02_ConfigReader.LoadConfiguration - Configuration loaded successfully."
    Exit Function

LoadConfiguration_Error_MainHandler:
    Call ReportConfigError(m_errorOccurred, "LoadConfiguration", "N/A", "実行時エラー " & Err.Number & ": " & Err.Description, targetWorkbook, configStruct.ErrorLogSheetName, True, "FATAL_RUNTIME") ' configStruct.ErrorLogSheetName might be empty here
    LoadConfiguration = False
End Function
