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
    If DEBUG_MODE_DETAIL Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M02_ConfigReader.LoadStringList - Loading " & listDescription & " from " & columnLetter & firstRow & ":" & columnLetter & lastRow

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
    If DEBUG_MODE_DETAIL Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M02_ConfigReader.LoadStringList - " & listDescription & " loaded with " & tempCollection.Count & " items."
End Sub

Private Sub ReportConfigError(ByRef overallErrorFlag As Boolean, errorSourceContext As String, errorSourceCell As String, errorMessageText As String, _
                                ByVal wbForLog As Workbook, ByVal errorLogSheetNameForLog As String, Optional isFatal As Boolean = True)
    ' 設定読み込みエラーを報告し、必要に応じて全体エラーフラグを立てます。
    If isFatal Then overallErrorFlag = True
    
    Dim logSourceModule As String
    logSourceModule = "M02_ConfigReader." & errorSourceContext ' errorSourceContext will be like "LoadConfiguration (A-1)"

    If DEBUG_MODE_ERROR Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - CONFIG_ERROR" & IIf(isFatal, " (FATAL)", " (WARNING)") & ": " & logSourceModule & " - Cell: " & errorSourceCell & " - Message: " & errorMessageText
    
    Call SafeWriteErrorLog(wbForLog, errorLogSheetNameForLog, logSourceModule, errorSourceCell, errorMessageText, 0, "")
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

' --- Stubs for future implementation ---
Private Function ParseOffset(offsetString As String, ByRef resultOffset As tOffset) As Boolean
    ' TODO: Implement offset string parsing
    If DEBUG_MODE_WARNING Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - WARNING: M02_ConfigReader.ParseOffset - STUB called for: " & offsetString
    ParseOffset = True ' Placeholder
    Exit Function
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
    If DEBUG_MODE_DETAIL Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M02_ConfigReader.LoadConfiguration - Attempting to get Config sheet: '" & configSheetName & "'"
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
    If DEBUG_MODE_DETAIL Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M02_ConfigReader.LoadConfiguration - Config sheet found: '" & configStruct.ConfigSheetFullName & "'"

    ' --- A. 全般設定 ---
    If DEBUG_MODE_DETAIL Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M02_ConfigReader.LoadConfiguration - Reading Section A: General Settings"
    
    ' A-1: デバッグモードフラグ (O3)
    tempVal = GetCellValue(wsConfig, "O3", "LoadConfiguration (A-1)", m_errorOccurred, "デバッグモードフラグ", targetWorkbook, configStruct.ErrorLogSheetName, False, "Boolean")
    If Not m_errorOccurred Then ' Check if GetCellValue itself caused a fatal error
        If Not IsEmpty(tempVal) Then
            configStruct.DebugModeFlag = tempVal
        Else ' Value was empty or not a valid boolean string (but not a read error from GetCellValue)
            configStruct.DebugModeFlag = False ' Default value
            If DEBUG_MODE_WARNING Then Call ReportConfigError(m_errorOccurred, "LoadConfiguration (A-1)", "O3", "デバッグモードフラグ(O3)が不正または空のためFalseを適用", targetWorkbook, configStruct.ErrorLogSheetName, False) ' Not fatal
        End If
    End If
    ' Note: m_errorOccurred could be true if GetCellValue had a fatal read error.

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
            If DEBUG_MODE_WARNING Then Call ReportConfigError(m_errorOccurred, "LoadConfiguration (A-7)", "O122", "工程パターンデータ取得方法(O122)が不正または空のためFalse(VBA方式)を適用", targetWorkbook, configStruct.ErrorLogSheetName, False) ' Not fatal
        End If
    End If

    ' --- B. 工程表ファイル内 設定 ---
    If DEBUG_MODE_DETAIL Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M02_ConfigReader.LoadConfiguration - Reading Section B: Schedule File Settings"
    
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

    ' --- Final Check ---
    If m_errorOccurred Then
        If DEBUG_MODE_ERROR Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - ERROR: M02_ConfigReader.LoadConfiguration - One or more configuration errors occurred. See logs."
        MsgBox "Configシートの読み込み中に1つ以上のエラーが発生しました。詳細はエラーログを確認してください。", vbExclamation, "設定エラー"
        Exit Function ' Returns False
    End If

    LoadConfiguration = True
    If DEBUG_MODE_DETAIL Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M02_ConfigReader.LoadConfiguration - Configuration loaded successfully."
    Exit Function

LoadConfiguration_Error_MainHandler:
    Call ReportConfigError(m_errorOccurred, "LoadConfiguration", "N/A", "実行時エラー " & Err.Number & ": " & Err.Description, targetWorkbook, configStruct.ErrorLogSheetName) ' configStruct.ErrorLogSheetName might be empty here
    LoadConfiguration = False
End Function
