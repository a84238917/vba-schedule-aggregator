Option Explicit
' このモジュールは、「Config」シートから全ての設定情報を読み込み、検証し、g_configSettingsグローバル変数に格納する役割を担います。

Public Function LoadConfiguration(ByRef configSettings As tConfigSettings, ByVal configSheet As Worksheet) As Boolean
    ' 「Config」シートから基本的な設定（当面はログシート名のみ）を読み込み、引数のconfigSettings構造体に格納します。
    ' 今後、全ての設定項目の読み込みと検証をこの関数で実施する予定です。
    ' 引数:
    '   configSettings: (I/O) tConfigSettings型。読み込んだ設定が格納されます。
    '   configSheet: (I) Worksheet型。読み込み元のConfigシートオブジェクト。
    ' 戻り値:
    '   Boolean: 必須設定項目（ログシート名）の読み込みに成功した場合はTrue、それ以外はFalse。

    LoadConfiguration = False ' Default to failure
    On Error GoTo LoadConfiguration_Error ' Basic error handling for unexpected issues
    
    Dim errSheetNameVal As Variant
    Dim logSheetNameVal As Variant
    
    errSheetNameVal = configSheet.Range("O45").Value
    logSheetNameVal = configSheet.Range("O44").Value

    If DEBUG_MODE_DETAIL Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M02_ConfigReader.LoadConfiguration - Attempting to read O45 (ErrorLogSheetName), O44 (SearchConditionLogSheetName)"
    If DEBUG_MODE_DETAIL Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M02_ConfigReader.LoadConfiguration - Raw O45 value: '" & CStr(errSheetNameVal) & "'"
    If DEBUG_MODE_DETAIL Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M02_ConfigReader.LoadConfiguration - Raw O44 value: '" & CStr(logSheetNameVal) & "'"

    ' Validate ErrorLogSheetName (O45)
    If Not IsError(errSheetNameVal) And Len(Trim(CStr(errSheetNameVal))) > 0 Then
        configSettings.ErrorLogSheetName = Trim(CStr(errSheetNameVal))
    Else
        If DEBUG_MODE_ERROR Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - ERROR: M02_ConfigReader.LoadConfiguration - ErrorLogSheetName (O45) is empty or invalid."
        Exit Function ' Returns False
    End If

    ' Validate SearchConditionLogSheetName (O44)
    If Not IsError(logSheetNameVal) And Len(Trim(CStr(logSheetNameVal))) > 0 Then
        configSettings.SearchConditionLogSheetName = Trim(CStr(logSheetNameVal))
    Else
        If DEBUG_MODE_ERROR Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - ERROR: M02_ConfigReader.LoadConfiguration - SearchConditionLogSheetName (O44) is empty or invalid."
        Exit Function ' Returns False
    End If

    If DEBUG_MODE_DETAIL Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M02_ConfigReader.LoadConfiguration - Successfully read ErrorLogSheetName: '" & configSettings.ErrorLogSheetName & "', SearchConditionLogSheetName: '" & configSettings.SearchConditionLogSheetName & "'"
    LoadConfiguration = True ' All essential names read successfully
    Exit Function

LoadConfiguration_Error:
    If DEBUG_MODE_ERROR Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - ERROR: M02_ConfigReader.LoadConfiguration - Runtime error " & Err.Number & ": " & Err.Description
    LoadConfiguration = False ' Ensure False on error
    Err.Clear
End Function
