' バージョン：v0.5.0
Option Explicit
' このモジュールは、処理対象となる工程表ファイルのリストを取得・管理します。このステップでは、Configシートの特定セルから単一のファイルパスを取得する簡易版を実装します。

Private Function LogFileProcessor_IsArrayInitialized(arr As Variant) As Boolean
    ' 配列が有効に初期化されているか（少なくとも1つの要素を持つか）を確認します。
    ' Variant型が配列でない場合、または配列であっても要素が割り当てられていない場合（Dim arr() のみでReDimされていない状態など）はFalseを返します。
    On Error GoTo NotAnArrayOrNotInitialized_FP
    If IsArray(arr) Then
        Dim lBoundCheck As Long
        lBoundCheck = LBound(arr)
        LogFileProcessor_IsArrayInitialized = True
        Exit Function
    End If
NotAnArrayOrNotInitialized_FP:
    LogFileProcessor_IsArrayInitialized = False
End Function

Private Function IsExcelFile(ByVal fileName As String) As Boolean
    ' 指定されたファイル名が一般的なExcelファイルの拡張子を持つかどうかを判定します。
    Dim ext As String
    If InStrRev(fileName, ".") = 0 Then ' No extension
        IsExcelFile = False
        Exit Function
    End If
    ext = LCase(Right(fileName, Len(fileName) - InStrRev(fileName, ".")))
    
    Select Case ext
        Case "xlsx", "xls", "xlsm"
            IsExcelFile = True
        Case Else
            IsExcelFile = False
    End Select
End Function

Public Function GetTargetFiles(ByRef config As tConfigSettings, ByVal mainAppWorkbook As Workbook, ByRef targetFilesCollection As Collection) As Boolean
    ' Config設定 (P557) から単一の処理対象ファイルパスを取得し、検証後、コレクションに追加します。(ステップ4限定仕様)
    ' Arguments:
    '   config: (I) tConfigSettings型。設定情報を保持します。ErrorLogSheetName と TargetFileFolderPaths(0) を使用します。
    '   mainAppWorkbook: (I) Workbook型。マクロ本体（ログシートが存在する）のワークブックオブジェクト。
    '   targetFilesCollection: (O) Collection型。検証済みのファイルパスが追加されます。
    ' Returns:
    '   Boolean: ファイルパスの取得と検証に成功し、コレクションに追加できた場合はTrue、それ以外はFalse。

    Dim fso As Object
    Dim filePathFromConfig As String
    Dim successFlag As Boolean ' Local flag for clarity before setting function return
    
    successFlag = False ' Default to failure
    On Error GoTo GetTargetFiles_Error

    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 設定構造体からファイルパスを取得 (ステップ4では config.TargetFileFolderPaths(0) にP557の値が入っている想定)
    If LogFileProcessor_IsArrayInitialized(config.TargetFileFolderPaths) Then
        If UBound(config.TargetFileFolderPaths) >= LBound(config.TargetFileFolderPaths) Then
            filePathFromConfig = config.TargetFileFolderPaths(LBound(config.TargetFileFolderPaths))
        Else
            filePathFromConfig = "" ' Array is initialized but empty
        End If
    Else
        filePathFromConfig = "" ' Array not initialized
    End If

    If Len(filePathFromConfig) > 0 Then
        If fso.FileExists(filePathFromConfig) Then
            If IsExcelFile(filePathFromConfig) Then
                targetFilesCollection.Add filePathFromConfig
                successFlag = True
                If config.TraceDebugEnabled Then Debug.Print Format(Now, "yyyy/mm/dd hh:nn:ss") & " - DEBUG_DETAIL: M05_FileProcessor.GetTargetFiles - Added target file: '" & filePathFromConfig & "'"
            Else
                Call M04_LogWriter.SafeWriteErrorLog("WARNING", mainAppWorkbook, config.ErrorLogSheetName, "M05_FileProcessor", "GetTargetFiles", "指定されたファイルはExcelファイルではありません: " & filePathFromConfig, 0, "")
            End If
        Else
            Call M04_LogWriter.SafeWriteErrorLog("ERROR", mainAppWorkbook, config.ErrorLogSheetName, "M05_FileProcessor", "GetTargetFiles", "指定されたファイルが見つかりません: " & filePathFromConfig, 0, "")
        End If
    Else
        Call M04_LogWriter.SafeWriteErrorLog("WARNING", mainAppWorkbook, config.ErrorLogSheetName, "M05_FileProcessor", "GetTargetFiles", "ConfigシートP557に処理対象ファイルパスが指定されていません。", 0, "")
    End If
    
    GetTargetFiles = successFlag
    Set fso = Nothing
    Exit Function

GetTargetFiles_Error:
    Call M04_LogWriter.SafeWriteErrorLog("ERROR", mainAppWorkbook, config.ErrorLogSheetName, "M05_FileProcessor", "GetTargetFiles", "実行時エラー " & Err.Number & ": " & Err.Description, Err.Number, Err.Description)
    GetTargetFiles = False
    Set fso = Nothing
End Function
