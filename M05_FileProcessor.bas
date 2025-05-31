' バージョン：v0.5.1
Option Explicit
' このモジュールは、処理対象となるファイルパスの特定と管理を行います。
' 設定に基づいてフォルダを探索し、対象となるExcelファイルのリストを作成します。

Private Const MODULE_NAME As String = "M05_FileProcessor"

' Public Function: GetTargetFiles
' 設定情報に基づき、処理対象となる全てのExcelファイルのフルパスと、
' 対応する工程パターン識別子を収集します。
' Parameters:
'   ByRef config As tConfigSettings - 設定情報
'   ByVal mainWorkbook As Workbook - メインワークブック (ログ記録用)
'   ByRef foundFilesDetails As Collection - (Output) 見つかったファイル詳細(PathとPatternIDを持つDictionary)を格納するコレクション
' Returns: Boolean - 処理が正常に完了した場合はTrue、エラー発生時はFalse
Public Function GetTargetFiles(ByRef config As tConfigSettings, ByVal mainWorkbook As Workbook, ByRef foundFilesDetails As Collection) As Boolean
    Dim funcName As String: funcName = "GetTargetFiles"
    Dim fso As Object ' FileSystemObject
    Dim i As Long
    Dim pathItem As String
    Dim patternIdItem As String
    Dim fileItem As Object ' Scripting.File
    Dim folderItem As Object ' Scripting.Folder
    Dim fileDetail As Object ' Scripting.Dictionary to store path and patternID

    On Error GoTo ErrorHandler

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set foundFilesDetails = New Collection ' 出力用コレクションを初期化

    If Not General_IsArrayInitialized(config.TargetFileFolderPaths) Then
        Call M04_LogWriter.WriteErrorLog("INFORMATION", MODULE_NAME, funcName, "処理対象ファイル/フォルダパスリスト(TargetFileFolderPaths)が空または未初期化です。")
        GetTargetFiles = True ' 処理は成功（ファイルが見つからなかっただけ）
        Exit Function
    End If

    For i = LBound(config.TargetFileFolderPaths) To UBound(config.TargetFileFolderPaths)
        pathItem = Trim(CStr(config.TargetFileFolderPaths(i)))

        ' 対応する工程パターン識別子を取得
        If General_IsArrayInitialized(config.FilePatternIdentifiers) And _
           i >= LBound(config.FilePatternIdentifiers) And _
           i <= UBound(config.FilePatternIdentifiers) Then
            patternIdItem = Trim(CStr(config.FilePatternIdentifiers(i)))
        Else
            patternIdItem = "" ' デフォルトまたはエラーケース
            Call M04_LogWriter.WriteErrorLog("WARNING", MODULE_NAME, funcName, "工程パターン識別子リストのインデックス" & i & "に対応する値がありません。パス「" & pathItem & "」には空のパターンIDが使用されます。")
        End If

        If pathItem = "" Then
            Call M04_LogWriter.WriteErrorLog("WARNING", MODULE_NAME, funcName, "処理対象パスが空です。スキップします。(インデックス: " & i & ")")
            GoTo NextPath
        End If

        If fso.FolderExists(pathItem) Then
            Set folderItem = fso.GetFolder(pathItem)
            If folderItem.Files.Count = 0 Then
                 Call M04_LogWriter.WriteErrorLog("INFORMATION", MODULE_NAME, funcName, "対象フォルダ「" & pathItem & "」にファイルが存在しません。")
                 GoTo NextPath
            End If
            For Each fileItem In folderItem.Files
                If IsExcelFile(fileItem.Name, fso) Then
                    Set fileDetail = CreateObject("Scripting.Dictionary")
                    fileDetail("Path") = fileItem.Path
                    fileDetail("PatternID") = patternIdItem
                    foundFilesDetails.Add fileDetail
                End If
            Next fileItem
        ElseIf fso.FileExists(pathItem) Then
            If IsExcelFile(pathItem, fso) Then
                Set fileDetail = CreateObject("Scripting.Dictionary")
                fileDetail("Path") = pathItem
                fileDetail("PatternID") = patternIdItem
                foundFilesDetails.Add fileDetail
            Else
                Call M04_LogWriter.WriteErrorLog("WARNING", MODULE_NAME, funcName, "指定されたファイル「" & pathItem & "」はExcelファイルではありません。スキップします。")
            End If
        Else
            Call M04_LogWriter.WriteErrorLog("ERROR", MODULE_NAME, funcName, "指定されたパス「" & pathItem & "」が見つかりません。")
            ' GetTargetFiles = False ' 処理を中断する場合はFalseにする
        End If
NextPath:
    ' Erroneous "Loop" keyword removed from here. The For...Next i handles the loop continuation.
    Next i


    If foundFilesDetails.Count = 0 Then
        Call M04_LogWriter.WriteErrorLog("INFORMATION", MODULE_NAME, funcName, "処理対象となるExcelファイルは見つかりませんでした。")
    Else
        Call M04_LogWriter.WriteErrorLog("INFORMATION", MODULE_NAME, funcName, foundFilesDetails.Count & "個の処理対象Excelファイル(と対応パターンID)が見つかりました。")
    End If

    GetTargetFiles = True ' 処理完了
    GoTo CleanupAndExit

ErrorHandler:
    Call M04_LogWriter.WriteErrorLog("CRITICAL", MODULE_NAME, funcName, "ファイル特定処理中に予期せぬエラーが発生しました。", Err.Number, Err.Description)
    GetTargetFiles = False

CleanupAndExit:
    Set fso = Nothing
    Set fileItem = Nothing
    Set folderItem = Nothing
    Set fileDetail = Nothing
End Function

' Private Helper Function: IsExcelFile
Private Function IsExcelFile(ByVal fileNameOrPath As String, ByVal fso As Object) As Boolean
    Dim extension As String
    On Error Resume Next
    extension = fso.GetExtensionName(fileNameOrPath)
    If Err.Number <> 0 Then
        IsExcelFile = False
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0

    Select Case LCase(extension)
        Case "xls", "xlsx", "xlsm", "xlsb"
            IsExcelFile = True
        Case Else
            IsExcelFile = False
    End Select
End Function

Public Function General_IsArrayInitialized(arr As Variant) As Boolean
    If Not IsArray(arr) Then
        General_IsArrayInitialized = False
        Exit Function
    End If

    ' 配列であれば、ReDimされているとみなし、初期化済みとする
    ' LBoundやUBoundのチェックは、要素が存在するかどうかの判断であり、
    ' 配列が「初期化されているか（DimやReDimされたか）」の判断とは異なる場合がある。
    ' 特にユーザー定義型の配列の場合、LBound等がエラーになることがあるため、
    ' IsArray(arr) が True であれば、ここでは初期化済みと判断する。
    General_IsArrayInitialized = True

    ' もし「要素が実際に存在するか」を確認したい場合は、別途 UBound(arr) >= LBound(arr) のようなチェックを行う。
    ' ここでは「配列として使える状態か」を返すことに注力する。
End Function
