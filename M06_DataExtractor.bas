' バージョン：v0.5.1
Option Explicit
' このモジュールは、個別の工程表ファイルからデータを抽出する役割を担います。
' 設定に基づいてファイルを開き、指定されたオフセットに従って情報を読み取ります。

Private Const MODULE_NAME As String = "M06_DataExtractor"

' Public Function: ExtractDataFromFile
' 指定されたExcelファイルからデータを抽出します。
' Parameters:
'   ByVal targetFilePath As String - 処理対象のExcelファイルパス
'   ByRef config As tConfigSettings - 設定情報
'   ByVal mainWorkbook As Workbook - マクロ本体のワークブック (ログ記録用)
' Returns: Boolean - 抽出成功/失敗 (v0.1では主に処理完了を示す)
Public Function ExtractDataFromFile(ByVal targetFilePath As String, ByRef config As tConfigSettings, ByVal mainWorkbook As Workbook) As Boolean
    Dim funcName As String: funcName = "ExtractDataFromFile"
    Dim wbTarget As Workbook
    Dim wsTarget As Worksheet
    Dim sheetName As Variant ' ループ用
    Dim i As Long, dayLoopIndex As Long, procLoopIndex As Long ' ループカウンタ
    Dim yearVal As Long, monthVal As Long, dayValAsLong As Long ' dayValをLong型に
    Dim dataRow As Long, dataCol As Long
    Dim extractedValue As String
    Dim outputLine As String
    Dim dayCellVal As Variant ' 日付セルからの読み取り用

    On Error GoTo ErrorHandler

    If Dir(targetFilePath) = "" Then
        Call M04_LogWriter.WriteErrorLog("ERROR", MODULE_NAME, funcName, "指定されたファイルが見つかりません: " & targetFilePath)
        ExtractDataFromFile = False
        Exit Function
    End If

    Call M04_LogWriter.WriteErrorLog("INFORMATION", MODULE_NAME, funcName, "処理開始: " & targetFilePath & " (パターンID: " & config.CurrentPatternIdentifier & ")")

    On Error Resume Next
    Set wbTarget = Workbooks.Open(Filename:=targetFilePath, ReadOnly:=True, UpdateLinks:=0)
    If Err.Number <> 0 Then
        Call M04_LogWriter.WriteErrorLog("ERROR", MODULE_NAME, funcName, "ファイルを開けませんでした: " & targetFilePath & ". エラー: " & Err.Description, Err.Number)
        ExtractDataFromFile = False
        Err.Clear
        Exit Function
    End If
    On Error GoTo ErrorHandler

    If Not General_IsArrayInitialized(config.TargetSheetNames) Then
        Call M04_LogWriter.WriteErrorLog("WARNING", MODULE_NAME, funcName, "対象シート名リストが設定されていません。ファイル処理をスキップ: " & targetFilePath)
        GoTo CleanupAndExit
    End If

    For Each sheetName In config.TargetSheetNames
        If Trim(CStr(sheetName)) = "" Then GoTo NextSheetLoop

        On Error Resume Next
        Set wsTarget = wbTarget.Sheets(CStr(sheetName))
        If Err.Number <> 0 Then
            Call M04_LogWriter.WriteErrorLog("WARNING", MODULE_NAME, funcName, "シート「" & sheetName & "」がファイル「" & targetFilePath & "」内に見つかりません。スキップします。", Err.Number)
            Err.Clear
            Set wsTarget = Nothing
            GoTo NextSheetLoop
        End If
        On Error GoTo ErrorHandler

        Call M04_LogWriter.WriteErrorLog("INFORMATION", MODULE_NAME, funcName, "シート処理中: " & wsTarget.Name)

        ' 年月取得 (Config設定から)
        On Error Resume Next
        yearVal = CLng(wsTarget.Range(config.YearCellAddress).Value)
        monthVal = CLng(wsTarget.Range(config.MonthCellAddress).Value)
        If Err.Number <> 0 Or yearVal = 0 Or monthVal = 0 Then
             Call M04_LogWriter.WriteErrorLog("WARNING", MODULE_NAME, funcName, "シート「" & wsTarget.Name & "」から年または月を取得できませんでした。YearCell: " & config.YearCellAddress & ", MonthCell: " & config.MonthCellAddress & ". スキップします。", Err.Number)
             Err.Clear
             GoTo NextSheetLoop
        End If
        On Error GoTo ErrorHandler

        For dayLoopIndex = 1 To config.MaxDaysPerSheet
            dayCellVal = wsTarget.Cells(config.HeaderRowCount + ((dayLoopIndex - 1) * config.RowsPerDay) + config.DayRowOffset, config.DayColumnLetter).Value
            If Not IsNumeric(dayCellVal) Or Trim(CStr(dayCellVal)) = "" Then GoTo NextDayLoop
            dayValAsLong = CLng(dayCellVal)
            If dayValAsLong <= 0 Or dayValAsLong > 31 Then GoTo NextDayLoop

            For procLoopIndex = 1 To config.ProcessesPerDay
                outputLine = Format(Now(), "yyyy-mm-dd hh:nn:ss") & " - " & _
                             targetFilePath & " | " & wsTarget.Name & " | " & _
                             yearVal & "年" & monthVal & "月" & dayValAsLong & "日" & " | " & _
                             "工程" & procLoopIndex & ":"

                If General_IsArrayInitialized(config.OffsetItemMasterNames) And General_IsArrayInitialized(config.OffsetDefinitions) Then
                    If UBound(config.OffsetItemMasterNames) = UBound(config.OffsetDefinitions) And _
                       LBound(config.OffsetItemMasterNames) = LBound(config.OffsetDefinitions) Then
                        For i = LBound(config.OffsetItemMasterNames) To UBound(config.OffsetItemMasterNames)
                            Dim currentItemName As String
                            Dim currentOffset As tOffset
                            currentItemName = config.OffsetItemMasterNames(i)
                            currentOffset = config.OffsetDefinitions(i)

                            ' データ読み取り位置計算 (v0.1 シンプル版)
                            ' プロセスごとの列オフセットはCurrentPatternIdentifierとProcessPatternColNumbersから取得する想定だがv0.1では未実装
                            ' ここではdayLoopIndexとprocLoopIndexに基づく行オフセットと、config.Offsetsの列オフセットのみ考慮
                            dataRow = config.HeaderRowCount + ((dayLoopIndex - 1) * config.RowsPerDay) + config.DayRowOffset + currentOffset.Row
                                      '+ procLoopIndex -1 ; Process毎のベース行が異なる場合、この行調整も必要
                            dataCol = Columns(config.DayColumnLetter).Column + currentOffset.Col ' 日付列を基準とした列オフセット

                            If dataRow > 0 And dataCol > 0 And dataRow <= wsTarget.Rows.Count And dataCol <= wsTarget.Columns.Count Then
                                extractedValue = Trim(CStr(wsTarget.Cells(dataRow, dataCol).Value))
                                outputLine = outputLine & " [" & currentItemName & " (" & currentOffset.Row & "," & currentOffset.Col & "): '" & extractedValue & "']"
                            Else
                                outputLine = outputLine & " [" & currentItemName & " (" & currentOffset.Row & "," & currentOffset.Col & "): (範囲外 R:" & dataRow & ",C:" & dataCol & ")]"
                            End If
                        Next i
                    Else
                        outputLine = outputLine & " [エラー: オフセット定義の配列数/範囲が不一致です]"
                    End If
                Else
                    outputLine = outputLine & " [エラー: オフセット定義が初期化されていません]"
                End If
                Debug.Print outputLine
            Next procLoopIndex
NextDayLoop:
        Next dayLoopIndex
NextSheetLoop:
    Next sheetName

CleanupAndExit:
    If Not wbTarget Is Nothing Then
        wbTarget.Close SaveChanges:=False
    End If
    Set wbTarget = Nothing
    Set wsTarget = Nothing
    Call M04_LogWriter.WriteErrorLog("INFORMATION", MODULE_NAME, funcName, "処理終了: " & targetFilePath)
    ExtractDataFromFile = True
    Exit Function

ErrorHandler:
    Call M04_LogWriter.WriteErrorLog("CRITICAL", MODULE_NAME, funcName, "データ抽出処理中に予期せぬエラー。ファイル: " & targetFilePath, Err.Number, Err.Description)
    ExtractDataFromFile = False
    If Not wbTarget Is Nothing Then
        wbTarget.Close SaveChanges:=False
    End If
    Set wbTarget = Nothing
    Set wsTarget = Nothing
End Function

Public Function General_IsArrayInitialized(arr As Variant) As Boolean
    If Not IsArray(arr) Then Exit Function
    On Error Resume Next
    Dim lBoundCheck As Long: lBoundCheck = LBound(arr)
    If Err.Number = 0 Then General_IsArrayInitialized = True
    On Error GoTo 0
End Function
