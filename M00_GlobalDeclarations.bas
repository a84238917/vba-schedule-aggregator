Option Explicit
' このモジュールは、プロジェクト全体で共有されるグローバル定数、Publicなユーザー定義型、およびPublicなグローバル変数を一元的に宣言・管理します。

' Global Debug Flag Constants
Public Const DEBUG_MODE_ERROR As Boolean = True    ' エラー関連の詳細デバッグ情報を出力するかどうか
Public Const DEBUG_MODE_WARNING As Boolean = False ' 警告レベルのデバッグ情報を出力するかどうか
Public Const DEBUG_MODE_DETAIL As Boolean = False  ' 詳細な処理追跡情報を出力するかどうか

' Fixed Setting Value Constant
Public Const CONFIG_SHEET_DEFAULT_NAME As String = "Config (2)" ' 設定シートのデフォルト名

' User-Defined Type: tOffset
Public Type tOffset
    Row As Long ' 行オフセット
    Col As Long ' 列オフセット
End Type

' User-Defined Type: tProcessDetail
Public Type tProcessDetail
    Kankatsu1 As String ' 管轄1情報
    Kankatsu2 As String ' 管轄2情報
End Type

' User-Defined Type: tConfigSettings
Public Type tConfigSettings
    ' A. General Settings
    DebugModeFlag As Boolean              ' O3 デバッグモードフラグ
    DefaultFolderPath As String           ' O12 デフォルトフォルダパス
    OutputSheetName As String             ' O43 抽出結果出力シート名
    SearchConditionLogSheetName As String ' O44 検索条件ログシート名
    ErrorLogSheetName As String           ' O45 エラーログシート名
    ConfigSheetName As String             ' O46 設定ファイルシート名
    GetPatternDataMethod As Boolean       ' O122 工程パターンデータ取得方法 (TRUE:数式, FALSE:VBA)

    ' B. Schedule File Settings
    TargetSheetNames() As String    ' O66-O75 工程表内 検索対象シート名リスト
    HeaderRowCount As Long          ' O87 工程表ヘッダー行数
    HeaderColCount As Long          ' O88 工程表ヘッダー列数
    RowsPerDay As Long              ' O89 1日のデータが占める行数
    MaxDaysPerSheet As Long         ' O90 1シート内の最大日数
    YearCellAddress As String       ' O101 「年」のセルアドレス
    MonthCellAddress As String      ' O102 「月」のセルアドレス
    DayColumnLetter As String       ' O103 「日」の値がある列文字
    DayRowOffset As Long            ' O104 「日」の値の行オフセット
    ProcessesPerDay As Long         ' O114 1日の工程数

    ' C. Process Pattern Definition
    CurrentPatternIdentifier As String       ' O126 現在処理中ファイル適用工程パターン識別子
    ProcessKeys() As String                  ' I129-I(128+O114) 工程キーリスト
    Kankatsu1List() As String                ' J129-J(128+O114) 管内1リスト
    Kankatsu2List() As String                ' K129-K(128+O114) 管内2リスト
    Bunrui1List() As String                  ' L129-L(128+O114) 分類1リスト
    Bunrui2List() As String                  ' M129-M(128+O114) 分類2リスト
    Bunrui3List() As String                  ' N129-N(128+O114) 分類3リスト
    ProcessColCountSheetHeaders() As String  ' O128-X128 工程列数定義用シート名ヘッダー
    ProcessColCounts() As Variant            ' O129-X(128+O114) 工程パターン別 工程列数定義
    ProcessDetails() As tProcessDetail       ' C-3, C-4から派生 各工程の管轄情報
    ProcessPatternColNumbers() As Variant    ' C-9から派生 現在のシートに対応する工程列数

    ' D. Filter Conditions
    WorkerFilterLogic As String         ' O242 作業員フィルター検索論理
    WorkerFilterList() As String        ' O243-O262 作業員フィルターリスト
    Kankatsu1FilterList() As String     ' O275-O294 管内1フィルターリスト
    Kankatsu2FilterList() As String     ' O305-O334 管内2フィルターリスト
    Bunrui1Filter As String             ' O346 分類1フィルター
    Bunrui2Filter As String             ' O367 分類2フィルター
    Bunrui3Filter As String             ' O388 分類3フィルター
    KoujiShuruiFilterList() As String   ' O409-O418 工事種類フィルターリスト
    KoubanFilterList() As String        ' O431-O440 工番フィルターリスト
    SagyoushuruiFilterList() As String  ' O451-O470 作業種類フィルターリスト
    TantouFilterList() As String        ' O481-O490 担当の名前フィルターリスト
    NinzuFilter As String               ' O503 人数フィルター (数値だが文字列で読み込み空を判定)
    IsNinzuFilterOriginallyEmpty As Boolean ' O503 人数フィルターが元々空だったか
    SagyouKashoKindFilter As String     ' O514 作業箇所の種類フィルター
    SagyouKashoFilterList() As String   ' O525-O544 作業箇所フィルターリスト

    ' E. Target File Definition
    TargetFileFolderPaths() As String ' P557-P756 処理対象ファイル/フォルダパスリスト
    FilePatternIdentifiers() As String ' Q557-Q756 各処理対象ファイル適用工程パターン識別子

    ' F. Extraction Data Offset Definition
    OffsetItemNames() As String             ' N778-N792 オフセット項目名リスト
    OffsetValuesRaw() As String             ' O778-O792 オフセット値リスト (Raw "row,col" strings)
    Offsets() As tOffset                    ' F-2 オフセット値 (パース後)
    IsOffsetKoubanOriginallyEmpty As Boolean ' F-1 工番オフセットが元々空だったか
    IsOffsetHensendenjoOriginallyEmpty As Boolean ' F-1 変電所オフセットが元々空だったか
    IsOffsetSagyomei1OriginallyEmpty As Boolean ' F-1 作業名1オフセットが元々空だったか
    IsOffsetSagyomei2OriginallyEmpty As Boolean ' F-1 作業名2オフセットが元々空だったか
    IsOffsetTantouOriginallyEmpty As Boolean ' F-1 担当の名前オフセットが元々空だったか
    IsOffsetKoujiShuruiOriginallyEmpty As Boolean ' F-1 工事種類オフセットが元々空だったか
    IsOffsetNinzuOriginallyEmpty As Boolean ' F-1 人数オフセットが元々空だったか
    IsOffsetSagyoinOriginallyEmpty As Boolean ' F-1 作業員オフセットが元々空だったか
    IsOffsetSonotaOriginallyEmpty As Boolean ' F-1 旧その他オフセットが元々空だったか
    IsOffsetShuuryoJikanOriginallyEmpty As Boolean ' F-1 終了時間オフセットが元々空だったか
    IsOffsetBunrui1ExtSrcOriginallyEmpty As Boolean ' F-1 分類1抽出元オフセットが元々空だったか

    ' G. Output Sheet Settings
    OutputHeaderRowCount As Long    ' O811 出力シートヘッダー行数
    OutputHeaderContents() As String ' O812-O821 出力シートヘッダー内容
    OutputDataOption As String      ' O1124 出力データオプション
    HideSheetMethod As String       ' O1126 非表示方式
    HideSheetNames() As String      ' O1127-O1146 マクロ実行後非表示シートリスト

    ' Additional Members
    StartTime As Date               ' マクロ実行開始時刻
    ScriptFullName As String        ' マクロファイルのフルパス
    WorkSheetName As String         ' Workシート名 (固定値または設定による)
    ConfigSheetFullName As String   ' Configシートのフルネーム (Workbook名を含む)
End Type

' Global Variables
Public g_configSettings As tConfigSettings   ' Configシートから読み込まれた全ての設定情報を格納するグローバル変数
Public g_errorLogWorksheet As Worksheet      ' エラーログを書き込むワークシートオブジェクト
Public g_nextErrorLogRow As Long             ' エラーログシートの次に書き込む行番号
