このステップでは、マクロの動作を制御する上で最も基本的な設定項目（全般設定、シート名など）を「Config」シートから読み込み、グローバル設定構造体に格納する機能を実装します。基本的な入力値の検証も行います。

---

## プロンプト (AI: Jules向け) - ステップ3: Configシート読み込み機能の基盤実装

**依頼プロジェクト名:** VBA Schedule Aggregator (工程表データ集約マクロ) - 再構築 (ステップ3)

**現在の開発ステップ:** 【ステップ3】Configシート読み込み機能の基盤実装

**このステップの目的:**
VBAマクロの動作を決定づける中核となる「Config」シートから、基本的な設定情報を読み取り、グローバル設定構造体（`g_configSettings`）に格納する機能を実装します。具体的には、全般設定（デバッグモード、各種シート名、デフォルトフォルダパスなど）と、工程表ファイル内の基本的な構造に関する設定（検索対象シート名リスト、ヘッダー行数・列数、1日の行数・日数、年月日の取得セルなど）の読み込みと、簡単な入力値検証を行います。このステップは、マクロがユーザーの設定に基づいて動作するための最初の重要なステップです。

**参照必須ドキュメント:**
1.  `docs/00_Project_Overview.md` (Configシートの役割)
2.  `docs/01_System_Instructions_for_AI.md` (特に、原則1「設定駆動」、原則2「可読性」、原則3「堅牢性」、原則5「モジュール性」、原則7「禁止事項」を厳守)
3.  `docs/03_Functional_Specification.md` (セクション3.1「設定管理機能」、3.2「Configシート読み込み」)
4.  `docs/04_Config_Sheet_Definition.md` (特に、**Aセクション「全般設定」** および **Bセクション「工程表ファイル内 設定」** の全項目定義)
5.  `docs/07_Naming_Conventions_and_Glossary.md` (特に、接頭辞 `g_`, `t`, `ws`, モジュール名 `M02_` の規約、および `tConfigSettings` のメンバー構成)

**前提条件:**
*   「【ステップ1】」で作成された `M00_GlobalDeclarations` モジュール（特に `Public Const CONFIG_SHEET_DEFAULT_NAME`, `Public Type tConfigSettings`, `Public g_configSettings As tConfigSettings`）が存在すること。
*   「【ステップ2】」で作成された `M04_LogWriter` モジュール（特に `SafeWriteErrorLog` プロシージャ）が利用可能であること（エラー発生時のログ記録のため）。

**あなたのタスク:**
以下の仕様に基づき、`M02_ConfigReader` モジュールを新規作成し、`M01_MainControl` モジュールを更新するVBAコードを生成してください。

**1. `M02_ConfigReader` モジュール (新規作成):**

*   **目的:** 「Config」シートから全ての設定情報を読み込み、検証し、グローバル設定構造体 `g_configSettings` に格納する処理を専門に担当します。
*   **実装要件:**
    *   モジュールの先頭に `Option Explicit` を記述してください。
    *   モジュールヘッダーコメントとして、このモジュールの役割を日本語で記述してください。
    *   **`Public Function LoadConfiguration(ByRef configStruct As tConfigSettings, ByVal targetWorkbook As Workbook, ByVal configSheetName As String) As Boolean` プロシージャ:**
        *   プロシージャヘッダーコメント（目的、引数、戻り値の説明）を記述。
        *   ローカル変数（`wsConfig As Worksheet`, `errorOccurred As Boolean` など）を宣言し、コメント付与。
        *   `On Error GoTo LoadConfiguration_Error` でエラーハンドラを設定。
        *   引数 `configSheetName` で指定されたConfigシートオブジェクトを `wsConfig` に取得。
            *   取得失敗時は `SafeWriteErrorLog` でログ記録、ユーザーにMsgBox表示、`LoadConfiguration = False` で終了。
        *   `configStruct.ConfigSheetFullName` に取得したシートのフルネームを設定。
        *   **「Configシート定義」のAセクション「全般設定」の読み込み:**
            *   `O3` (デバッグモードフラグ): `GetCellValue` を使用して読み込み、Boolean型に変換して `configStruct.DebugModeEnabled` (仮メンバー名、`tConfigSettings`で定義) に格納。`TRUE`/`FALSE`以外の文字列の場合は警告ログを出し、デフォルト`FALSE`とする。
            *   `O12` (デフォルトフォルダパス): `GetCellValue` で読み込み、String型として `configStruct.DefaultFolderPath` に格納。
            *   `O43` (抽出結果出力シート名): `GetCellValue` (必須指定) で読み込み。
            *   `O44` (検索条件ログシート名): `GetCellValue` (必須指定) で読み込み。
            *   `O45` (エラーログシート名): `GetCellValue` (必須指定) で読み込み。
            *   `O46` (設定ファイルシート名): `GetCellValue` で読み込み。
            *   `O122` (工程パターンデータ取得方法): `GetCellValue` で読み込み、Boolean型に変換して `configStruct.UseExcelFormulasForPatternData` (仮メンバー名) に格納。`TRUE`/`FALSE`以外の文字列の場合は警告ログを出し、デフォルト`FALSE`とする。
        *   **「Configシート定義」のBセクション「工程表ファイル内 設定」の読み込み:**
            *   `O66`-`O75` (検索対象シート名リスト): `LoadStringList` ヘルパー関数（後述）を呼び出して `configStruct.TargetSheetNames` 配列に格納。
            *   `O87` (工程表ヘッダー行数): `GetCellValue` (必須、最小値0) で読み込み、Long型に変換。
            *   `O88` (工程表ヘッダー列数): `GetCellValue` (必須、最小値0) で読み込み、Long型に変換。
            *   `O89` (1日のデータが占める行数): `GetCellValue` (必須、最小値1) で読み込み、Long型に変換。
            *   `O90` (1シート内の最大日数): `GetCellValue` (必須、最小値1) で読み込み、Long型に変換。
            *   `O101` (年のセルアドレス): `GetCellValue` (必須) で読み込み。`IsValidCellAddress` で形式検証。
            *   `O102` (月のセルアドレス): `GetCellValue` (必須) で読み込み。`IsValidCellAddress` で形式検証。
            *   `O103` (日の値がある列文字): `GetCellValue` (必須) で読み込み、UCase変換。列文字として有効か簡易検証。
            *   `O104` (日の値の行オフセット): `GetCellValue` (必須、最小値1) で読み込み、Long型に変換。
            *   `O114` (1日の工程数): `GetCellValue` (必須、最小値1) で読み込み、Long型に変換。
        *   `errorOccurred` フラグが `True` なら `LoadConfiguration = False` で終了。そうでなければ `True` で終了。
        *   エラーハンドラ `LoadConfiguration_Error:` では、`SafeWriteErrorLog` でエラーを記録し、`LoadConfiguration = False` で終了。
    *   **`Private Function GetCellValue(...) As Variant` プロシージャ:**
        *   引数: `targetSheet As Worksheet`, `cellAddressString As String`, `callerProcName As String`, `ByRef errorFlag As Boolean`, `itemDescription As String`, `Optional isRequiredField As Boolean = False`, `Optional validationMinValue As Variant`, `Optional validationMaxValue As Variant`。
        *   機能: 指定されたセルの値を読み取り、必須チェック、数値範囲チェック（オプション）を行う。エラーがあれば `errorFlag` を `True` にし、`ReportConfigError` を呼び出す。
        *   詳細な日本語コメントを記述。
    *   **`Private Sub LoadStringList(...)` プロシージャ:**
        *   引数: `ByRef targetStringArray() As String`, `sourceSheet As Worksheet`, `columnLetter As String`, `firstRow As Long`, `lastRow As Long`, `callerProcName As String`, `listDescription As String`。
        *   機能: 指定された列の範囲から空白でない文字列を読み込み、`targetStringArray` 動的配列に格納する。
        *   詳細な日本語コメントを記述。
    *   **`Private Function ParseOffset(...) As Boolean` プロシージャ (今回は枠組みのみでOK):**
        *   引数: `offsetString As String`, `ByRef resultOffset As tOffset`。
        *   機能: オフセット文字列("行,列")を解析する。今回は呼び出されないので、`ParseOffset = True` を返すだけのスタブで可。
    *   **`Private Sub LoadProcessPatternColNumbers(...)` プロシージャ (今回は枠組みのみでOK):**
        *   引数: `ByRef configStruct As tConfigSettings`, `sourceSheet As Worksheet`, `callerProcName As String`, `ByRef errorFlag As Boolean`。
        *   機能: 工程パターン列数を読み込む。今回は呼び出されないので空のスタブで可。
    *   **`Private Sub LoadProcessDetails(...)` プロシージャ (今回は枠組みのみでOK):**
        *   引数: `ByRef configStruct As tConfigSettings`, `sourceSheet As Worksheet`, `callerProcName As String`, `ByRef errorFlag As Boolean`。
        *   機能: 工程詳細（管内など）を読み込む。今回は呼び出されないので空のスタブで可。
    *   **`Private Sub ReportConfigError(...)` プロシージャ:**
        *   引数: `ByRef overallErrorFlag As Boolean`, `callerProcName As String`, `errorCellOrArea As String`, `errorMessageText As String`。
        *   機能: `overallErrorFlag` を `True` に設定し、`SafeWriteErrorLog` を呼び出してエラーログシートにConfig設定エラーを記録する。
        *   詳細な日本語コメントを記述。
    *   **`Private Function IsValidCellAddress(cellAddressString As String) As Boolean` プロシージャ:**
        *   引数: `cellAddressString As String`。
        *   機能: 与えられた文字列がExcelの有効なセルアドレス形式（例: "A1", "FR2"）であるか簡易的に検証する。
        *   詳細な日本語コメントを記述。
    *   **`Private Function ConfigReader_IsArrayInitialized(arr As Variant) As Boolean` (このモジュール専用):**
        *   配列が有効に初期化されているか確認するヘルパー関数。
        *   詳細な日本語コメントを記述。

**2. `M01_MainControl` モジュール (更新):**

*   **`ExtractDataMain()` プロシージャ内:**
    *   `M02_ConfigReader.LoadConfiguration(g_configSettings, wbThis, CONFIG_SHEET_DEFAULT_NAME)` を呼び出す処理のコメントを解除し、実際に呼び出すようにします。
    *   `LoadConfiguration` の戻り値（Boolean）を確認し、`False` であればユーザーにエラーメッセージを表示して `FinalizeMacro_M01` へジャンプする処理を追加します。

**生成コードに関する期待:**
*   「System Instructions」で定義されたコーディング規約（コメント、命名、禁止事項など）を完全に遵守してください。
*   Configシートから「全般設定」と「工程表ファイル内 設定」のセクションを正しく読み込み、グローバル設定構造体 `g_configSettings` の対応するメンバーに格納できること。
*   読み込み時の基本的な検証（必須、型、範囲、形式）と、エラー発生時のログ記録、エラーフラグ管理が実装されていること。
*   まだ実装しない機能（オフセット読み込み、複雑な工程パターン読み込み、フィルター読み込みなど）については、対応するプロシージャの枠組み（スタブ）だけを作成し、中身は空または単純な `Exit Sub`/`Exit Function` としてください。
*   全てのプロシージャ、変数、定数には、その役割や意図を説明する詳細な日本語コメントを必ず付与してください。

**成果物:**
上記の指示に基づき生成された、`M02_ConfigReader` モジュールの完全なVBAコード、および修正・更新された `M01_MainControl` モジュールのVBAコード。

