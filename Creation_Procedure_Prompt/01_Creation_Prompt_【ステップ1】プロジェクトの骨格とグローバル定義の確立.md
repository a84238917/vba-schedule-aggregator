このステップでは、マクロ全体の基本的な構造と、中心となるデータ型を定義することに焦点を当てます。AIには、保守性と可読性を最大限に高めるような骨格作りを依頼します。

---

## プロンプト (AI: Jules向け) - ステップ1: プロジェクト骨格とグローバル定義

**依頼プロジェクト名:** VBA Schedule Aggregator (工程表データ集約マクロ) - 再構築 (ステップ1)

**現在の開発ステップ:** 【ステップ1】プロジェクトの骨格とグローバル定義の確立

**このステップの目的:**
VBAマクロプロジェクトの最も基本的な骨組みを構築します。具体的には、グローバルな宣言（定数、ユーザー定義型、変数）を行う専用モジュールと、マクロ全体の処理フローを制御するメインモジュールを作成し、基本的な初期化処理とエラーハンドリングの枠組みを設けます。このステップで作成するコードは、今後の全開発フェーズの基盤となります。

**参照必須ドキュメント:**
1.  `docs/01_System_Instructions_for_AI.md` (特に、原則1, 2, 3, 5, 7 を厳守)
2.  `docs/04_Config_Sheet_Definition.md` (特に、`tConfigSettings` 型のメンバー構成を決定するために全体を把握)
3.  `docs/07_Naming_Conventions_and_Glossary.md` (特に、接頭辞 `g_`, `t`, モジュール名 `M00_`, `M01_` の規約)

**あなたのタスク:**
以下の仕様に基づき、2つの新しい標準モジュールの完全なVBAコードを生成してください。

**1. `M00_GlobalDeclarations` モジュール:**

*   **目的:** プロジェクト全体で共有されるグローバル定数、Publicなユーザー定義型、およびPublicなグローバル変数を一元的に宣言・管理します。
*   **実装要件:**
    *   モジュールの先頭に `Option Explicit` を記述してください。
    *   モジュールヘッダーコメントとして、このモジュールの役割（上記目的）を日本語で記述してください。
    *   **グローバルデバッグフラグ定数:**
        *   `Public Const DEBUG_MODE_ERROR As Boolean = True`
        *   `Public Const DEBUG_MODE_WARNING As Boolean = True` (今回は `False` に初期設定してください)
        *   `Public Const DEBUG_MODE_DETAIL As Boolean = True` (今回は `False` に初期設定してください)
        *   各定数の役割を説明するコメントを記述してください。
    *   **固定設定値定数:**
        *   `Public Const CONFIG_SHEET_DEFAULT_NAME As String = "Config (2)"` (実際のConfigシート名に合わせてください)
        *   この定数の役割を説明するコメントを記述してください。
    *   **ユーザー定義型 (`Public Type`):**
        *   `tOffset`: `Row As Long`, `Col As Long` をメンバーとして持つ。各メンバーの役割コメントを記述。
        *   `tProcessDetail`: `Kankatsu1 As String`, `Kankatsu2 As String` をメンバーとして持つ。各メンバーの役割コメントを記述。
        *   `tConfigSettings`: 「`docs/04_Config_Sheet_Definition.md`」に記載されている**全てのConfigシート設定項目**（AセクションからGセクションまで）を網羅するメンバーを持つように設計してください。
            *   各メンバーのデータ型は、Configシート定義の「設定例/詳細説明」から適切に判断してください（例: `O3`は`Boolean`、`O12`は`String`、`O87`は`Long`、`O66-O75`は`String`型の動的配列、`ProcessDetails()`は`tProcessDetail`型の動的配列など）。
            *   **特に重要なメンバー:** `Is...OriginallyEmpty` フラグ（例: `IsOffsetSonotaOriginallyEmpty As Boolean`）、`ProcessPatternColNumbers() As Variant`、`ProcessDetails() As tProcessDetail` なども忘れずに含めてください。
            *   `tConfigSettings` 型の各メンバー宣言の横には、対応するConfigシートの代表セル範囲（例: `O43`) と簡単な項目名をコメントとして記述してください。
    *   **グローバル変数 (`Public`):**
        *   `g_configSettings As tConfigSettings` (型名は正確に `tConfigSettings` としてください)
        *   `g_errorLogWorksheet As Worksheet`
        *   `g_nextErrorLogRow As Long`
        *   各グローバル変数の役割を説明するコメントを記述してください。
*   **注意:** このモジュールは必ず**標準モジュール**として作成されることを想定しています。

**2. `M01_MainControl` モジュール:**

*   **目的:** マクロ全体の実行エントリーポイントを提供し、主要な処理フェーズの呼び出しフローを定義し、基本的な初期化処理と包括的なエラーハンドリングを行います。
*   **実装要件:**
    *   モジュールの先頭に `Option Explicit` を記述してください。
    *   モジュールヘッダーコメントとして、このモジュールの役割を日本語で記述してください。
    *   **メインプロシージャ `ExtractDataMain() As Sub`:**
        *   プロシージャヘッダーコメントとして、このプロシージャがマクロの実行開始点であることを記述してください。
        *   ローカル変数を宣言（例: `wbThis As Workbook`, `startTime As Double`など）。各変数には役割コメントを付与。
        *   **初期化シーケンス:**
            *   `On Error GoTo GlobalErrorHandler_M01` (このプロシージャ専用のエラーハンドララベル) を設定。
            *   `Application.ScreenUpdating = False` など、Excelの基本的なアプリケーション設定変更処理を記述。
            *   `startTime = Timer` で処理開始時刻を記録。
            *   `Set wbThis = ThisWorkbook` を実行。
            *   `Call InitializeConfigStructure(g_configSettings)` を呼び出し。
            *   `g_configSettings.StartTime = Now()` と `g_configSettings.ScriptFullName = wbThis.FullName` を設定。
            *   デバッグモードがONの場合、イミディエイトウィンドウに「マクロ実行開始。初期化処理完了。」といったログを出力。
        *   **主要処理フェーズの呼び出しスタブ（コメントとして記述）:**
            *   `' --- 1. Configシート読み込みフェーズ ---`
            *   `' Call M02_ConfigReader.LoadConfiguration(...)`
            *   `' --- 2. 各種シート準備フェーズ ---`
            *   `' Call M03_SheetManager.PrepareSheets(...)`
            *   `' --- 3. 処理対象ファイル特定フェーズ ---`
            *   `' Call M05_FileProcessor.GetTargetFiles(...)`
            *   `' --- 4. 出力/ログ準備フェーズ ---`
            *   `' Call M03_SheetManager.PrepareOutputSheet(...)`
            *   `' --- 5. 検索条件ログ出力フェーズ ---`
            *   `' Call M04_LogWriter.WriteFilterLog(...)`
            *   `' --- 6. メインループフェーズ (ファイルごとのデータ抽出処理) ---`
            *   `' For Each ... Call M06_DataExtractor.ExtractDataFromFile(...) ... Next`
        *   **終了処理シーケンスラベル `FinalizeMacro_M01:`:**
            *   `On Error Resume Next` (終了処理中のエラーは基本的に無視)。
            *   `Application.ScreenUpdating = True` など、Excelのアプリケーション設定を元に戻す処理。
            *   `endTime = Timer` で処理終了時刻を記録。
            *   完了メッセージボックス表示のスタブ（コメントとして「MsgBox "処理完了 (仮)"」など）。
            *   `Set wbThis = Nothing` など、オブジェクト変数の解放処理。
            *   デバッグモードがONの場合、イミディエイトウィンドウに「マクロ実行正常終了。」といったログを出力。
        *   **エラーハンドララベル `GlobalErrorHandler_M01:`:**
            *   エラー情報を取得（`Err.Number`, `Err.Description`, `Err.Source`）。
            *   デバッグモードがONの場合、エラー情報をイミディエイトウィンドウに出力。
            *   （将来的に実装する`M04_LogWriter.WriteErrorLog`または`SafeWriteErrorLog`を呼び出すコメントスタブ）
            *   ユーザーにエラーメッセージボックスを表示。
            *   `Resume FinalizeMacro_M01` で終了処理へジャンプ。
    *   **補助プロシージャ `InitializeConfigStructure(ByRef configStruct As tConfigSettings) As Sub`:**
        *   `Private`スコープで宣言。
        *   プロシージャヘッダーコメントとして、引数で受け取った`tConfigSettings`型の構造体の全メンバー（特に動的配列）を初期化（`Erase`）する役割を記述。
        *   `tConfigSettings` 型の全ての動的配列メンバーに対して `Erase` を実行するコードを記述。
        *   デバッグモードがONの場合、処理の開始と終了をイミディエイトウィンドウに出力。
    *   **補助関数 `LogMain_IsArrayInitialized(arr As Variant) As Boolean`:**
        *   `Private`スコープで宣言。
        *   引数で受け取ったVariant型変数が、有効な要素を持つ配列として初期化されているか確認する。
        *   （前回の回答で提示した `LogMain_IsArrayInitialized` の実装をそのまま使用してください）

**生成コードに関する期待:**
*   「System Instructions」で定義されたコーディング規約（コメント、命名、禁止事項など）を完全に遵守してください。
*   現時点では、実際にConfigシートを読み込んだり、ファイルを処理したりするロジックは実装せず、あくまでプロジェクトの骨格と、今後の開発の土台となるグローバル定義の確立に集中してください。
*   全てのプロシージャ、変数、定数には、その役割や意図を説明する詳細な日本語コメントを必ず付与してください。

**成果物:**
上記の指示に基づき生成された、`M00_GlobalDeclarations` モジュールと `M01_MainControl` モジュールの完全なVBAコード。
