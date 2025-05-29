このステップでは、今後の開発とデバッグに不可欠となるエラーログと検索条件ログの基本的な書き込み機能を実装し、それらのログシートが存在しない場合には自動生成する仕組みを作ります。

---

## プロンプト (AI: Jules向け) - ステップ2: ログ出力機能の基盤実装

**依頼プロジェクト名:** VBA Schedule Aggregator (工程表データ集約マクロ) - 再構築 (ステップ2)

**現在の開発ステップ:** 【ステップ2】ログ出力機能の基盤実装

**このステップの目的:**
マクロの実行状況や発生したエラーを記録するための基本的なログ出力機能を実装します。具体的には、エラー情報を記録する「エラーログシート」と、マクロ実行時の設定や主要な処理ステップを記録する「検索条件ログシート（処理ログシート）」への書き込みプロシージャ、およびこれらのシートが存在しない場合に自動生成する機能を開発します。このログ機能は、以降の全開発ステップにおけるデバッグと動作検証の基盤となります。

**参照必須ドキュメント:**
1.  `docs/00_Project_Overview.md` (プロジェクトの全体像とログの役割)
2.  `docs/01_System_Instructions_for_AI.md` (特に、原則1, 2, 3, 5, 7 を厳守)
3.  `docs/03_Functional_Specification.md` (セクション3.3, 3.7, 4 のログ関連要件)
4.  `docs/04_Config_Sheet_Definition.md` (特に、`O44`「検索条件ログシート名」、`O45`「エラーログシート名」の設定項目)
5.  `docs/07_Naming_Conventions_and_Glossary.md` (特に、接頭辞 `g_`, `ws`, モジュール名 `M03_`, `M04_` の規約)

**前提条件:**
*   「【ステップ1】プロジェクトの骨格とグローバル定義の確立」で作成された `M00_GlobalDeclarations` モジュール（特に `Public g_Config As tConfigSettings`, `Public g_wsErrorLog As Worksheet`, `Public g_NextErrorLogRow As Long` のグローバル変数宣言）および `M01_MainControl` モジュールの基本構造が存在することを前提とします。

**あなたのタスク:**
以下の仕様に基づき、`M04_LogWriter` モジュールを新規作成し、`M03_SheetManager` モジュールに機能を追加し、`M01_MainControl` モジュールを更新するVBAコードを生成してください。

**1. `M04_LogWriter` モジュール (新規作成):**

*   **目的:** エラーログおよび検索条件ログのシートへの書き込み処理を専門に担当します。
*   **実装要件:**
    *   モジュールの先頭に `Option Explicit` を記述してください。
    *   モジュールヘッダーコメントとして、このモジュールの役割を日本語で記述してください。
    *   **`Public Sub WriteErrorLog(...)` プロシージャ:**
        *   引数: `moduleName As String`, `procedureName As String`, `relatedInfo As String`, `errorNumber As Long`, `errorDescription As String`, `Optional actionTaken As String = ""`, `Optional variableInfo As String = ""`。
        *   機能: グローバル変数 `g_wsErrorLog` で参照されるワークシートの `g_NextErrorLogRow` 行目に、引数で受け取ったエラー情報を書き込みます。書き込み列はA列から順に「発生日時(Now)」「モジュール名」「プロシージャ名」「関連情報」「エラー番号」「エラー内容（先頭にアポストロフィを付加）」「対処内容」「変数情報（長すぎる場合は先頭32767文字）」とします。
        *   書き込み後、`g_NextErrorLogRow` をインクリメントします。
        *   `g_wsErrorLog` が `Nothing` の場合は、デバッグログにエラーを出力して何もせずに終了します。
        *   このプロシージャ内でのエラーは、`WriteErrorLog_InternalError:` ラベルで捕捉し、デバッグログに詳細を出力します（エラーログへの再帰的書き込みは避ける）。
        *   詳細な日本語コメントを記述してください。
    *   **`Public Sub SafeWriteErrorLog(...)` プロシージャ:**
        *   引数: `targetWorkbook As Workbook`, `errorLogSheetNameAttempt As String`, `moduleName As String`, `procedureName As String`, `relatedInfo As String`, `errorNumber As Long`, `errorDescription As String`, `Optional actionTaken As String = ""`, `Optional variableInfo As String = ""`。
        *   機能: `g_wsErrorLog` が未設定の可能性がある状況（例: Config読み込み前）でもエラーログの記録を試みます。引数で指定されたワークブック内の `errorLogSheetNameAttempt` シートに書き込みます。シートが存在しない場合は、簡易ヘッダー（「発生日時」～「変数情報」）付きで新規作成を試みます。
        *   書き込みロジックは `WriteErrorLog` と同様です。
        *   このプロシージャ内でのエラーは `On Error Resume Next` で処理し、呼び出し元に影響を与えないようにします（ログ記録の試みなので、失敗しても致命的ではない）。
        *   詳細な日本語コメントを記述してください。
    *   **`Public Sub WriteFilterLog(...)` プロシージャ (枠組みのみ):**
        *   引数: `ByRef config As tConfigSettings`, `ByVal targetWorkbook As Workbook`。
        *   機能（今回は枠組み）: 指定されたフィルターログシートに、マクロ実行開始を示すログ（実行日時、区切り線、実行ファイルパスなど）のみを書き込む処理を実装します。
        *   ログシートが存在しない場合の処理は、`M03_SheetManager.EnsureSheetExists` に委ねることを想定（後述）。
        *   このプロシージャ内でのエラーはエラーハンドラで捕捉し、`SafeWriteErrorLog` でエラーログに記録してください。
        *   詳細な日本語コメントを記述してください。
    *   **`Private Sub WriteFilterLogEntry(...)` プロシージャ:**
        *   引数: `targetLogSheet As Worksheet`, `ByRef nextLogRow As Long`, `itemName As String`, `itemValue As String`。
        *   機能: 指定されたログシートの指定行に、「実行日時(Now)」「項目名」「値」を書き込み、行番号をインクリメントします。
        *   詳細な日本語コメントを記述してください。
    *   **`Private Sub WriteFilterLogArrayEntry(...)` プロシージャ:**
        *   引数: `targetLogSheet As Worksheet`, `ByRef nextLogRow As Long`, `itemName As String`, `ByRef itemArray() As String`。
        *   機能: `itemArray` が初期化されていれば、その内容を `Join` でカンマ区切り文字列にし、`WriteFilterLogEntry` を呼び出して書き込みます。配列が空や未初期化の場合は、その旨を示す文字列（例: "(リスト空)", "(リスト未設定)"）を値として書き込みます。
        *   詳細な日本語コメントを記述してください。
    *   **`Private Function LogWriter_IsArrayInitialized(arr As Variant) As Boolean`:**
        *   配列が有効に初期化されているか確認するヘルパー関数（前回のコードを流用）。
        *   詳細な日本語コメントを記述してください。

**2. `M03_SheetManager` モジュール (機能追加):**

*   **`Public Function PrepareSheets(...)` プロシージャの拡充:**
    *   引数: `ByRef config As tConfigSettings`, `ByVal targetWorkbook As Workbook`。
    *   機能: Config設定に基づき、エラーログシートとフィルターログシートの準備を行います。
        *   `config.ErrorLogSheetName` で指定されたエラーログシートに対し、`EnsureSheetExists` を呼び出して存在確認と必要なら自動生成（ヘッダー作成指定は `True`）を行います。
        *   `config.FilterLogSheetName` で指定されたフィルターログシートに対し、`EnsureSheetExists` を呼び出して存在確認と必要なら自動生成（ヘッダー作成指定は `True`）を行います。
        *   これらのシート準備に失敗した場合は、`allSheetsPrepared` フラグを `False` に設定します。
    *   詳細な日本語コメントを記述してください。
*   **`Private Function EnsureSheetExists(...)` プロシージャのヘッダー作成ロジック修正:**
    *   引数: `targetWorkbook As Workbook`, `sheetNameToEnsure As String`, `ByRef config As tConfigSettings`, `callerFuncName As String`, `createHeaders As Boolean`。
    *   機能: `createHeaders` が `True` の場合、
        *   もし `sheetNameToEnsure` が `config.OutputSheetName` と一致する場合、`config.OutputHeaderRowCount` と `config.OutputHeaderRows` に基づいてヘッダーを作成します（このロジックはステップ5で本格実装するので、今回はコメントアウトまたは簡単なスタブでも可）。
        *   もし `sheetNameToEnsure` が `config.ErrorLogSheetName` と一致する場合、1行目に固定のヘッダー（例: A1="発生日時", B1="モジュール", C1="プロシージャ", D1="関連情報", E1="エラー番号", F1="エラー内容", G1="対処内容", H1="変数情報"）を作成します。
        *   もし `sheetNameToEnsure` が `config.FilterLogSheetName` と一致する場合、1行目に固定のヘッダー（例: A1="実行日時", B1="フィルター項目", C1="条件", D1="備考"）を作成します。
    *   詳細な日本語コメントを記述してください。
*   **`Private Function IsArrayInitialized(arr As Variant) As Boolean`:**
    *   配列が有効に初期化されているか確認するヘルパー関数（前回のコードを流用）。
    *   詳細な日本語コメントを記述してください。

**3. `M01_MainControl` モジュール (更新):**

*   **`ExtractDataMain()` プロシージャ内:**
    *   `M03_SheetManager.PrepareSheets` 呼び出し後、返り値を確認し、失敗ならエラーメッセージ表示後 `FinalizeMacro_M01` へ。
    *   `g_wsErrorLog` オブジェクトと `g_NextErrorLogRow` の設定処理を実装（`PrepareSheets` 成功後）。
    *   `M04_LogWriter.WriteFilterLog` の呼び出しスタブのコメントを解除し、実際に呼び出すようにします。
*   **`GlobalErrorHandler_M01` ラベル内:**
    *   `M04_LogWriter.WriteErrorLog` または `SafeWriteErrorLog` を呼び出すコメントスタブを解除し、実際にエラー情報をエラーログシートに書き込むようにします。`g_wsErrorLog` が設定されているかで呼び出すプロシージャを分岐させてください。

**生成コードに関する期待:**
*   「System Instructions」で定義されたコーディング規約（コメント、命名、禁止事項など）を完全に遵守してください。
*   現時点では、Configシートからログシート名を取得し、それらのシートを準備し、基本的なログ（実行開始ログ、テスト用のエラーログ）を書き込める状態にすることを目指します。
*   全てのプロシージャ、変数、定数には、その役割や意図を説明する詳細な日本語コメントを必ず付与してください。

**成果物:**
上記の指示に基づき生成された、`M04_LogWriter` モジュールの完全なVBAコード、および修正・更新された `M03_SheetManager` と `M01_MainControl` モジュールのVBAコード。
