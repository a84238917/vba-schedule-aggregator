このステップでは、実際に1つの工程表ファイルを開き、そこから基本的な情報（年月、各日付）を抽出するコアなロジックの第一歩を実装します。工程パターンは最もシンプルな「パターン1」固定とし、フィルター処理はまだ行いません。ログ出力機能を活用して、抽出した情報を確認できるようにします。

---

## プロンプト (AI: Jules向け) - ステップ4: 単一ファイル指定と基本情報抽出

**依頼プロジェクト名:** VBA Schedule Aggregator (工程表データ集約マクロ) - 再構築 (ステップ4)

**現在の開発ステップ:** 【ステップ4】単一ファイルの指定と基本情報抽出 (工程パターン固定、フィルターなし)

**このステップの目的:**
マクロの核心機能であるデータ抽出処理の最初の段階として、**単一の指定された工程表ファイル**を開き、そのファイル内の**指定された1枚のシート**から、**年月情報および各日付情報**を正しく読み取る機能を実装します。このステップでは、複雑な工程パターンやフィルター処理は導入せず、最も基本的なデータアクセスとループ処理の骨格を確立します。抽出された情報は、デバッグログまたはフィルターログ（処理ログ）シートに出力して確認します。

**参照必須ドキュメント:**
1.  `docs/00_Project_Overview.md` (データ抽出の基本的な流れ)
2.  `docs/01_System_Instructions_for_AI.md` (特に、原則1, 2, 3, 5, 6, 7 を厳守)
3.  `docs/03_Functional_Specification.md` (セクション3.2「対象ファイル処理機能」、3.3「データ抽出機能」の初期段階、3.6.1-3.6.3の該当部分)
4.  `docs/04_Config_Sheet_Definition.md` (特に、**Bセクション「工程表ファイル内 設定」** (`O66`-`O75`, `O87`-`O114`)、および**Cセクション「工程パターン定義」** のうち、`O126`（今回は固定値"1"を想定）、`I列`(工程キー)、`J列`(管内1)、`K列`(管内2)、`L列`(分類1)、`M列`(分類2)、`N列`(分類3)、そして**`O列` (パターン1の工程列数)** の読み込みに関連する部分、**E-2「処理対象ファイル/フォルダパスリスト」**(`P557`の最初の1行のみ対象))
5.  `docs/05_Expected_Behavior.md` (セクション4「処理対象ファイルの特定」の簡易版、セクション6「メイン処理ループ」のファイルオープン、シート処理、年月取得、日取得までの基本フロー)
6.  `docs/07_Naming_Conventions_and_Glossary.md` (特に、接頭辞 `g_`, `t`, `ws`, `wb`, モジュール名 `M02_`, `M05_`, `M06_` の規約)

**前提条件:**
*   「【ステップ1】」で作成された `M00_GlobalDeclarations` モジュール。
*   「【ステップ2】」で作成された `M04_LogWriter` モジュールおよび `M03_SheetManager` モジュールの基本ログ機能。
*   「【ステップ3】」で作成された `M02_ConfigReader` モジュール（全般設定と工程表基本フォーマット設定の読み込みが実装済みであること）。

**あなたのタスク:**
以下の仕様に基づき、`M05_FileProcessor` モジュールを新規作成し、`M02_ConfigReader` モジュールおよび `M06_DataExtractor` モジュールに機能を追加/新規作成し、`M01_MainControl` モジュールを更新するVBAコードを生成してください。

**1. `M05_FileProcessor` モジュール (新規作成):**

*   **目的:** 処理対象となる工程表ファイルのリストを取得・管理します。このステップでは、Configシートの特定セルから単一のファイルパスを取得する簡易版を実装します。
*   **実装要件:**
    *   モジュールの先頭に `Option Explicit`、モジュールヘッダーコメントを記述。
    *   **`Public Function GetTargetFiles(ByRef config As tConfigSettings, ByRef targetFilesCollection As Collection) As Boolean` プロシージャ:**
        *   プロシージャヘッダーコメントを記述。
        *   ローカル変数（`fso As Object`, `filePathFromConfig As String` など）を宣言し、コメント付与。
        *   `FileSystemObject` をレイトバインディングで生成。
        *   **このステップでは、`config.FilePaths(LBound(config.FilePaths))` のようにして、Configから読み込まれた `FilePaths` 配列の最初の要素のみを処理対象とする簡易実装で構いません。** (Configシートの `P557` セルの値を直接参照する形で読み込み、その1行だけを処理対象とする、よりシンプルな実装でも可です。AIの判断に任せます。)
        *   取得したファイルパスが存在し、かつExcelファイル（`IsExcelFile` ヘルパー関数で判定）であれば、`targetFilesCollection` に追加します。
        *   ファイルが存在しない、またはExcelファイルでない場合はエラーログに記録し、`GetTargetFiles = False` を返します。
        *   1件でも有効なファイルが取得できれば `GetTargetFiles = True` を返します。
    *   **`Private Function IsExcelFile(ByVal fileName As String) As Boolean` プロシージャ:**
        *   指定されたファイル名がExcel拡張子(`.xlsx`, `.xls`, `.xlsm`)を持つか判定（前回のコードを流用可）。
    *   **`Private Function LogFileProcessor_IsArrayInitialized(arr As Variant) As Boolean` (このモジュール専用):**
        *   配列初期化確認ヘルパー（前回のコードを流用可）。

**2. `M02_ConfigReader` モジュール (機能追加):**

*   **`LoadConfiguration` プロシージャの拡充:**
    *   **Bセクション「工程表ファイル内 設定」(`O66`-`O114`)の読み込み処理を実装または確認・完成させてください。**
    *   **Cセクション「工程パターン定義」の読み込み（このステップでの限定版）:**
        *   `configStruct.FileProcessPatternToUse` には、このステップでは固定で `1` （またはConfigの`O127`から読み込んだ値で、その値が"1"であることを想定）を設定するようにしてください。
        *   `configStruct.ProcessesPerDay` の値に基づき、以下の配列を `ReDim` します。
            *   `configStruct.ProcessDetails(0 To configStruct.ProcessesPerDay - 1) As tProcessDetail`
            *   `configStruct.ProcessPatternColNumbers(1 To 1)` とし、`configStruct.ProcessPatternColNumbers(1)` も `(0 To configStruct.ProcessesPerDay - 1) As Long` で `ReDim`。 (MaxProcessPatterns は 1 固定)
        *   `LoadProcessDetails` を呼び出し、Configシートの **`I列` (工程キー用、今回は読み飛ばし可)、`J列` (管内1)、`K列` (管内2)、`L列` (分類1)、`M列` (分類2)、`N列` (分類3)** の `129`行目から `ProcessesPerDay` 行数分のデータを読み込み、`configStruct.ProcessDetails` の各要素の対応するメンバーに格納してください。
        *   `LoadProcessPatternColNumbers` を呼び出し、Configシートの **`O列` (パターン1のシート名ヘッダー"第1週"等に対応する工程列数)** の `129`行目から `ProcessesPerDay` 行数分の数値を読み込み、`configStruct.ProcessPatternColNumbers(1)` 配列に格納してください。
    *   **Eセクション「処理対象ファイル定義」の読み込み（このステップでの限定版）:**
        *   `P557`セルからファイルパスを1つだけ読み込み、`configStruct.FilePaths(0)` に格納する処理を `LoadStringList` を使わずに直接記述するか、`LoadStringList` が1行だけでも対応できるようにしてください。(`LoadStringList`を呼び出し、`P557`から`P557`までの範囲を指定するのでも可)
    *   **Fセクション「抽出データオフセット定義」はまだ読み込まないでください（次のステップで実装）。**

**3. `M06_DataExtractor` モジュール (新規作成/枠組みからの拡充):**

*   **目的:** 単一の工程表ファイルから基本的な情報（年月、日）を抽出し、ログに出力します。
*   **実装要件:**
    *   モジュールの先頭に `Option Explicit`、モジュールヘッダーコメントを記述。
    *   **`Public Function ExtractDataFromFile(...) As Boolean` プロシージャ (枠組みから拡充):**
        *   引数は前回定義通り (`kouteiFilePath`, `config`, `wsOutput`, `outputNextRow`, `currentFileNum`, `totalExtractedCount`)。
        *   ローカル変数（`wbKoutei`, `wsKoutei`, `currentYear`, `currentMonth`, `dayLoopIdx`, `dayValueFromCell`, `currentDateInLoop` など）を宣言し、コメント付与。
        *   **ファイルオープン処理:** `kouteiFilePath` を読み取り専用で開きます。失敗時はエラーログ記録、`ExtractDataFromFile = False` で終了。
        *   **工程パターンは固定"1"を想定:** `config.FileProcessPatternToUse` (または `config.ProcessPatternColNumbers(1)` など) を参照して、固定パターンで処理を進める準備をします（このステップではまだ複雑なパターン切り替えは不要）。
        *   **シート処理ループ (今回はリストの最初のシートのみ対象):**
            *   `config.TargetSheetNames(LBound(config.TargetSheetNames))` のようにして、Configから読み込んだ「検索対象シート名リスト」の最初のシート名のみを処理対象とします。
            *   該当シートが存在しない場合はエラーログ記録、このファイルの処理を（部分的）失敗として終了。
            *   **年月取得:** `config.YearCellAddress`, `config.MonthCellAddress` から年月を取得。失敗時はエラーログ記録、このシートの処理を失敗として終了（または流用ロジックのスタブ）。
            *   **日処理ループ:** 1日から `config.DaysPerSheet` までループ。
                *   **日付確定:** `config.DayColumnLetter`, `config.DayRowOffset` から日付セル値を取得し、`DateSerial` で日付を確定。
                    *   **エラー処理/検証:** 日付セルが空、非数値、不正な日などの場合はエラーログに記録し、その日の処理をスキップ（`GoTo NextDayInLoop_Label`など）。
                *   **ログ出力:** 確定した日付（`currentDateInLoop`）を、`M04_LogWriter.WriteFilterLogEntry_General` を使って**検索条件ログシート**に記録します（カテゴリ「データ抽出テスト」、メッセージ「日付抽出成功」、詳細「ファイル名/シート名/日付」など）。
                *   **このステップでは、工程処理ループと詳細なデータ項目抽出はまだ行いません。**
        *   ファイルクローズ処理。
    *   **`Private Function GetValueFromOffset(...) As Variant` プロシージャ (スタブ):**
        *   今回はまだ使用しないため、`GetValueFromOffset = ""` を返すだけのスタブで構いません。
    *   **`Private Function PerformFilterCheck(...) As Boolean` プロシージャ (スタブ):**
        *   今回はまだ使用しないため、`PerformFilterCheck = True` を返すだけのスタブで構いません。
    *   **`Private Function LogExtractor_IsArrayInitialized(arr As Variant) As Boolean` (このモジュール専用):**
        *   配列初期化確認ヘルパー（前回のコードを流用可）。

**4. `M01_MainControl` モジュール (更新):**

*   **`ExtractDataMain()` プロシージャ内:**
    *   `M05_FileProcessor.GetTargetFiles` 呼び出しのコメントを解除し、実際に呼び出す。
    *   ファイル処理ループ内で `M06_DataExtractor.ExtractDataFromFile` 呼び出しのコメントを解除し、実際に呼び出す。

**生成コードに関する期待:**
*   指定された単一の工程表ファイルを開き、その中の最初の対象シートから、年月と各日付を正しく読み取り、それらをログシートに記録できること。
*   ファイルオープンエラー、シート不在エラー、日付取得エラーなどが発生した場合に、適切にエラーログが記録され、処理が安全に継続（または終了）すること。
*   工程パターンは「1」固定で、それに関連するConfig情報（管内1/2、分類1-3、O列の工程列数）が読み込まれること（ただし、これらの情報をまだ抽出ロジックでは使用しない）。
*   オフセットに基づく詳細なデータ項目抽出やフィルター処理はまだ実装しない。
*   「System Instructions」のコーディング規約（特にコメントと命名）を遵守すること。

**成果物:**
上記の指示に基づき生成された、`M05_FileProcessor` モジュールのVBAコード、および修正・更新された `M02_ConfigReader`, `M06_DataExtractor`, `M01_MainControl` モジュールのVBAコード。
