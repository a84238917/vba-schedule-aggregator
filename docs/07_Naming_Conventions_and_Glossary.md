## このドキュメントの目的 (存在意義)

この「Naming Conventions and Glossary (命名規則と主要な定義一覧)」ドキュメントは、**「VBA Schedule Aggregator (工程表データ集約マクロ)」プロジェクト内で使用されるVBAコードの要素（変数、定数、プロシージャ、ユーザー定義型、モジュールなど）に対する命名規則と、特に重要なグローバルレベルの定義（定数、ユーザー定義型、変数）の一覧とその意味を明確に定めるもの**です。
この文書の目的は、以下の通りです。

1.  **コードの可読性向上:** 一貫した命名規則を適用することで、コードを読む人が各要素の役割やデータ型を名前から容易に推測できるようにし、プログラム全体の理解を助けます。
2.  **開発効率の向上:** 開発者が命名に迷う時間を減らし、より本質的なロジックの実装に集中できるようにします。
3.  **保守性の向上:** 将来的にコードを修正したり機能を追加したりする際に、既存のコードの構造や命名規則を容易に把握できるようにし、安全かつ効率的なメンテナンスを可能にします。
4.  **チーム開発の円滑化 (該当する場合):** 複数人で開発を行う場合に、全員が共通のルールに従うことで、コードの品質を均一化し、コミュニケーションコストを削減します。

このドキュメントは、AIがコードを生成する際の命名ガイドラインとして、また、人間がコードをレビューしたり保守したりする際の参照資料として機能します。

---

**対象システム:** VBA Schedule Aggregator (工程表データ集約マクロ)
**バージョン:** 1.1 (再構築版)

### 1. 命名規則の基本方針

本プロジェクトにおけるVBAコードの全ての命名対象（変数、定数、プロシージャ、ユーザー定義型、モジュール）は、以下の基本方針に従います。

1.  **言語:** 原則として**英語**を使用します。ただし、Configシートの項目名など、日本語に由来する固有名詞を扱う場合は、その意味が明確に伝わるローマ字表記（例: `Kankatsu1FilterList`）を用いることを許容します。
2.  **明確性と具体性:** 名前は、その要素が持つ役割、格納するデータの内容、または実行する処理が、**曖昧さなく具体的に推測できる**ものとします。極端な省略形や、意味の不明瞭な頭字語の使用は避けてください。
3.  **一貫性:** プロジェクト全体を通じて、同様の役割を持つ要素には一貫した命名パターンを適用します。
4.  **VBAの予約語との衝突回避:** VBAの組み込み関数名、キーワード、オブジェクト名など、予約語と完全に同一の名前は使用しないでください。

### 2. 大文字・小文字の使用規則 (ケーシング)

1.  **プロシージャ名 (Sub, Function):** パスカルケース (PascalCase) を使用します。各単語の先頭を大文字にします。
    *   例: `ExtractDataFromFile`, `InitializeGlobalConfiguration`, `IsValidFile`
2.  **モジュール名:** パスカルケースを使用し、役割を示す接頭辞 `MXX_` (XXは連番) を付与します。
    *   例: `M00_GlobalDeclarations`, `M01_MainControl`
3.  **変数名 (ローカル、モジュールレベル、グローバル):** キャメルケース (camelCase) を使用します。最初の単語の先頭は小文字、以降の単語の先頭を大文字にします。
    *   例: `targetFilePath`, `currentSheetName`, `outputStartRow`, `g_configSettings` (グローバル変数も同様)
4.  **定数名 (Const):** 全て大文字とし、単語間をアンダースコア `_` で区切ります (SNAKE_CASE)。
    *   例: `CONFIG_SHEET_DEFAULT_NAME`, `MAX_FILTER_ITEMS`, `ERROR_FILE_NOT_FOUND`
5.  **ユーザー定義型名 (Type):** パスカルケースを使用し、型であることを明確にするために接頭辞 `t` (小文字) を付けます。
    *   例: `tConfigSettings`, `tProcessDetail`, `tOffset`
6.  **ユーザー定義型のメンバー名:** パスカルケースを使用します（型名内部のメンバーなので、変数名のキャメルケースとは区別）。
    *   例: `Public Type tConfigSettings` の中の `OutputSheetName As String`

### 3. 推奨される接頭辞とその意味

変数のスコープや主要なデータ型を名前から推測しやすくするために、以下の接頭辞の使用を推奨します。これは強制ではありませんが、一貫して使用することでコードの可読性が向上します。

| 接頭辞 | 意味                        | 使用対象の例                                 | 具体例                                      |
| :----- | :-------------------------- | :------------------------------------------- | :------------------------------------------ |
| `g_`   | グローバル変数 (Global)     | `Public` で宣言された標準モジュールレベル変数  | `g_configSettings`, `g_errorLogWorksheet`   |
| `m_`   | モジュールレベル変数 (Module) | `Private` で宣言された標準モジュールレベル変数 | `m_isInitialized`, `m_defaultPatternNumber` |
| `t`    | ユーザー定義型 (Type)       | `Public Type` または `Private Type` 宣言     | `tConfigSettings`, `tProcessDetail`         |
| `ws`   | ワークシート (Worksheet)    | `Worksheet` オブジェクト変数                 | `wsConfig`, `wsOutput`, `wsKoutei`          |
| `wb`   | ワークブック (Workbook)     | `Workbook` オブジェクト変数                  | `wbThis`, `wbKouteiTarget`                  |
| `rng`  | 範囲 (Range)                | `Range` オブジェクト変数                     | `rngDataArea`, `rngHeader`                  |
| `arr`  | 配列 (Array)                | 配列型の変数                                 | `arrFilePaths`, `arrFilterItems`            |
| `col`  | コレクション (Collection)   | `Collection` オブジェクト変数                | `colTargetFiles`, `colErrorMessages`        |
| `dic`  | ディクショナリ (Dictionary) | `Scripting.Dictionary` オブジェクト変数      | `dicPatternCache` (レイトバインディング時)  |
| `obj`  | 汎用オブジェクト (Object)   | レイトバインディング等で具体的な型が事前不明な場合 | `objFileSystem`, `objRegex`                 |
| `b`    | ブール型 (Boolean)          | `Boolean` 型の変数 (主にローカル)            | `bIsValid`, `bContinueProcessing`           |
| `s`    | 文字列 (String)             | `String` 型の変数 (主にローカル)             | `sFilePath`, `sSheetName`                   |
| `l`    | 長整数型 (Long)             | `Long` 型の変数 (主にローカル、ループカウンタ等) | `lRowCount`, `lFileIndex`                   |
| `dt`   | 日付型 (Date)               | `Date` 型の変数 (主にローカル)               | `dtTargetDate`, `dtStartDate`               |

**注意:** `b`, `s`, `l`, `dt` のような純粋なデータ型を示す接頭辞（いわゆるハンガリアン記法）は、VBAのように型宣言が強制される環境では冗長と見なされることもあります。これらは必須ではなく、変数名全体で意味が明確であれば省略可能です。しかし、`g_`, `m_`, `t`, `ws`, `wb` のようなスコープや主要オブジェクト型を示す接頭辞は、コードの構造理解に役立つため、一貫した使用を推奨します。

### 4. 主要なグローバル定義一覧 (M00_GlobalDeclarations で定義)

以下は、本マクロプロジェクトの中核となるグローバルスコープの定数、ユーザー定義型、および変数です。これらはマクロ全体の動作設定や中心的なデータ構造を定義します。

#### 4.1. グローバル定数

| 定数名 (Const Name)         | データ型 (Data Type) | 推奨値/説明                                                                                                   |
| :-------------------------- | :----------------- | :---------------------------------------------------------------------------------------------------------- |
| `DEBUG_MODE_ERROR`          | `Boolean`          | `TRUE` / `FALSE`。エラー関連の詳細デバッグ情報をイミディエイトウィンドウに出力するかどうかのフラグ。                       |
| `DEBUG_MODE_WARNING`        | `Boolean`          | `TRUE` / `FALSE`。警告レベル（データ不整合の可能性など）のデバッグ情報をイミディエイトウィンドウに出力するかどうかのフラグ。 |
| `DEBUG_MODE_DETAIL`         | `Boolean`          | `TRUE` / `FALSE`。詳細な処理追跡情報（ステップ実行、変数内容など）をイミディエイトウィンドウに出力するかどうかのフラグ。   |
| `CONFIG_SHEET_DEFAULT_NAME` | `String`           | `"Config"` (または実際のConfigシート名)。マクロが設定を読み込むConfigシートのデフォルト名を定義します。                |

#### 4.2. 主要ユーザー定義型 (Public Type)

| 型名 (Type Name)  | 説明                                                                                                | 主要メンバー (詳細は「Configシート定義」および他のドキュメント参照)                                                                                                                                                                                                                                                                                                                                                                  |
| :---------------- | :-------------------------------------------------------------------------------------------------- | :--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `tOffset`         | 工程表内の基準セルからの相対的な位置（行オフセット、列オフセット）を格納するためのデータ構造。                         | `Row As Long`, `Col As Long`                                                                                                                                                                                                                                                                                                                                                                                                   |
| `tProcessDetail`  | 各工程（横方向の作業ブロック）に固有の情報を格納するためのデータ構造。主に「Config」シートの工程パターン定義から読み込まれる。 | `Kankatsu1 As String` (管内1情報), `Kankatsu2 As String` (管内2情報)。将来的には分類情報などもここに含めることを検討。                                                                                                                                                                                                                                                                                                             |
| `tConfigSettings` | 「Config」シートから読み込まれるマクロ全体の動作設定を一元的に管理するための、最も重要なグローバルデータ構造。             | `StartTime As Date`, `ScriptFullName As String`, `DefaultFolderPath As String`, `OutputSheetName As String`, (各種ログシート名), `ActualConfigSheetName As String`, `WorkSheetName As String`, `IsConfigDataValid As Boolean` (新設検討: Config読み込み後の検証結果), `DebugModeEnabled As Boolean` (新設: `O3`の値), `UseExcelFormulasForPatternData As Boolean` (新設: `O122`の値), `TargetSheetNames() As String`, (工程表基本フォーマット関連メンバー), (工程定義関連メンバー: `ProcessesPerDay`, `FileProcessCountForPattern`, `FileProcessPatternToUse`, `MaxProcessPatterns`, `ProcessPatternColNumbers() As Variant`, `ProcessDetails() As tProcessDetail`), (各種抽出データオフセットメンバー: `OffsetKouban As tOffset` など、および対応する `Is...OriginallyEmpty As Boolean` フラグ), (各種フィルター条件メンバー: `WorkerFilterLogic As String`, `FilterWorkerNames() As String` など), `FilePaths() As String`, (出力シートヘッダー関連メンバー), `OutputDataOption As String`, `HideSheetOption As String`, `HideSheetNames() As String`, `ConfigSheetFullName As String` など。**詳細は「Configシート定義」および「System Instructions」で列挙されたメンバーを網羅すること。** |

#### 4.3. グローバル変数 (Public Variables)

| 変数名 (Variable Name) | データ型 (Data Type) | 説明                                                                                                                                                                                             |
| :-------------------- | :----------------- | :----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `g_configSettings`    | `tConfigSettings`  | `M02_ConfigReader`によって「Config」シートから読み込まれ、検証された全設定情報を格納する、マクロ全体で共有される最重要変数。この変数の内容は、マクロの全動作を規定します。                                              |
| `g_errorLogWorksheet` | `Worksheet`        | エラーログを書き込むためのワークシートオブジェクトへの参照。`M03_SheetManager`でシート準備後に設定され、`M04_LogWriter`で使用されます。マクロ実行中にエラーが発生した場合に、このオブジェクトを通じてログが記録されます。             |
| `g_nextErrorLogRow`   | `Long`             | エラーログシートにおいて、次にエラー情報を書き込むべき行の番号。`M04_LogWriter`によって管理・更新されます。                                                                                             |

### 5. モジュール命名規則 (提案)

*   `M00_GlobalDeclarations`: グローバルスコープの定数、ユーザー定義型、変数の宣言専用。
*   `M01_MainControl`: マクロ全体の実行フローの開始、主要処理フェーズの呼び出し、終了処理を統括。
*   `M02_ConfigReader`: 「Config」シートから全ての設定情報を読み込み、検証し、`g_configSettings`に格納。
*   `M03_SheetManager`: 出力シート、ログシート等の存在確認、自動生成（ヘッダー含む）、準備、マクロ終了時の表示/非表示制御。
*   `M04_LogWriter`: エラーログシートおよび検索条件ログシートへの整形された情報の書き込み。
*   `M05_FileProcessor`: 「Config」シートの指定またはファイルダイアログから、処理対象となる工程表ファイルのリストを特定し、管理。
*   `M06_DataExtractor`: 個々の工程表ファイルを開き、指定されたシートからデータを抽出し、フィルター条件を適用し、結果を出力シートに書き込む。
*   `M_Utilities` (オプション): プロジェクト内の複数のモジュールで共通して使用できる汎用的なヘルパー関数（例: 配列の初期化状態チェック、文字列の安全な数値変換、日付書式設定など）を格納。

この命名規則と定義一覧が、本プロジェクトのコード品質と保守性の向上に貢献することを期待します。