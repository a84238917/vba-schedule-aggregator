このステップでは、前ステップで実装したファイルオープンと日付取得の基盤の上に、Configシートで定義されたオフセット情報に基づいて具体的なデータ項目を抽出し、それらを整形して「一覧」シートにヘッダー付きで書き出す機能を実装します。フィルター処理はまだ導入せず、抽出できたデータは原則として全て出力します。

---

## プロンプト (AI: Jules向け) - ステップ5: オフセット抽出と全件出力

**依頼プロジェクト名:** VBA Schedule Aggregator (工程表データ集約マクロ) - 再構築 (ステップ5)

**現在の開発ステップ:** 【ステップ5】オフセットに基づくデータ項目抽出と「一覧」シートへの全件出力 (単一ファイル、工程パターン固定、フィルターなし)

**このステップの目的:**
「Config」シートで定義されたオフセット情報に基づき、工程表ファイル内の各工程ブロックから具体的なデータ項目（工番、作業場所、作業内容、担当者、人数、作業員名など）を抽出するロジックを実装します。さらに、抽出されたデータを整形し、指定された「一覧」シートにヘッダー行と共に全件書き出す機能を確立します。このステップでは、工程パターンは依然として固定（例：パターン1）とし、データ絞り込みのためのフィルター処理はまだ実装しません。

**参照必須ドキュメント:**
1.  `docs/00_Project_Overview.md` (データ抽出と出力の基本概念)
2.  `docs/01_System_Instructions_for_AI.md` (特に、原則1, 2, 3, 5, 6, 7 を厳守)
3.  `docs/03_Functional_Specification.md` (セクション3.3「データ抽出機能」、3.4「工程パターン適応機能」の初期段階、3.6「データ出力機能」)
4.  `docs/04_Config_Sheet_Definition.md` (特に、**Fセクション「抽出データオフセット定義」** (`N778`-`O792`)、**Gセクション「出力シート設定」** (`O811`-`O821`, `O1124`)、および**Cセクション「工程パターン定義」** のうち、`O列`（パターン1の工程列数）の利用)
5.  `docs/05_Expected_Behavior.md` (セクション6「メイン処理ループ」内の工程処理ループ、データ項目抽出、空白行判定、データ書き出しに関する記述)
6.  `docs/07_Naming_Conventions_and_Glossary.md` (命名規則全般)

**前提条件:**
*   「【ステップ1】～【ステップ4】」で作成・更新された `M00`～`M06` の各モジュールが存在し、特に `M02_ConfigReader` でConfigのA, BセクションおよびCセクションの一部（パターン1関連）が読み込める状態であること。また、`M06_DataExtractor` で単一ファイルのオープン、シート処理、年月・日取得、およびログ出力の基本機能が実装済みであること。

**あなたのタスク:**
以下の仕様に基づき、`M02_ConfigReader`、`M03_SheetManager`、および `M06_DataExtractor` モジュールに機能を追加・拡充し、`M01_MainControl` モジュールを適宜更新するVBAコードを生成してください。

**1. `M02_ConfigReader` モジュール (機能追加):**

*   **`LoadConfiguration` プロシージャの拡充:**
    *   **Fセクション「抽出データオフセット定義」 (`N778`-`O792`) の読み込み処理を実装してください。**
        *   `N列`の項目名と`O列`のオフセット値（文字列 "行,列"）をペアで読み込み、`g_configSettings` 構造体内の対応する `tOffset` 型メンバー（例: `g_configSettings.OffsetKouban`, `g_configSettings.OffsetSagyouinStart` など）に格納します。
        *   オフセット文字列の解析には、以前スタブとして作成した `ParseOffset(offsetString As String, ByRef resultOffset As tOffset) As Boolean` ヘルパー関数を本格実装し、これを使用してください。`ParseOffset` が `False` を返した場合（書式不正など）は、エラーログに記録し、該当オフセットは(0,0)として扱うか、あるいはエラーフラグを立ててください。
        *   `Is...OriginallyEmpty` フラグ（例: `IsOffsetSonotaOriginallyEmpty`）も、オフセット文字列が元々Configシートで空欄だったかどうかに基づいて適切に設定してください。
    *   **Gセクション「出力シート設定」 (`O811`-`O821`, `O1124`) の読み込み処理を実装してください。**
        *   `O811` (ヘッダー行数): Long型で読み込み、範囲検証 (1-10)。
        *   `O812`-`O821` (ヘッダー内容): `OutputHeaderRowCount` 分の文字列を動的配列 `OutputHeaderRows()` に格納。
        *   `O1124` (出力データオプション): 文字列で読み込み ("リセット" or "引継ぎ")。

**2. `M03_SheetManager` モジュール (機能追加):**

*   **`EnsureSheetExists` プロシージャのヘッダー作成ロジックの本格実装:**
    *   引数 `createHeaders` が `True` で、かつ対象シートが `config.OutputSheetName` と一致する場合、`config.OutputHeaderRowCount` と `config.OutputHeaderRows` 配列の内容に基づき、出力シートにヘッダー行を正確に作成するロジックを実装してください。ヘッダー文字列はタブ区切りで複数列に対応します。
*   **`PrepareOutputSheet` プロシージャの拡充:**
    *   引数 `config.OutputDataOption` が "リセット" の場合に、出力シートの既存データ（ヘッダー行を除く）をクリアする処理を確実に実装してください。「ヘッダー行を除く」の判定には、`GetHeaderRowCount` 関数の結果を利用してください。

**3. `M06_DataExtractor` モジュール (機能拡充):**

*   **`ExtractDataFromFile` プロシージャ内の工程処理ループ (`For processLoopIdx = ...`) を本格実装:**
    *   **データ項目抽出の本格実装:**
        *   現在の `processLoopIdx` に対応する「管内1」「管内2」「分類1/2/3」の情報を、ステップ4で内部変数に格納したパターンデータ（Configの`J`～`N`列由来）から取得し、`extractedData` 配列の所定のインデックスに格納します。
        *   `g_configSettings` に格納された各抽出データ項目のオフセット定義（`OffsetKouban`, `OffsetHenshousho` など）を使用し、`GetValueFromOffset` ヘルパー関数を呼び出して、対応する値を工程表シートから抽出し、`extractedData` 配列の所定のインデックスに格納します。
        *   `Is...OriginallyEmpty` フラグが `True` の項目は、抽出処理を行わず `extractedData` の対応要素を `""` とします。
    *   **作業員名の抽出ロジック:**
        *   `g_configSettings.OffsetSagyouinStart` を作業員1のオフセット基準とします。
        *   現在処理中の工程表シート名と現在の `processLoopIdx` に対応する「工程列数」（パターン1の場合は`g_configSettings.ProcessPatternColNumbers(1)(processLoopIdx)` から取得）を `currentProcessActualNumCols` として取得します。
        *   作業員カウンタを1から最大10までループさせ、かつカウンタが `currentProcessActualNumCols` 以下である間、列オフセットを（作業員カウンタ - 1）だけ加算して作業員名を取得し、`extractedData` 配列の作業員名格納領域（インデックス11～20）に格納します。
        *   抽出できた作業員の実数を `actualExtractedWorkerCount` にカウントします。
    *   **空白行判定ロジックの実装:**
        *   抽出された `extractedData` 配列内の主要項目（仕様書で定義されたもの、例: 日付、管内1/2を除く工番、変電所、作業名1, 作業名2, 分類1, 人数）と、`actualExtractedWorkerCount` が0であるかを総合的に判断し、全ての情報が実質的に空であれば `isRowAllEmpty = True` とします。
        *   `isRowAllEmpty` が `True` の場合は、その工程の処理をスキップし（`GoTo NextProcess_LoopEnd_Label`）、ログに「空白行スキップ」の旨を記録します。
    *   **フィルター処理 (スタブのまま):** `PerformFilterCheck` 関数の呼び出しは行いますが、このステップでは常に `True` を返すスタブのままで構いません。
    *   **「一覧」シートへの書き出し処理:**
        *   （フィルターを通過したと仮定して）`extractedData` 配列の内容を、`wsOutput`（出力シート）の `outputNextRow` 行目に書き込みます。
        *   書き出す列の順序は、「Configシート定義」の「G-2. 出力シートヘッダー内容」で定義されたヘッダーの列順に合わせることを意識してください（AIに列のマッピングを適切に行わせる）。
        *   日付は "yyyy/mm/dd(aaa)" 形式で書き込みます。
        *   書き込み後、`outputNextRow` をインクリメントします。
        *   `totalExtractedCount` をインクリメントします。
*   **`GetValueFromOffset(ws As Worksheet, baseRow As Long, baseCol As Long, offsetVal As tOffset) As Variant` プロシージャの本格実装:**
    *   引数で渡された基準セル位置とオフセットに基づき、対象セルの値を読み取って返す処理を実装します。
    *   オフセット後のセル座標がシート範囲外になる場合はエラーログに記録し、空文字列を返します。
    *   セル読み取り時にエラーが発生した場合もエラーログに記録し、空文字列を返します。
    *   取得した値は `Trim(CStr(value))` で整形して返します。

**4. `M01_MainControl` モジュール (更新):**

*   **`ExtractDataMain()` プロシージャ内:**
    *   `M03_SheetManager.PrepareOutputSheet` の呼び出しが正しく行われ、`outputStartRow` が設定されることを確認します（既存のはず）。
    *   `totalExtractedCount` をループ開始前に0で初期化していることを確認します。

**生成コードに関する期待:**
*   指定された単一の工程表ファイル（固定パターン1に対応する形式）から、Configで定義されたオフセットに従って全てのデータ項目（作業員最大10名含む）を抽出し、「一覧」シートにヘッダー行と共に出力できること。
*   「Config」シートの「出力データオプション」が「リセット」の場合、実行前に一覧シートのデータがクリアされること。
*   空白と判断された工程データは出力されないこと。
*   オフセット定義が不正な場合や、セルアクセスでエラーが発生した場合に、適切にログが記録され、処理が継続（該当データは空白として扱われるなど）すること。
*   「System Instructions」のコーディング規約（特にコメントと命名）を遵守すること。

**成果物:**
上記の指示に基づき修正・拡充された、`M02_ConfigReader`, `M03_SheetManager`, `M06_DataExtractor`, および `M01_MainControl` モジュールのVBAコード。
