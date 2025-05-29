このステップでは、これまで抽出してきたデータに対して、ユーザーが「Config」シートで定義した様々な条件に基づいて絞り込みを行うフィルター機能を実装します。1つのフィルター機能から順に、段階的に追加・テストしていくアジャイルなアプローチを推奨します。

---

## プロンプト (AI: Jules向け) - ステップ8: 各種フィルター機能の実装 (段階的アプローチ)

**依頼プロジェクト名:** VBA Schedule Aggregator (工程表データ集約マクロ) - 再構築 (ステップ8)

**現在の開発ステップ:** 【ステップ8】各種フィルター機能の実装 (1つずつ段階的に)

**このステップの目的:**
抽出された工程データに対し、「Config」シートで定義された多様なフィルター条件を適用し、条件に合致するデータのみを出力対象とする機能を実装します。このステップでは、**一度に全てのフィルターを実装するのではなく、1種類ずつ（または論理的に関連する数種類ずつ）フィルター機能を追加し、その都度テストと検証を行う**ことで、複雑なロジックを段階的かつ確実に構築します。最終的には、「Configシート定義」Dセクションの全てのフィルター項目に対応することを目指します。

**参照必須ドキュメント:**
1.  `docs/01_System_Instructions_for_AI.md` (特に、原則1「設定駆動」、原則2「可読性」、原則3「堅牢性」、原則7「禁止事項」を厳守)
2.  `docs/03_Functional_Specification.md` (セクション3.5「データフィルタリング機能」)
3.  `docs/04_Config_Sheet_Definition.md` (特に、**Dセクション「フィルター条件」** (`O242`-`O544`) の全項目定義、およびフィルター対象となるデータが「Config」シートのどの部分（CセクションやFセクション）から来るかの関連性)
4.  `docs/05_Expected_Behavior.md` (セクション6-6-ii-6「フィルター判定」の具体的な動作フロー)
5.  `docs/07_Naming_Conventions_and_Glossary.md` (関連する変数名・型名の規約)

**前提条件:**
*   「【ステップ1】～【ステップ7】」で作成・更新された `M00`～`M06` の各モジュールが存在し、特に `M02_ConfigReader` でConfigのA, B, C, E, F, Gセクションが読み込める状態であること。
*   `M06_DataExtractor` で、ファイルごとの工程パターンが適用され、オフセットに基づいた全データ項目（作業員含む）が `extractedData` 配列に正しく抽出できる状態であること。
*   `M06_DataExtractor` 内に `PerformFilterCheck` 関数のスタブ（現在は常に`True`を返す）が存在すること。

**あなたのタスク (段階的実装):**
`M02_ConfigReader` モジュールにDセクション「フィルター条件」の読み込み処理を完全実装し、`M06_DataExtractor` モジュール内の `PerformFilterCheck` 関数を本格実装します。以下のフィルターを**指定された順序で1つずつ（または指示されたグループ単位で）**実装し、各実装後にテストと検証を行えるようにしてください。

**1. `M02_ConfigReader` モジュール (機能追加):**

*   **`LoadConfiguration` プロシージャの拡充:**
    *   **Dセクション「フィルター条件」 (`O242`-`O544`) の全ての項目の読み込み処理を実装してください。**
        *   `O242` (作業員フィルター検索論理): 文字列として読み込み、`AND`/`OR` 以外ならデフォルト`OR`。
        *   `O243`-`O262` (作業員フィルターリスト): `LoadStringList` を使用して `g_configSettings.FilterWorkerNames` 配列に格納。
        *   `O275`-`O294` (管内1フィルターリスト): `LoadStringList` を使用して `g_configSettings.FilterKankatsu1` 配列に格納。
        *   `O305`-`O334` (管内2フィルターリスト): `LoadStringList` を使用して `g_configSettings.FilterKankatsu2` 配列に格納。
        *   `O346` (分類1フィルター): カンマまたは" OR "区切り文字列を配列 `g_configSettings.FilterBunrui1` に格納。
        *   `O367` (分類2フィルター): 同様に配列 `g_configSettings.FilterBunrui2` (型定義に追加要) に格納。
        *   `O388` (分類3フィルター): 同様に配列 `g_configSettings.FilterBunrui3` (型定義に追加要) に格納。
        *   `O409`-`O418` (工事種類フィルターリスト): `LoadStringList` を使用して `g_configSettings.FilterKoujiShurui` 配列に格納。
        *   `O431`-`O440` (工番フィルターリスト): `LoadStringList` を使用して `g_configSettings.FilterKouban` 配列に格納。
        *   `O451`-`O470` (作業種類フィルターリスト): `LoadStringList` を使用して `g_configSettings.FilterSagyouShurui` 配列に格納。
        *   `O481`-`O490` (担当の名前フィルターリスト): `LoadStringList` を使用して `g_configSettings.FilterTantouNameList` (型定義に追加要、単一文字列の `FilterTantouName` とは別) に格納。
        *   `O503` (人数フィルター): 文字列として読み込み、数値変換可能か検証。`g_configSettings.FilterNinzuu` に格納。
        *   `O514` (作業箇所の種類フィルター): カンマ区切り文字列を配列 `g_configSettings.FilterSagyouKashoType` に格納。
        *   `O525`-`O544` (作業箇所フィルターリスト): `LoadStringList` を使用して `g_configSettings.FilterSagyouKasho` 配列に格納。
    *   **注意:** `tConfigSettings` 型定義に、上記で新規に必要となる配列メンバー（例: `FilterBunrui2() As String`, `FilterBunrui3() As String`, `FilterTantouNameList() As String`）を忘れずに追加してください (`M00_GlobalDeclarations`の修正)。

**2. `M06_DataExtractor` モジュール (段階的機能拡充):**

*   **`Private Function PerformFilterCheck(...) As Boolean` プロシージャの本格実装 (段階的に):**
    *   この関数は、引数で渡された `extractedData` 配列と `actualExtractedWorkerCount`、そしてグローバルな `g_configSettings` 内の各フィルター設定を参照して、総合的なフィルター判定を行います。
    *   **基本ロジック:** 関数冒頭で `PerformFilterCheck = True` と初期化。各フィルター条件を順にチェックし、いずれかの条件で**不一致**となった場合は、即座に `PerformFilterCheck = False` を設定し、`Exit Function` します（AND結合のイメージ）。全ての有効なフィルター条件をクリアした場合のみ `True` が返ります。
    *   フィルターリストが空、またはフィルター条件セル自体が空の場合は、そのフィルターは「適用なし（常に合致）」として扱います。

    **以下の順序で、1つ（または指示されたグループ）ずつフィルターロジックを実装し、AIにその部分のコード生成を依頼してください。各実装後はテストと検証を行います。**

    *   **段階1: 作業員フィルター (`D-1`, `D-2`)**
        *   `g_configSettings.FilterWorkerNames` 配列が要素を持つ場合のみ処理。
        *   `g_configSettings.WorkerFilterLogic` (`AND`/`OR`) に従い、`extractedData(11)`～`extractedData(10 + actualExtractedWorkerCount)` とフィルターリストを比較（完全一致、`vbTextCompare`）。
        *   `AND` の場合: リスト内の全ての作業員が抽出作業員に含まれていなければ `False`。
        *   `OR` の場合: リスト内のいずれかの作業員も抽出作業員に含まれていなければ `False`。
        *   デバッグログに判定過程と結果を出力。

    *   **段階2: 管内1フィルター (`D-3`) と 管内2フィルター (`D-4`)**
        *   `g_configSettings.FilterKankatsu1` 配列が要素を持つ場合: `extractedData(2)`（管内1）とリスト内の各項目を比較（完全一致、`vbBinaryCompare`）。いずれにも一致しなければ `False`。
        *   `g_configSettings.FilterKankatsu2` 配列が要素を持つ場合: `extractedData(3)`（管内2）とリスト内の各項目を比較（完全一致、`vbBinaryCompare`）。いずれにも一致しなければ `False`。
        *   デバッグログに判定過程と結果を出力。

    *   **段階3: 分類フィルター (`D-5`, `D-6`, `D-7`)**
        *   `g_configSettings.FilterBunrui1` 配列が要素を持つ場合: `extractedData(9)`（抽出された分類1、ConfigのL列由来）とリスト内の各キーワードを比較（部分一致、`vbTextCompare`）。いずれにも一致しなければ `False`。
        *   `g_configSettings.FilterBunrui2` 配列が要素を持つ場合: `extractedData` 配列の対応するインデックス（仮に22番、ConfigのM列由来の分類2を格納）とリスト内の各キーワードを比較（部分一致、`vbTextCompare`）。いずれにも一致しなければ `False`。 **(注意: `extractedData` 配列のインデックス定義と格納ロジックを要確認・調整)**
        *   `g_configSettings.FilterBunrui3` 配列が要素を持つ場合: `extractedData` 配列の対応するインデックス（仮に23番、ConfigのN列由来の分類3を格納）とリスト内の各キーワードを比較（**完全一致**、`vbTextCompare`）。いずれにも一致しなければ `False`。 **(注意: `extractedData` 配列のインデックス定義と格納ロジックを要確認・調整)**
        *   デバッグログに判定過程と結果を出力。

    *   **段階4: 工事種類フィルター (`D-8`) と 工番フィルター (`D-9`)**
        *   `g_configSettings.FilterKoujiShurui` 配列が要素を持つ場合: `extractedData(20)`（工事種類）とリスト内の各キーワードを比較（部分一致、`vbTextCompare`）。いずれにも一致しなければ `False`。
        *   `g_configSettings.FilterKouban` 配列が要素を持つ場合: `extractedData(4)`（工番）とリスト内の各キーワードを比較（部分一致、`vbTextCompare`）。いずれにも一致しなければ `False`。
        *   デバッグログに判定過程と結果を出力。

    *   **段階5: 作業種類フィルター (`D-10`) と 担当の名前フィルター (`D-11`)**
        *   `g_configSettings.FilterSagyouShurui` 配列が要素を持つ場合: `extractedData(6)`（作業名1）または `extractedData(7)`（作業名2）のいずれかと、リスト内の各キーワードを比較（部分一致、`vbTextCompare`）。いずれの作業名に対しても、どのキーワードとも一致しなければ `False`。
        *   `g_configSettings.FilterTantouNameList` 配列が要素を持つ場合: `extractedData(21)`（担当の名前）とリスト内の各キーワードを比較（部分一致、`vbTextCompare`）。いずれにも一致しなければ `False`。
        *   デバッグログに判定過程と結果を出力。

    *   **段階6: 人数フィルター (`D-12`)**
        *   `g_configSettings.FilterNinzuu` が空でない場合: `extractedData(10)`（人数）とフィルター値を比較（数値として完全一致）。一致しなければ `False`。
        *   デバッグログに判定過程と結果を出力。

    *   **段階7: 作業箇所の種類フィルター (`D-13`) と 作業箇所フィルター (`D-14`)**
        *   `g_configSettings.FilterSagyouKashoType` 配列が要素を持つ場合: `extractedData(5)`（変電所/作業箇所名）とリスト内の各キーワードを比較（部分一致、`vbTextCompare`）。いずれにも一致しなければ `False`。
        *   `g_configSettings.FilterSagyouKasho` 配列が要素を持つ場合: `extractedData(5)`（変電所/作業箇所名）とリスト内の各キーワードを比較（部分一致、`vbTextCompare`）。いずれにも一致しなければ `False`。
        *   デバッグログに判定過程と結果を出力。

    *   詳細な日本語コメントを各フィルターロジックに記述してください。

**生成コードに関する期待:**
*   `M02_ConfigReader` が「Configシート定義」Dセクションの全てのフィルター条件を正しく読み込み、`g_configSettings` の対応するメンバー（必要に応じて新規追加）に格納できること。
*   `M06_DataExtractor` の `PerformFilterCheck` 関数が、上記で段階的に実装された各フィルターロジックを正しく実行し、総合的な合致判定（全ての有効なフィルターをAND条件で満たすか）を行えること。
*   フィルターリストが空、またはフィルター条件セルが空の場合、そのフィルターは適用されず、常に「合致」として扱われること。
*   各フィルターの比較ロジック（完全一致/部分一致、大文字小文字の区別、AND/ORの扱い）が、「Configシート定義」および「仕様書」の指示通りであること。
*   デバッグモード時には、どのフィルターがどのように判定されたかの詳細がイミディエイトウィンドウに出力されること。
*   「System Instructions」のコーディング規約を遵守すること。

**成果物:**
上記の指示に基づき、Dセクションの読み込みが完全実装された `M02_ConfigReader` モジュール、および `PerformFilterCheck` 関数が段階的に本格実装された `M06_DataExtractor` モジュールのVBAコード。そして、必要に応じて更新された `M00_GlobalDeclarations`（`tConfigSettings`型定義の拡張）。
