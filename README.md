# vba-schedule-aggregator
----
# Directory memo
vba-schedule-aggregator/               # リポジトリルート
├── .gitignore                         # Gitで追跡しないファイルを指定
├── README.md                          # プロジェクトの説明、フォルダ構成の日本語解説など
│
├── src/                               # ソースコード (マクロ本体)
│   └── schedule_extractor.xlsm        # マクロが含まれるExcelファイル (旧: Excelファイル抽出ソフト.xlsm)
│
├── docs/                              # ドキュメント類
│   ├── specification.md               # 仕様書 (旧: 仕様書.md)
│   ├── config_sheet_definition.md     # Configシート定義 (旧: Configシート定義.md)
│   ├── expected_behavior.md           # 期待する動作 (旧: 期待する動作 (プログラムの挙動).md)
│   ├── System Instructions.md         # AIへの指示・補足など (旧: System Instructions (AIへの役割指示 - 更新版))
│   ├── development_notes.md           # AIへの指示・補足など (旧: AI関連ファイルを統合)
│   └── prompt.md                      # AIへの指示・補足など (旧: プロンプト (AIへの具体的な指示))
│
└── samples/                           # 参考ファイル、サンプルデータ
    ├── examples/                      # マクロ本体のExcelシートのサンプル
    │   ├── config_sheet_example.csv   # ConfigシートのCSVサンプル (旧: Excelファイル抽出ソフト_Configシート.csv)
    │   └── work_sheet_example.csv     # データベースサンプル (旧: Excelファイル抽出ソフト_Workシート.csv が出力例と仮定)
    ├── input_excel/                   # 入力となるExcelファイルのサンプル
    │   ├── monthly_schedule_202504_data_example.csv # 抽出元CSV (旧: 抽出元_2025.04 月間予定表.csv)
    │   ├── source_schedule_202504_screenshot.png    # 抽出元カレンダーのスクリーンショット (旧: 抽出元_2025.04 月間予定表.png)
    │   └── monthly_schedule_202504_example.xlsx     # 抽出元Excel (旧: 抽出元_2025.04 月間予定表.xlsx)
    └── output_examples/               # 出力結果のサンプル
        └── output_example.csv         # 出力結果のサンプル
