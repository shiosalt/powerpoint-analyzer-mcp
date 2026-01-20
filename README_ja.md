# PowerPoint Analyzer MCP Server
Model Context Protocol (MCP) を使用して、AIエージェントがPowerPoint (.pptx) ファイルから構造化されたコンテンツとテキストフォーマットを抽出できるようにするMCPサーバーです。

## 背景
PowerPointをサポートすると謳うほとんどのAIツールは、プレーンテキストのみを抽出し、構造とフォーマット情報を失っています。このMCPサーバーは、PowerPointの構造、フォーマット属性を保持し、プレゼンテーションコンテンツのクエリを可能にします。

## 機能

- **テキストフォーマット検出**: 太字、斜体、下線、取り消し線、ハイライト、ハイパーリンク
- **フォント分析**: フォント色
- **スライドクエリ**: タイトル、コンテンツ、レイアウト、スピーカーノートによるフィルタリング
- **テーブルデータ抽出**: 複数の出力形式（row/col/value、HTML）とオプションのフォーマット
- **Pythonスタイルのスライド選択**: スライス記法（`:10`、`5:20`、`25:`）
- **コンテキスト消費を削減**: 構造化されたデータ出力
- **外部依存なし**: PowerPoint処理にPython標準ライブラリのみを使用
- **FastMCP 2.0で構築**: MCPサーバーフレームワーク

## プロジェクト構造

```
powerpoint-analyzer/
├── main.py                     # FastMCPサーバーのメインエントリーポイント
├── powerpoint_mcp_server/      # コアサーバー実装
│   ├── server.py              # メインMCPサーバー実装
│   ├── config.py              # 設定管理
│   ├── core/                  # コア機能
│   └── utils/                 # ユーティリティモジュール
├── tests/                      # テストファイル
│   ├── test_powerpoint_fastmcp.py  # メインサーバーテスト
│   ├── test_formatting_detection.py # フォーマット検出テスト
│   └── ...                         # その他のテストファイル
├── scripts/                    # ユーティリティスクリプト
│   ├── health_check.py        # サーバーヘルスチェック
│   └── start_server.py        # 代替サーバー起動
├── requirements.txt            # Python依存関係
├── pytest.ini                 # テスト設定
└── README.md                   # ドキュメント
```

## インストール

1. リポジトリをクローン:
```bash
git clone <repository-url>
cd powerpoint-analyzer
```

2. 依存関係をインストール:
```bash
pip install -r requirements.txt
```

3. MCPクライアント（Claude Desktop、Clineなど）を設定:

**Claude Desktopの場合:**

設定ファイルの場所: `~/Library/Application Support/Claude/claude_desktop_config.json` (macOS) または `%APPDATA%\Claude\claude_desktop_config.json` (Windows)

```json
{
  "mcpServers": {
    "powerpoint-analyzer": {
      "command": "python",
      "args": ["/absolute/path/to/powerpoint-analyzer/main.py"],
      "env": {
        "POWERPOINT_MCP_LOG_LEVEL": "INFO"
      }
    }
  }
}
```

**Cline/その他のMCPクライアントの場合:**

MCP設定ファイル（通常は `.kiro/settings/mcp.json` など）に追加:

```json
{
  "mcpServers": {
    "powerpoint-analyzer": {
      "command": "python",
      "args": ["C:\\path\\to\\powerpoint-analyzer\\main.py"],
      "env": {
        "POWERPOINT_MCP_LOG_LEVEL": "INFO"
      }
    }
  }
}
```

**注意**: `args` パラメータには絶対パスを使用してください。Windowsでは二重バックスラッシュ（`\\`）またはスラッシュ（`/`）を使用します。

## 技術的アプローチ

このサーバーは以下のアプローチでPowerPointファイルを処理します:

- **直接ZIP処理**: .pptxファイルはPythonの `zipfile` モジュールを使用してZIPアーカイブとして処理
- **XML解析**: PowerPoint内部のXML構造を `xml.etree.ElementTree` で名前空間サポート付きで解析
- **デュアルフォーマット検出**: テキストフォーマットプロパティのXML属性と子要素の両方の形式をサポート
- **外部依存なし**: PowerPoint処理にPython標準ライブラリモジュールのみを使用
- **処理**: プレゼンテーション全体をメモリにロードせず、必要な情報のみを抽出
- **キャッシング**: 繰り返し操作のパフォーマンス向上のためのキャッシングシステム

## テキストフォーマット検出

サーバーはテキストフォーマット検出機能を提供します:

### サポートされているフォーマットタイプ
- **太字テキスト**: テキスト要素の太字フォーマットを検出
- **斜体テキスト**: スライド全体の斜体スタイルを識別
- **下線テキスト**: 下線スタイル付きの下線テキストを検出
- **取り消し線テキスト**: 取り消し線フォーマットを検出
- **ハイライトテキスト**: ハイライト/背景色付きテキストを識別
- **ハイパーリンク**: ハイパーリンク情報とリレーションシップIDを抽出
- **フォントプロパティ**: フォントサイズ、色（RGBとスキームカラー）を分析

### 技術的実装
- **デュアル検出方式**: XML属性（`b="1"`）と子要素（`<a:b val="1"/>`）の両方をチェック
- **名前空間対応解析**: Office Open XML名前空間の処理
- **デバッグ機能**: フォーマット検出の問題をトラブルシューティングするためのデバッグツール

### テスト
- **テストスイート**: `tests/test_formatting_detection.py` がフォーマットタイプを検証
- **デバッグツール**: `tests/debug_formatting_detection.py` がXML分析を提供

## 使用方法

### サーバーの起動

```bash
python main.py
```

## 利用可能なMCPツール

このサーバーはPowerPoint分析のための4つのコアツールを提供します:

**推奨ツール（AIエージェント向け）:**
- `query_slides` - コンテキスト消費を削減したスライドクエリ
- `extract_table_data` - コンテキスト消費を削減したテーブル抽出

**レガシーツール（非推奨 - 大量のコンテキスト消費）:**
- `extract_formatted_table_data` - 完全なフォーマットメタデータ（必要な場合のみ使用）
- `extract_formatted_text` - 詳細なフォーマット分析（必要な場合のみ使用）

### 1. query_slides

指定された条件でスライドをクエリおよびフィルタリング。コンテキスト消費を削減した構造化スライド情報を返します。

**パラメータ:**

| パラメータ | 型 | 必須 | 説明 |
|-----------|------|----------|-------------|
| `file_path` | string | はい | PowerPointファイル（.pptx）への完全パス |
| `search_criteria` | object | はい | フィルタリング条件（下記参照） |
| `return_fields` | array | いいえ | 結果に含めるフィールド（デフォルト: `["slide_number", "title", "text"]`） |
| `slide_numbers` | int/string/array | いいえ | クエリするスライド（デフォルト: 全スライド） |
| `output_type` | string | いいえ | テキスト出力タイプ: `"preview_text_3boxes"` (デフォルト) または `"full_text"` |
| `output_format` | string | いいえ | 出力形式: `"simple"` (デフォルト) または `"formatted"` |
| `limit` | integer | いいえ | 返す最大結果数（デフォルト: 1000、最大: 10000） |

**search_criteria 構造:**

```json
{
  "title": {
    "contains": "Sales",           // タイトルにテキストを含む
    "starts_with": "Chapter",      // タイトルがテキストで始まる
    "ends_with": "Summary",        // タイトルがテキストで終わる
    "regex": "^Q[1-4] 202[0-9]$", // タイトルが正規表現にマッチ
    "one_of": ["Intro", "Outro"]  // タイトルがこれらの値のいずれか
  },
  "content": {
    "contains_text": "revenue",    // スライドテキストに文字列を含む
    "has_tables": true,            // スライドにテーブルがある
    "has_charts": true,            // スライドにチャートがある
    "has_images": true             // スライドに画像がある
  },
  "notes": {
    "contains": "important",       // スピーカーノートにテキストを含む
    "regex": "TODO.*",             // ノートが正規表現にマッチ
    "is_empty": false              // ノートが空でない
  },
  "sections": ["Introduction", "Conclusion"]  // セクション名でフィルタ
}
```

**return_fields オプション:**

- `"slide_number"`: スライド番号（常に含まれる）
- `"title"`: スライドタイトル
- `"subtitle"`: スライドサブタイトル
- `"text"`: テキストコンテンツ（形式は `output_type` で制御）
- `"extracted_tables"`: 簡素化形式のテーブルデータ

**output_type オプション:**

- `"preview_text_3boxes"` (デフォルト): タイトル + コンテンツプレースホルダー + 最大3つのテキストボックスを表示
- `"full_text"`: 制限なしで全テキスト要素を表示

**output_format オプション:**

- `"simple"` (デフォルト): テキスト/テーブルにフォーマット情報なし
- `"formatted"`: フォーマット情報を含む（太字、斜体、色など）

**使用例:**

```python
# タイトルに"Sales"を含むスライドを検索
query_slides("C:\\temp\\presentation.pptx", {"title": {"contains": "Sales"}})

# テーブルを含むスライドを検索、全テキストを返す
query_slides("C:\\temp\\presentation.pptx", 
            {"content": {"has_tables": true}},
            output_type="full_text")

# カスタムフィールドで特定のスライドをクエリ
query_slides("C:\\temp\\presentation.pptx", {},
            return_fields=["slide_number", "title", "extracted_tables"],
            slide_numbers="1,5,10")
```

---

### 2. extract_table_data

コンテキスト消費を削減した簡素化形式でテーブルデータを抽出。

**パラメータ:**

| パラメータ | 型 | 必須 | 説明 |
|-----------|------|----------|-------------|
| `file_path` | string | はい | PowerPointファイル（.pptx）へのパス |
| `slide_numbers` | int/string/array | いいえ | 抽出元のスライド（デフォルト: 全スライド） |
| `column_selection` | object | いいえ | カラムフィルタリング設定 |
| `output_format` | string | いいえ | 出力形式（デフォルト: `"row_col_value"`） |

**output_format オプション:**

- `"row_col_value"` (デフォルト): 値のみの `[row, col, value]` 形式
- `"row_col_formattedvalue"`: フォーマットを含む `[row, col, value]` 形式
- `"html"`: フォーマット付きHTMLテーブル（colspan/rowspanサポート）
- `"simple_html"`: フォーマットなしHTMLテーブル（colspan/rowspanサポート）

**column_selection 構造:**

```json
{
  "specific_columns": ["Name", "Price", "Quantity"],  // 名前で特定のカラムを選択
  "column_patterns": [".*_total$", "^sum_.*"],       // 正規表現パターンにマッチするカラムを選択
  "exclude_columns": ["Notes", "Internal_ID"],       // 特定のカラムを除外
  "all_columns": true                                 // 全カラムを含む（デフォルト）
}
```

**出力構造:**

`row_col_value` / `row_col_formattedvalue` の場合:
```json
{
  "extracted_tables": [
    {
      "slide_number": 3,
      "rows": 5,
      "columns": 3,
      "headers": ["Product", "Price", "Quantity"],
      "data": [
        [1, 0, "Widget A"],
        [1, 1, "$10.00"],
        [1, 2, "100"],
        [2, 0, "Widget B"],
        [2, 1, "$15.00"],
        [2, 2, "50"]
      ]
    }
  ]
}
```

`html` / `simple_html` の場合:
```json
{
  "extracted_html_tables": [
    {
      "slide_number": 3,
      "rows": 5,
      "columns": 3,
      "headers": ["Product", "Price", "Quantity"],
      "htmldata": "<table style=\"white-space: pre;\">...</table>"
    }
  ]
}
```

**使用例:**

```python
# 全テーブルを簡単な配列として抽出
extract_table_data("C:\\temp\\presentation.pptx")

# フォーマット付きHTMLテーブルとして抽出
extract_table_data("C:\\temp\\presentation.pptx", output_format="html")

# 特定のスライドのみから抽出
extract_table_data("C:\\temp\\presentation.pptx", slide_numbers=[1, 3, 5])

# 特定のカラムを抽出
extract_table_data("C:\\temp\\presentation.pptx",
                  column_selection={"specific_columns": ["Name", "Total"]})
```

---

## 使用推奨事項

### AIエージェント向け（推奨）

コンテキスト消費を削減するために以下のツールを使用してください:

1. **query_slides**: スライド分析の主要ツール
   - コンテキスト消費を削減
   - 複数のフィルタリングオプション
   - 構造化された出力
   - フォーマット情報が必要な場合は `output_format="formatted"` を使用

2. **extract_table_data**: テーブル抽出の主要ツール
   - コンテキスト消費を削減
   - 複数の形式オプション（row/col/value、HTML）
   - 大規模プレゼンテーションに適用可能

### レガシーツール（控えめに使用）

これらのツールは大量のコンテキスト出力を生成するため、絶対に必要な場合のみ使用してください:

- **extract_formatted_table_data**: 詳細なフォーマットメタデータ分析のみ
- **extract_formatted_text**: 包括的なフォーマット調査のみ

ほとんどのフォーマットニーズには、代わりに `query_slides` を `output_format="formatted"` で使用してください。

---

### 3. extract_formatted_table_data

⚠️ **非推奨**: このツールは広範なフォーマットメタデータを含む非常に大きなコンテキスト出力を生成します。ほとんどのユースケースでは代わりに `extract_table_data` を使用してください。

包括的なフォーマット検出を備えたテーブルデータ抽出（完全なフォーマットサポートを持つレガシーツール）。

**使用すべき場合:**
- 詳細なフォーマットメタデータ（太字、斜体、色など）が必要な場合のみ
- 特殊なフォーマット分析要件がある場合
- 追加のコンテキスト消費が許容できる場合

**ほとんどのユースケースでは、代わりに `extract_table_data` を使用してください。**

**パラメータ:**

| パラメータ | 型 | 必須 | 説明 |
|-----------|------|----------|-------------|
| `file_path` | string | はい | PowerPointファイル（.pptx）へのパス |
| `slide_numbers` | int/string/array | いいえ | 抽出元のスライド（デフォルト: 全スライド） |
| `table_criteria` | object | いいえ | テーブルフィルタリング条件 |
| `column_selection` | object | いいえ | カラムフィルタリング設定 |
| `formatting_detection` | object | いいえ | フォーマット検出設定 |
| `output_format` | string | いいえ | 出力形式: `"structured"`, `"flat"`, `"grouped_by_slide"` |
| `include_metadata` | boolean | いいえ | テーブルメタデータを含む（デフォルト: true） |

**table_criteria 構造:**

```json
{
  "min_rows": 2,                                    // 最小行数
  "max_rows": 100,                                  // 最大行数
  "min_columns": 2,                                 // 最小列数
  "max_columns": 10,                                // 最大列数
  "header_contains": ["Total", "Summary"],         // ヘッダーにこれらの文字列を含む必要がある
  "header_patterns": ["^Q[1-4].*", ".*_total$"]   // ヘッダーがこれらの正規表現パターンにマッチする必要がある
}
```

**formatting_detection 構造:**

```json
{
  "detect_bold": true,           // 太字テキストを検出
  "detect_italic": true,         // 斜体テキストを検出
  "detect_underline": true,      // 下線テキストを検出
  "detect_highlight": true,      // ハイライトテキストを検出
  "detect_colors": true,         // テキスト色を検出
  "detect_hyperlinks": true,     // ハイパーリンクを検出
  "preserve_formatting": true    // 出力でフォーマットを保持
}
```

**使用例:**

```python
# フォーマット検出付きで抽出
extract_formatted_table_data("C:\\temp\\presentation.pptx",
                            formatting_detection={
                              "detect_bold": true,
                              "detect_colors": true
                            })

# 特定の条件でテーブルを抽出
extract_formatted_table_data("C:\\temp\\presentation.pptx",
                            table_criteria={
                              "min_rows": 3,
                              "header_contains": ["Total"]
                            })
```

---

### 4. extract_formatted_text

⚠️ **非推奨**: このツールは詳細なフォーマット分析を含む非常に大きなコンテキスト出力を生成します。より軽量なフォーマット情報には `query_slides` を `output_format="formatted"` で使用することを検討してください。

特定のフォーマット属性を持つテキストをスライドから抽出。

**使用すべき場合:**
- 全スライドにわたる包括的なフォーマット分析が必要な場合のみ
- 特殊なテキストフォーマット調査の場合
- 追加のコンテキスト消費が許容できる場合

**ほとんどのユースケースでは、代わりに適切なフィルタを使用した `query_slides` を使用してください。**

**パラメータ:**

| パラメータ | 型 | 必須 | 説明 |
|-----------|------|----------|-------------|
| `file_path` | string | はい | PowerPointファイル（.pptx）へのパス |
| `formatting_type` | string | はい | 抽出するフォーマットのタイプ |
| `slide_numbers` | int/string/array | いいえ | 分析するスライド（デフォルト: 全スライド） |

**formatting_type オプション:**

- `"bold"`: 太字テキストセグメントを抽出
- `"italic"`: 斜体テキストセグメントを抽出
- `"underlined"`: 下線テキストセグメントを抽出
- `"highlighted"`: ハイライトテキストセグメントを抽出
- `"strikethrough"`: 取り消し線テキストセグメントを抽出
- `"hyperlinks"`: URLとリンクタイプ付きハイパーリンクを抽出
- `"font_sizes"`: フォントサイズ情報付きテキストを抽出
- `"font_colors"`: 色情報（16進形式）付きテキストを抽出

**出力構造:**

```json
{
  "file_path": "C:\\temp\\presentation.pptx",
  "formatting_type": "bold",
  "summary": {
    "total_slides_analyzed": 10,
    "slides_with_formatting": 5,
    "total_formatted_segments": 12
  },
  "results_by_slide": [
    {
      "slide_number": 1,
      "title": "Introduction",
      "complete_text": "Welcome to our presentation...",
      "format": "bold",
      "formatted_segments": [
        {
          "text": "Important Note",
          "start_position": 25
        }
      ]
    }
  ]
}
```

**使用例:**

```python
# 全ての太字テキストを抽出
extract_formatted_text("C:\\temp\\presentation.pptx", "bold")

# 特定のスライドからハイパーリンクを抽出
extract_formatted_text("C:\\temp\\presentation.pptx", "hyperlinks", slide_numbers=[1, 2, 3])

# 最初の10スライドからフォント色を抽出
extract_formatted_text("C:\\temp\\presentation.pptx", "font_colors", slide_numbers=":10")
```

---

## スライド選択構文

全てのツールはPythonスタイルのスライシングを使用したスライド選択をサポートします:

| 形式 | 例 | 説明 |
|--------|---------|-------------|
| None | (パラメータ省略) | 全スライド |
| Integer | `3` | スライド3のみ |
| Array | `[1, 5, 10]` | 特定のスライド1, 5, 10 |
| String (カンマ) | `"1,5,10"` | 特定のスライド1, 5, 10 |
| String (スライス) | `":10"` | 最初の10スライド（1-10） |
| String (スライス) | `"5:20"` | スライド5-20 |
| String (スライス) | `"25:"` | スライド25から最後まで |

**例:**

```python
# 最初の10スライド
query_slides("file.pptx", {}, slide_numbers=":10")

# スライド5-20
extract_table_data("file.pptx", slide_numbers="5:20")

# スライド25から最後まで
extract_formatted_text("file.pptx", "bold", slide_numbers="25:")

# 特定のスライド
query_slides("file.pptx", {}, slide_numbers="1,3,5,10")
```

## 開発

### 要件

- Python 3.8+
- MCP (Model Context Protocol)
- FastMCP 2.0
- 標準Pythonライブラリ（zipfile、xml.etree.ElementTree）

## 最近の更新

### Version 2.2 - ドキュメント強化とツール改善
- **詳細なツールドキュメント**: 全MCPツールのパラメータ説明
- **出力形式オプション**: テーブル用の複数形式（row/col/value、HTML、フォーマット付き）
- **検索条件**: スライドとテーブルの検索条件
- **カラム選択**: 名前またはパターンでテーブルカラムをフィルタ
- **コンテキスト消費を削減**: 構造化されたデータ出力

### Version 2.1 - Pythonスタイルのスライド選択
- **スライド選択の強化**: Pythonスタイルのスライス記法（`:10`、`5:20`、`25:`）
- **複数の指定方法**: 単一スライド、範囲、カンマ区切りリスト、スライシング
- **パフォーマンス向上**: 必要なスライドのみを処理
- **後方互換性**: 既存の `[1, 5, 10]` 形式も引き続きサポート

### Version 2.0 - テキストフォーマット検出
- **フォーマット検出の修正**: 太字、斜体、下線、取り消し線が正しく検出される
- **デュアル検出サポート**: XML属性と子要素の両方の形式を処理
- **テストスイート**: 全フォーマットタイプのテスト
- **MCPツールの強化**: フォーマット付きテキスト抽出と分析のための新しいツール

## ライセンス

このプロジェクトはApache License 2.0の下でライセンスされています - 詳細は[LICENSE](LICENSE)ファイルを参照してください。
