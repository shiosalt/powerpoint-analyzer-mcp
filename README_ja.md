# PowerPoint analyze MCP Server
PowerPointの構造やテキストの書式属性を利用した検索・抽出が可能なMCPサーバーです。

## 背景
PowerPoint対応を謳うAI Agent検索が、PowerPointファイルの構造化を無視しテキストのみ抽出して検索するものが一般的です。
太字で記載したテキストを出力するなどを可能にしました。

## 機能

- PowerPoint (.pptx) ファイルから構造化コンテンツを抽出
- スライドから特定の属性（タイトル、サブタイトル、テキスト、テーブル、画像など）を取得
- 個別のスライド情報を取得
- フィルタリング条件でスライドをクエリ
- フォーマット検出付きテーブルデータ抽出
- **テキスト書式検出**: 太字、斜体、下線、取り消し線、ハイライト、ハイパーリンク
- **フォント分析**: フォントサイズ、フォントの色、スタイル情報
- プレゼンテーション概要と分析を取得
- スライドレイアウト、プレースホルダー、フォーマット情報をサポート
- **テストスイート** による書式検出の検証
- **デバッグツール** による書式検出問題のトラブルシューティング
- Python標準ライブラリを使用した実装（外部PowerPoint依存関係なし）
- 直接XML解析による処理
- FastMCP 2.0対応

## インストール

1. リポジトリをクローン：
```bash
git clone <repository-url>
cd powerpoint-analyzer
```

2. 依存関係をインストール：
```bash
pip install -r requirements.txt
```

3. AIエージェント（Claude Desktop等）の設定ファイルに以下を追加：

**mcp_settings.jsonの場所:**
- **macOS**: `~/Library/Application Support/Claude/mcp_settings.json`
- **Windows**: `%APPDATA%\Claude\mcp_settings.json`

```json
{
  "mcpServers": {
    "powerpoint-mcp-server": {
      "command": "python",
      "args": ["/path/to/your/powerpoint-analyzer/main.py"],
      "env": {
        "POWERPOINT_MCP_LOG_LEVEL": "INFO"
      }
    }
  }
}
```

**実際のパスを使用した例:**

macOS/Linux:
```json
{
  "mcpServers": {
    "powerpoint-mcp-server": {
      "command": "python",
      "args": ["/Users/username/powerpoint-analyzer/main.py"],
      "env": {
        "POWERPOINT_MCP_LOG_LEVEL": "INFO"
      }
    }
  }
}
```

Windows:
```json
{
  "mcpServers": {
    "powerpoint-mcp-server": {
      "command": "python",
      "args": ["C:\\Users\\username\\powerpoint-analyzer\\main.py"],
      "env": {
        "POWERPOINT_MCP_LOG_LEVEL": "INFO"
      }
    }
  }
}
```

## 技術的アプローチ

このサーバーはPowerPoint処理に以下のアプローチを使用しています：

- **直接ZIP処理**: .pptxファイルをPythonの`zipfile`モジュールを使用してZIPアーカイブとして処理
- **XML解析**: 名前空間サポートで`xml.etree.ElementTree`を使用して内部PowerPoint XML構造を解析
- **デュアル書式検出**: テキスト書式プロパティのXML属性と子要素の両方の形式をサポート
- **外部依存関係なし**: PowerPoint処理にPython標準ライブラリモジュールのみを使用
- **処理**: プレゼンテーション全体をメモリに読み込むことなく、必要な情報のみを抽出
- **キャッシュ**: 繰り返し操作でのパフォーマンス向上のためのキャッシュシステム

## テキスト書式検出

サーバーはテキスト書式検出機能を提供します：

### サポートされる書式タイプ
- **太字テキスト**: テキスト要素で太字書式を検出
- **斜体テキスト**: スライド全体で斜体スタイルを識別
- **下線テキスト**: 下線スタイルで下線テキストを検出
- **取り消し線テキスト**: 取り消し線書式を検出
- **ハイライトテキスト**: ハイライト/背景色付きテキストを識別
- **ハイパーリンク**: ハイパーリンク情報と関係IDを抽出
- **フォントプロパティ**: フォントサイズ、フォントの色（RGBとスキーム色）を分析

### 技術実装
- **デュアル検出方式**: XML属性（`b="1"`）と子要素（`<a:b val="1"/>`）の両方をチェック
- **名前空間対応解析**: Office Open XML名前空間の処理
- **検証**: 異なるPowerPointバージョンでの検出のためのテストスイート
- **デバッグ機能**: 書式検出問題のトラブルシューティング用のデバッグツール

### 検証とテスト
- **テストスイート**: `tests/test_formatting_detection.py`で書式タイプを検証
- **デバッグツール**: `tests/debug_formatting_detection.py`でXML分析を提供
- **検証**: 混合書式を含むPowerPointファイルでテスト済み

## 使用方法

### サーバーの実行

```bash
python main.py
```

### 利用可能なツール

#### コアコンテンツ抽出
1. **extract_powerpoint_content**: PowerPointファイルから構造化コンテンツを抽出
2. **get_powerpoint_attributes**: PowerPointスライドから特定の属性を取得
3. **get_slide_info**: 特定のスライドの情報を取得
4. **query_slides**: フィルタリング条件でスライドをクエリ
5. **extract_table_data**: 選択とフォーマット検出でテーブルデータを抽出

#### テキスト書式分析
6. **extract_bold_text**: 位置情報付きでスライドから太字テキストを抽出
7. **extract_formatted_text**: 特定の書式タイプ（太字、斜体、下線、取り消し線、ハイライト、ハイパーリンク）のテキストを抽出
8. **get_formatting_summary**: プレゼンテーション内のテキスト書式のサマリーを取得
9. **analyze_text_formatting**: スライド全体のテキスト書式パターンを分析

#### プレゼンテーション分析
10. **get_presentation_overview**: プレゼンテーション概要と分析を取得
11. **clear_cache**: キャッシュをクリア
12. **reload_file_content**: キャッシュをクリアして再抽出によりファイルコンテンツを再読み込み



## 開発

### プロジェクト構造

```
powerpoint_mcp_server/
├── __init__.py
├── server.py              # メインMCPサーバー実装
├── config.py              # 設定管理
├── core/
│   ├── __init__.py
│   ├── content_extractor.py    # 書式検出付きPowerPointコンテンツ抽出
│   ├── attribute_processor.py  # 属性フィルタリングと処理
│   ├── presentation_analyzer.py # プレゼンテーション分析
│   └── xml_parser.py           # XML解析ユーティリティ
├── utils/
│   ├── __init__.py
│   ├── file_validator.py       # ファイル検証
│   ├── zip_extractor.py        # ZIPアーカイブ処理
│   └── cache_manager.py        # キャッシュユーティリティ
└── tests/
    ├── test_formatting_detection.py  # 書式検出テスト
    ├── debug_formatting_detection.py # 書式問題用デバッグツール
    └── ...                           # その他のテストファイル
```

### 要件

- Python 3.8+
- MCP (Model Context Protocol)
- FastMCP 2.0
- Python標準ライブラリ (zipfile, xml.etree.ElementTree)

## 最新の更新

### バージョン 2.0 - テキスト書式検出
- **書式検出バグを修正**: 太字、斜体、下線、取り消し線属性が正しく検出されるようになりました
- **デュアル検出サポート**: XML属性と子要素の両方の形式を処理
- **テストスイート**: 書式タイプのテストを追加
- **デバッグツール**: 書式問題のトラブルシューティング用のデバッグユーティリティ
- **MCPツール**: 書式付きテキスト抽出と分析のためのツール

### 対応している書式タイプ
下記の書式タイプを確認済み：
- ✅ **太字テキスト**
- ✅ **斜体テキスト**
- ✅ **下線テキスト**
- ✅ **取り消し線テキスト**
- ✅ **ハイライトテキスト**
- ✅ **ハイパーリンク**
- ✅ **フォントの色**

## ライセンス

TBD