# PowerPoint MCP Server

PowerPointファイルから構造化された情報を抽出するためのModel Context Protocol (MCP) サーバーです。

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

### スタンドアロン使用の場合

1. リポジトリをクローン：
```bash
git clone <repository-url>
cd powerpoint-analyzer
```

2. 依存関係をインストール：
```bash
pip install -r requirements.txt
```

3. パッケージをインストール：
```bash
pip install -e .
```

### AIエージェント統合の場合

1. 上記のスタンドアロンインストール手順を実行

2. インストールディレクトリの絶対パスを確認：
```bash
pwd
# 出力例: /Users/username/powerpoint-analyzer
```

3. ステップ2のパスを使用してAIエージェントを設定（下記の[AIエージェント統合](#aiエージェント統合)セクションを参照）

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

またはインストールされたコンソールスクリプトを使用：
```bash
powerpoint-mcp-server
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

## AIエージェント統合

このMCPサーバーは、Claude Desktop、Claude Code、その他のMCP対応アプリケーションなど、Model Context ProtocolをサポートするAIエージェントと統合できます。

### Claude Desktop設定

Claude Desktopの`mcp_settings.json`ファイルに以下の設定を追加してください：

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
        "POWERPOINT_MCP_LOG_LEVEL": "INFO",
        "POWERPOINT_MCP_MAX_FILE_SIZE": "100",
        "POWERPOINT_MCP_CACHE_ENABLED": "true"
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

### Claude Code設定

Claude Codeの場合、`mcp_settings.json`を作成または更新してください：

```json
{
  "mcpServers": {
    "powerpoint-analyzer": {
      "command": "python",
      "args": ["/absolute/path/to/powerpoint-analyzer/main.py"],
      "env": {
        "POWERPOINT_MCP_LOG_LEVEL": "DEBUG",
        "POWERPOINT_MCP_DEBUG": "true"
      }
    }
  }
}
```

### 代替案: 起動スクリプトの使用

より多くの設定オプションのために拡張起動スクリプトを使用することもできます：

```json
{
  "mcpServers": {
    "powerpoint-mcp-server": {
      "command": "python",
      "args": [
        "/path/to/your/powerpoint-analyzer/scripts/start_server.py",
        "--log-level", "INFO",
        "--max-file-size", "150"
      ]
    }
  }
}
```

### 設定オプション

AIエージェントと統合する際、環境変数を使用してサーバーの動作をカスタマイズできます：

| 環境変数 | デフォルト | 説明 |
|---------|-----------|------|
| `POWERPOINT_MCP_LOG_LEVEL` | `INFO` | ログレベル (DEBUG, INFO, WARNING, ERROR, CRITICAL) |
| `POWERPOINT_MCP_MAX_FILE_SIZE` | `100` | 最大ファイルサイズ（MB） |
| `POWERPOINT_MCP_TIMEOUT` | `300` | 処理タイムアウト（秒） |
| `POWERPOINT_MCP_CACHE_ENABLED` | `true` | キャッシュの有効/無効 |
| `POWERPOINT_MCP_DEBUG` | `false` | デバッグモードの有効化 |

### AIエージェントでの使用例

設定後、AIエージェントに以下のような依頼ができます：

1. **プレゼンテーション構造の分析**:
   ```
   /Users/username/Documents/quarterly-report.pptxのプレゼンテーションからコンテンツを抽出してください
   ```

2. **特定のスライド情報の取得**:
   ```
   /Users/username/Documents/quarterly-report.pptxのスライド3からタイトルとテキストコンテンツを取得できますか？
   ```

3. **特定の属性の抽出**:
   ```
   /Users/username/Documents/quarterly-report.pptxの全スライドからタイトルとオブジェクト数のみを取得してください
   ```

4. **プレゼンテーションメタデータの分析**:
   ```
   /Users/username/Documents/quarterly-report.pptxのスライドサイズと総スライド数は何ですか？
   ```

5. **テーブルと構造化データの抽出**:
   ```
   /Users/username/Documents/quarterly-report.pptxからテーブルを抽出して、その内容を表示してください
   ```

6. **テキスト書式の分析**:
   ```
   /Users/username/Documents/quarterly-report.pptxから太字テキストを抽出して、どのスライドに表示されているか教えてください
   ```

7. **書式サマリーの取得**:
   ```
   /Users/username/Documents/quarterly-report.pptxのテキスト書式（太字、斜体、下線など）を分析できますか？
   ```

8. **特定の書式タイプの抽出**:
   ```
   /Users/username/Documents/quarterly-report.pptxから太字と斜体の書式が適用されたテキストを抽出してください
   ```

### AIエージェント統合のトラブルシューティング

1. **サーバーが起動しない**: `mcp_settings.json`のパスが絶対パスで正しいことを確認
2. **権限エラー**: Python実行ファイルとスクリプトファイルに適切な権限があることを確認
3. **ファイルアクセスの問題**: AIエージェントが分析したいPowerPointファイルにアクセスできることを確認
4. **デバッグ情報**: 詳細なログのために`POWERPOINT_MCP_DEBUG=true`を設定

### ヘルスチェック

AIエージェントで設定する前に、サーバーが正しく動作することを確認してください：

```bash
python scripts/health_check.py
```

これにより、すべての依存関係と設定が検証されます。

### AIエージェント統合の確認

AIエージェントを設定した後：

1. **AIエージェントを再起動**（Claude Desktop、Claude Codeなど）

2. **サーバーが認識されているか確認**: AIエージェントに質問してください：
   ```
   利用可能なMCPツールは何ですか？
   ```
   PowerPoint MCPサーバーツールがリストに表示されるはずです。

3. **サンプルファイルでテスト**: PowerPointファイルからコンテンツを抽出して、すべてが正しく動作することを確認してください。

4. **ログを確認**: 問題が発生した場合は、AIエージェントのログを確認するか、デバッグモードを有効にしてください：
   ```json
   "env": {
     "POWERPOINT_MCP_LOG_LEVEL": "DEBUG",
     "POWERPOINT_MCP_DEBUG": "true"
   }
   ```

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