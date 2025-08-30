# PowerPoint MCP Server

PowerPointファイルから構造化された情報を抽出するためのModel Context Protocol (MCP) サーバーです。

## 機能

- PowerPoint (.pptx) ファイルから完全な構造化コンテンツを抽出
- スライドから特定の属性（タイトル、サブタイトル、テキスト、テーブル、画像など）を取得
- 個別のスライド情報を取得
- スライドレイアウト、プレースホルダー、フォーマット情報をサポート
- Python標準ライブラリを使用した軽量実装（外部PowerPoint依存関係なし）
- 高速で効率的な処理のための直接XML解析

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

このサーバーはPowerPoint処理に軽量なアプローチを使用しています：

- **直接ZIP処理**: .pptxファイルをPythonの`zipfile`モジュールを使用してZIPアーカイブとして処理
- **XML解析**: 内部PowerPoint XML構造を`xml.etree.ElementTree`を使用して解析
- **外部依存関係なし**: PowerPoint処理にPython標準ライブラリモジュールのみを使用
- **効率的な処理**: プレゼンテーション全体をメモリに読み込むことなく、必要な情報のみを抽出

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

1. **extract_powerpoint_content**: PowerPointファイルから完全な構造化コンテンツを抽出
2. **get_powerpoint_attributes**: PowerPointスライドから特定の属性を取得
3. **get_slide_info**: 特定のスライドの情報を取得

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
   /Users/username/Documents/quarterly-report.pptxのプレゼンテーションから完全なコンテンツを抽出してください
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
   /Users/username/Documents/quarterly-report.pptxから全てのテーブルを抽出して、その内容を表示してください
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
│   ├── content_extractor.py    # PowerPointコンテンツ抽出
│   ├── attribute_processor.py  # 属性フィルタリングと処理
│   └── xml_parser.py           # XML解析ユーティリティ
└── utils/
    ├── __init__.py
    ├── file_validator.py       # ファイル検証
    ├── zip_extractor.py        # ZIPアーカイブ処理
    └── cache_manager.py        # キャッシュユーティリティ
```

### 要件

- Python 3.8+
- MCP (Model Context Protocol)
- Python標準ライブラリ (zipfile, xml.etree.ElementTree)

## ライセンス

TBD