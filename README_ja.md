# PowerPoint analyzer MCP Server
PowerPointの構造やテキストの書式属性を利用した検索・抽出が可能なMCPサーバーです。

## 背景
PowerPoint対応を謳うAI Agent検索は、PowerPointファイルの構造化を無視しテキストのみ抽出して検索するものが一般的で、定型のPowerPoint資料から強調された文字を検索するなどはできません。
このMCPサーバを利用すれば、PowerPointの構造（セクション、スライドタイトル、テーブル、メモ）やテキスト属性（太字、斜体、下線、取り消し線、ハイライト、ハイパーリンク）を検索の条件に指定できます。
Claude Sonnet 3.5 相当など高度なAIを利用している場合、AIとMCP間で自動でラリーを行い情報を分析して探すこともできます。

## 機能

- **テキスト書式検出**: 太字、斜体、下線、取り消し線、ハイライト、ハイパーリンクの検出と抽出
- **フォント分析**: フォントサイズ、フォントの色、スタイル情報の分析
- **スライド検索**: 柔軟なフィルタリング条件でスライドをクエリ・検索
- **テーブルデータ抽出**: フォーマット検出付きテーブルデータ抽出
- **テストスイート** による書式検出の検証
- Python標準ライブラリを使用した実装（外部PowerPoint依存関係なし）
- 直接XML解析による処理
- FastMCP 2.0利用

## プロジェクト構造

```
powerpoint-analyzer/
├── main.py                     # メインFastMCPサーバーエントリーポイント
├── powerpoint_mcp_server/      # コアサーバー実装
│   ├── server.py              # メインMCPサーバー実装
│   ├── config.py              # 設定管理
│   ├── core/                  # コア機能
│   └── utils/                 # ユーティリティモジュール
├── tests/                      # テストファイル
│   ├── test_powerpoint_fastmcp.py  # メインサーバーテスト
│   ├── test_formatting_detection.py # 書式検出テスト
│   └── ...                         # その他のテストファイル
├── scripts/                    # ユーティリティスクリプト
│   ├── health_check.py        # サーバーヘルスチェック
│   └── start_server.py        # 代替サーバー起動
├── requirements.txt            # Python依存関係
├── pytest.ini                 # テスト設定
└── README.md                   # ドキュメント
```

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

このMCPサーバーは以下の3つの主要ツールを提供します：

1. **extract_formatted_text**: 特定の書式タイプ（太字、斜体、下線、取り消し線、ハイライト、ハイパーリンク、フォントサイズ、フォント色）のテキストを抽出
2. **query_slides**: 柔軟なフィルタリング条件でスライドをクエリ・検索
3. **extract_table_data**: 選択とフォーマット検出でテーブルデータを抽出



## 開発

### 要件

- Python 3.8+
- MCP (Model Context Protocol)
- FastMCP 2.0
- Python標準ライブラリ (zipfile, xml.etree.ElementTree)

## 最新の更新

### テキスト書式検出
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

このプロジェクトはApache License 2.0の下でライセンスされています - 詳細は[LICENSE](LICENSE)ファイルを参照してください。