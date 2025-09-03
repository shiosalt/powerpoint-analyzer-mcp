# MCP Test CLI Tool

PowerPoint Analyzer MCP用のコマンドラインテストツールです。MCPサーバーとstdio通信を行い、ツールの動作確認やデバッグを簡単に行えます。

## 📋 ファイル構成

- `mcp_test_cli.py` - メインのMCPテストCLIツール
- `test_tools.py` - よく使用するテストシナリオの簡易ラッパー
- `examples/test_examples.py` - 使用例を示すサンプルスクリプト
- `MCP_TEST_CLI_README.md` - このドキュメント

## 🚀 基本的な使用方法

### 1. 利用可能なツール一覧を表示

```bash
python mcp_test_cli.py
```

出力例：
```
📋 Available Tools (15 total):
==================================================

 1. extract_powerpoint_content
    Extract complete structured content from a PowerPoint file.

 2. get_powerpoint_attributes
    Get specific attributes from PowerPoint slides with selective extraction.

 3. extract_bold_text
    Extract bold text from slides with location information.
...
```

### 2. 特定ツールのヘルプを表示

```bash
python mcp_test_cli.py <tool_name>
```

例：
```bash
python mcp_test_cli.py extract_powerpoint_content
```

出力例：
```
🔧 Tool: extract_powerpoint_content
====================================
Description: Extract complete structured content from a PowerPoint file.

Parameters:
  --file_path <string> (required)
      Path to the PowerPoint file (.pptx). Must be a valid PowerPoint file.

Usage Example:
  python mcp_test_cli.py extract_powerpoint_content --file_path "example_value"
```

### 3. ツールを実行

```bash
python mcp_test_cli.py <tool_name> [options]
```

例：
```bash
python mcp_test_cli.py extract_powerpoint_content --file_path "tests/test_files/test_minimal.pptx"
```

## 🎯 簡易ラッパー (test_tools.py)

よく使用するテストシナリオ用の簡易コマンド：

```bash
# ツール一覧表示
python test_tools.py list

# ツールヘルプ表示
python test_tools.py help extract_powerpoint_content

# PowerPoint内容抽出
python test_tools.py extract tests/test_files/test_minimal.pptx

# 特定属性取得
python test_tools.py attrs tests/test_files/test_minimal.pptx title subtitle

# スライド情報取得
python test_tools.py slide tests/test_files/test_minimal.pptx 1

# 太字テキスト抽出
python test_tools.py bold tests/test_files/test_complex.pptx

# 特定フォーマット抽出
python test_tools.py format tests/test_files/test_complex.pptx italic
```

## 📚 使用例

### PowerPoint内容の完全抽出

```bash
python mcp_test_cli.py extract_powerpoint_content --file_path "presentation.pptx"
```

### 特定属性のみ取得

```bash
# カンマ区切り形式（推奨）
python mcp_test_cli.py get_powerpoint_attributes \
  --file_path "presentation.pptx" \
  --attributes title,subtitle,object_counts

# JSON形式（PowerShell）
python mcp_test_cli.py get_powerpoint_attributes \
  --file_path "presentation.pptx" \
  --attributes '["title", "subtitle", "object_counts"]'

# JSON形式（Windows CMD）
python mcp_test_cli.py get_powerpoint_attributes \
  --file_path "presentation.pptx" \
  --attributes "[""title"", ""subtitle"", ""object_counts""]"
```

### スライド情報取得

```bash
python mcp_test_cli.py get_slide_info \
  --file_path "presentation.pptx" \
  --slide_number 1
```

### 太字テキスト抽出

```bash
python mcp_test_cli.py extract_bold_text --file_path "presentation.pptx"
```

### 特定フォーマットのテキスト抽出

```bash
python mcp_test_cli.py extract_text_formatting \
  --file_path "presentation.pptx" \
  --formatting_type "italic"
```

### プレゼンテーション概要取得

```bash
python mcp_test_cli.py get_presentation_overview \
  --file_path "presentation.pptx" \
  --analysis_depth "detailed"
```

### スライドクエリ（条件検索）

```bash
# 簡略JSON形式（Windows CMD）
python mcp_test_cli.py query_slides \
  --file_path "presentation.pptx" \
  --search_criteria "{contains: bold, has_tables: true}"

# 標準JSON形式（PowerShell）
python mcp_test_cli.py query_slides \
  --file_path "presentation.pptx" \
  --search_criteria '{"contains": "bold", "has_tables": true}'

# Windows CMD標準JSON形式
python mcp_test_cli.py query_slides \
  --file_path "presentation.pptx" \
  --search_criteria "{""contains"": ""bold"", ""has_tables"": true}"

# ネストしたオブジェクト（Windows CMD簡略形式）
python mcp_test_cli.py query_slides \
  --file_path "presentation.pptx" \
  --search_criteria "{content: {contains_text: Slide}, layout: {type: content}}"

# 配列を含む複雑なクエリ
python mcp_test_cli.py query_slides \
  --file_path "presentation.pptx" \
  --search_criteria "{slide_numbers: [1, 2, 3], content: {min_elements: 2}}"
```

## 🔧 パラメータの指定方法

### 文字列パラメータ
```bash
--file_path "path/to/file.pptx"
```

### 数値パラメータ
```bash
--slide_number 1
```

### ブール値パラメータ
```bash
--include_sample_content true
--clear_cache false
```

### 配列パラメータ

#### 1. カンマ区切り形式（最も簡単）
```bash
--attributes title,subtitle,text_elements
--slide_numbers 1,2,3
```

#### 2. JSON形式 - PowerShell
```bash
--attributes '["title", "subtitle", "text_elements"]'
--slide_numbers '[1, 2, 3]'
```

#### 3. JSON形式 - Windows CMD
```bash
# ダブルクォートをダブルクォートでエスケープ
--attributes "[""title"", ""subtitle"", ""text_elements""]"

# または簡略形式（クォートは自動追加）
--attributes "[title, subtitle, text_elements]"
```

### オブジェクトパラメータ

#### 1. JSON形式 - PowerShell
```bash
--search_criteria '{"has_tables": true, "min_text_elements": 2}'
```

#### 2. JSON形式 - Windows CMD
```bash
# ダブルクォートをダブルクォートでエスケープ
--search_criteria "{""has_tables"": true, ""min_text_elements"": 2}"

# または簡略形式（クォートは自動追加）
--search_criteria "{has_tables: true, min_text_elements: 2}"

# ネストしたオブジェクト
--search_criteria "{""content"": {""contains_text"": ""Slide""}}"
--search_criteria "{content: {contains_text: Slide}}"

# 配列を含むオブジェクト
--search_criteria "{""slide_numbers"": [1, 2, 3]}"
--search_criteria "{slide_numbers: [1, 2, 3]}"
```

## 🧪 テスト例の実行

包括的なテスト例を実行：

```bash
python examples/test_examples.py
```

このスクリプトは以下を実行します：
- 全ツールの一覧表示
- 特定ツールのヘルプ表示
- 各種ツールの実行例
- 簡易ラッパーの使用例

## 🛠️ トラブルシューティング

### サーバーが起動しない場合

1. `main.py`が存在することを確認
2. 必要な依存関係がインストールされていることを確認
3. PowerPoint Analyzer MCPが正しく設定されていることを確認

### 通信エラーが発生する場合

1. サーバーのログを確認（`powerpoint_mcp_server.log`）
2. JSON-RPCメッセージの形式を確認
3. サーバープロセスが正常に動作していることを確認

### ツールが見つからない場合

1. サーバーが正しく初期化されていることを確認
2. `tools/list`リクエストが正常に動作することを確認
3. ツール名のスペルを確認

## 📝 カスタマイズ

### 異なるサーバーコマンドを使用

```python
cli = MCPTestCLI(server_command=["python", "path/to/your/server.py"])
```

### タイムアウト設定

サーバーの応答が遅い場合は、`asyncio.wait_for()`を使用してタイムアウトを設定できます。

### ログ出力

デバッグ用にログ出力を追加する場合は、`logging`モジュールを使用してください。

## 🎯 開発者向け情報

### MCPプロトコル対応

- JSON-RPC 2.0準拠
- MCP Protocol Version 2024-11-05対応
- FastMCP 2.0サーバー対応

### 拡張可能性

新しいテストシナリオを追加する場合は、`test_tools.py`に新しいコマンドを追加するか、`mcp_test_cli.py`を直接拡張してください。

## 📄 ライセンス

このツールはPowerPoint Analyzer MCPプロジェクトの一部として、同じライセンス（Apache License 2.0）の下で提供されます。