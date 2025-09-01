# Flexible PowerPoint MCP Tools Design

## 🎯 設計思想

### 柔軟性と対話性を重視した設計
- **段階的な情報収集**: AIエージェントが必要に応じて詳細な条件を指定
- **結果に基づく次のアクション**: 検索結果を見てから抽出条件を調整
- **複雑な条件の分解**: 複雑な要求を複数のシンプルなクエリに分解
- **中間結果の活用**: 前の結果を次のクエリの入力として使用

## 🔧 Core Flexible Tools

### 1. **query_slides**
**目的**: 柔軟な条件でスライドを検索・フィルタリング

```json
{
  "name": "query_slides",
  "description": "Search and filter slides using flexible criteria with support for complex conditions",
  "inputSchema": {
    "type": "object",
    "properties": {
      "file_path": {"type": "string", "description": "Path to PowerPoint file"},
      "filters": {
        "type": "object",
        "properties": {
          "title": {
            "type": "object",
            "properties": {
              "contains": {"type": "string", "description": "Title contains text"},
              "starts_with": {"type": "string", "description": "Title starts with text"},
              "ends_with": {"type": "string", "description": "Title ends with text"},
              "regex": {"type": "string", "description": "Title matches regex pattern"},
              "one_of": {"type": "array", "items": {"type": "string"}, "description": "Title matches any of these patterns"}
            }
          },
          "content": {
            "type": "object",
            "properties": {
              "contains_text": {"type": "string", "description": "Slide contains specific text"},
              "has_tables": {"type": "boolean", "description": "Slide has tables"},
              "has_charts": {"type": "boolean", "description": "Slide has charts"},
              "has_images": {"type": "boolean", "description": "Slide has images"},
              "object_count": {
                "type": "object",
                "properties": {
                  "min": {"type": "integer", "description": "Minimum object count"},
                  "max": {"type": "integer", "description": "Maximum object count"}
                }
              }
            }
          },
          "layout": {
            "type": "object",
            "properties": {
              "type": {"type": "string", "description": "Specific layout type"},
              "name": {"type": "string", "description": "Layout name pattern"}
            }
          },
          "slide_numbers": {"type": "array", "items": {"type": "integer"}, "description": "Specific slide numbers"},
          "section": {"type": "string", "description": "Section name"}
        }
      },
      "return_fields": {
        "type": "array",
        "items": {"type": "string", "enum": ["slide_number", "title", "subtitle", "layout", "object_counts", "preview_text", "table_info", "full_content"]},
        "default": ["slide_number", "title", "object_counts"],
        "description": "Fields to return for each matching slide"
      },
      "limit": {"type": "integer", "default": 50, "description": "Maximum number of slides to return"}
    },
    "required": ["file_path"]
  }
}
```

### 2. **extract_table_data**
**目的**: 指定されたスライドからテーブルデータを柔軟に抽出

```json
{
  "name": "extract_table_data",
  "description": "Extract table data from specified slides with flexible column selection and formatting detection",
  "inputSchema": {
    "type": "object",
    "properties": {
      "file_path": {"type": "string", "description": "Path to PowerPoint file"},
      "slide_numbers": {"type": "array", "items": {"type": "integer"}, "description": "Slide numbers to extract tables from"},
      "table_selection": {
        "type": "object",
        "properties": {
          "table_index": {"type": "integer", "description": "Specific table index (0-based) if multiple tables on slide"},
          "table_criteria": {
            "type": "object",
            "properties": {
              "min_rows": {"type": "integer", "description": "Minimum number of rows"},
              "min_columns": {"type": "integer", "description": "Minimum number of columns"},
              "header_contains": {"type": "array", "items": {"type": "string"}, "description": "Headers that table must contain"}
            }
          }
        }
      },
      "column_selection": {
        "type": "object",
        "properties": {
          "columns": {"type": "array", "items": {"type": "string"}, "description": "Specific column names to extract"},
          "column_patterns": {"type": "array", "items": {"type": "string"}, "description": "Regex patterns for column names"},
          "exclude_columns": {"type": "array", "items": {"type": "string"}, "description": "Column names to exclude"}
        }
      },
      "formatting_detection": {
        "type": "object",
        "properties": {
          "detect_bold": {"type": "boolean", "default": true, "description": "Detect bold text in cells"},
          "detect_italic": {"type": "boolean", "default": true, "description": "Detect italic text in cells"},
          "detect_highlight": {"type": "boolean", "default": true, "description": "Detect highlighted text in cells"},
          "detect_colors": {"type": "boolean", "default": false, "description": "Detect text colors"},
          "detect_hyperlinks": {"type": "boolean", "default": true, "description": "Detect hyperlinks in cells"}
        }
      },
      "output_format": {
        "type": "string",
        "enum": ["structured", "flat", "grouped_by_slide"],
        "default": "structured",
        "description": "How to structure the output data"
      },
      "include_metadata": {"type": "boolean", "default": true, "description": "Include table position and size metadata"}
    },
    "required": ["file_path", "slide_numbers"]
  }
}
```

### 3. **analyze_text_formatting**
**目的**: テキストのフォーマット情報を詳細に分析

```json
{
  "name": "analyze_text_formatting",
  "description": "Analyze text formatting (bold, italic, highlight, colors) in specified content",
  "inputSchema": {
    "type": "object",
    "properties": {
      "file_path": {"type": "string", "description": "Path to PowerPoint file"},
      "target": {
        "type": "object",
        "properties": {
          "slide_numbers": {"type": "array", "items": {"type": "integer"}, "description": "Slides to analyze"},
          "content_type": {"type": "string", "enum": ["all", "tables", "text_boxes", "titles", "bullets"], "default": "all"},
          "specific_text": {"type": "string", "description": "Analyze formatting of specific text content"}
        }
      },
      "formatting_types": {
        "type": "array",
        "items": {"type": "string", "enum": ["bold", "italic", "underline", "highlight", "strikethrough", "color", "font_size", "hyperlink"]},
        "default": ["bold", "italic", "highlight"],
        "description": "Types of formatting to detect"
      },
      "filter_criteria": {
        "type": "object",
        "properties": {
          "has_formatting": {"type": "array", "items": {"type": "string"}, "description": "Only return text that has these formatting types"},
          "text_contains": {"type": "string", "description": "Only analyze text containing this string"},
          "min_length": {"type": "integer", "description": "Minimum text length to analyze"}
        }
      },
      "group_by": {"type": "string", "enum": ["slide", "formatting_type", "content_type"], "default": "slide", "description": "How to group the results"}
    },
    "required": ["file_path", "target"]
  }
}
```

### 4. **filter_and_aggregate**
**目的**: 抽出されたデータをフィルタリング・集約

```json
{
  "name": "filter_and_aggregate",
  "description": "Filter and aggregate previously extracted data based on complex conditions",
  "inputSchema": {
    "type": "object",
    "properties": {
      "data_source": {
        "type": "object",
        "properties": {
          "type": {"type": "string", "enum": ["table_data", "text_analysis", "slide_query"], "description": "Type of source data"},
          "data": {"type": "object", "description": "The data to filter and aggregate"}
        }
      },
      "filters": {
        "type": "array",
        "items": {
          "type": "object",
          "properties": {
            "field": {"type": "string", "description": "Field name to filter on"},
            "condition": {"type": "string", "enum": ["equals", "contains", "starts_with", "ends_with", "regex", "not_empty", "has_formatting"], "description": "Filter condition"},
            "value": {"type": "string", "description": "Value to compare against"},
            "formatting_types": {"type": "array", "items": {"type": "string"}, "description": "Required formatting types for has_formatting condition"}
          }
        }
      },
      "aggregation": {
        "type": "object",
        "properties": {
          "group_by": {"type": "array", "items": {"type": "string"}, "description": "Fields to group by"},
          "operations": {
            "type": "array",
            "items": {
              "type": "object",
              "properties": {
                "field": {"type": "string", "description": "Field to aggregate"},
                "operation": {"type": "string", "enum": ["count", "list", "unique", "concat"], "description": "Aggregation operation"}
              }
            }
          }
        }
      },
      "sort": {
        "type": "object",
        "properties": {
          "field": {"type": "string", "description": "Field to sort by"},
          "order": {"type": "string", "enum": ["asc", "desc"], "default": "asc"}
        }
      }
    },
    "required": ["data_source"]
  }
}
```

### 5. **get_presentation_overview**
**目的**: プレゼンテーション全体の概要を取得（探索的分析の開始点）

```json
{
  "name": "get_presentation_overview",
  "description": "Get comprehensive overview of presentation structure and content for exploration",
  "inputSchema": {
    "type": "object",
    "properties": {
      "file_path": {"type": "string", "description": "Path to PowerPoint file"},
      "analysis_depth": {"type": "string", "enum": ["basic", "detailed", "comprehensive"], "default": "basic"},
      "include_samples": {"type": "boolean", "default": true, "description": "Include sample content from each slide type"},
      "detect_patterns": {"type": "boolean", "default": true, "description": "Detect common patterns in titles and content"}
    },
    "required": ["file_path"]
  }
}
```

## 🔄 AIエージェント自動ラリー対応設計

### Claude Sonnet 3.5相当での自動実行フロー

#### 🧠 AIエージェントの判断ロジック
1. **コンテキスト保持**: 前回の結果を次のクエリで自動参照
2. **条件の自動調整**: 結果が多すぎる/少なすぎる場合の自動調整
3. **エラー時の自動フォールバック**: 失敗時の代替アプローチ
4. **結果の自動検証**: 期待する形式かどうかの自動チェック

#### 自動実行例: "xxx サブプロジェクトA/B" の表から特定項目を抽出

**Step 1: 探索的分析（自動実行）**
```json
{
  "tool": "get_presentation_overview",
  "params": {
    "file_path": "project_report.pptx",
    "analysis_depth": "detailed",
    "detect_patterns": true
  }
}
```
→ **AIが自動判断**: スライド数、タイトルパターン、テーブル有無を確認

**Step 2: 対象スライドの特定（結果に基づく自動調整）**
```json
{
  "tool": "query_slides",
  "params": {
    "file_path": "project_report.pptx",
    "filters": {
      "title": {
        "one_of": [".*サブプロジェクトA.*", ".*サブプロジェクトB.*"]
      },
      "content": {
        "has_tables": true
      }
    },
    "return_fields": ["slide_number", "title", "table_info", "preview_text"]
  }
}
```
→ **AIが自動判断**: 
- 結果が0件 → タイトルパターンを緩和して再実行
- 結果が多すぎる → より具体的な条件を追加
- 適切な件数 → 次のステップへ

**Step 3: テーブル構造の事前確認（自動最適化）**
```json
{
  "tool": "extract_table_data",
  "params": {
    "file_path": "project_report.pptx",
    "slide_numbers": [3],  // まず1つのスライドで構造確認
    "column_selection": {},  // 全列を取得して構造を把握
    "formatting_detection": {
      "detect_bold": true,
      "detect_highlight": true
    },
    "output_format": "structured",
    "include_metadata": true
  }
}
```
→ **AIが自動判断**: 列名の正確なマッチング、データ形式の確認

**Step 4: 全対象スライドからの抽出（最適化された条件で）**
```json
{
  "tool": "extract_table_data",
  "params": {
    "file_path": "project_report.pptx",
    "slide_numbers": [3, 7, 12],  // Step 2の結果を自動使用
    "column_selection": {
      "columns": ["タスク名", "進捗", "課題", "重要度"]  // Step 3で確認した正確な列名
    },
    "formatting_detection": {
      "detect_bold": true,
      "detect_highlight": true
    },
    "output_format": "structured"
  }
}
```

**Step 5: 条件フィルタリング（自動実行）**
```json
{
  "tool": "filter_and_aggregate",
  "params": {
    "data_source": {
      "type": "table_data",
      "data": "{{step4_result}}"  // 前の結果を自動参照
    },
    "filters": [
      {
        "field": "課題",
        "condition": "has_formatting",
        "formatting_types": ["bold", "highlight"]
      },
      {
        "field": "課題",
        "condition": "not_empty"  // 空の課題は除外
      }
    ],
    "aggregation": {
      "group_by": ["slide_number", "タスク名"],
      "operations": [
        {"field": "課題", "operation": "list"},
        {"field": "重要度", "operation": "list"}
      ]
    },
    "sort": {
      "field": "重要度",
      "order": "desc"
    }
  }
}
```

## 🎯 設計の利点

### 1. **段階的な探索**
- 概要 → 検索 → 抽出 → フィルタリングの段階的アプローチ
- 各段階で結果を確認してから次の条件を決定

### 2. **柔軟な条件指定**
- 複雑な条件を組み合わせ可能
- 正規表現やパターンマッチングをサポート
- フォーマット情報の詳細な検出

### 3. **結果の再利用**
- 前の結果を次のクエリの入力として使用
- 中間結果の保存と参照が可能

### 4. **エラー処理とフォールバック**
- 条件が複雑すぎる場合は段階的に分解
- 部分的な結果でも有用な情報を提供

## 🔧 実装上の考慮事項

### パフォーマンス最適化
- **遅延評価**: 必要な部分のみを処理
- **結果キャッシュ**: 同じファイルの繰り返し処理を高速化
- **段階的処理**: 大きなファイルでも応答性を維持

### エラーハンドリング
- **部分的成功**: 一部のスライドでエラーが発生しても継続
- **詳細なエラー情報**: どの部分で問題が発生したかを明確に報告
- **代替案の提示**: エラー時に代替的なアプローチを提案

この設計により、AIエージェントは複雑な要求を段階的に処理し、ユーザーの意図を正確に理解して必要な情報を抽出できるようになります。
## 🤖
 AIエージェント自動ラリー対応の追加機能

### 6. **get_query_suggestions**
**目的**: AIエージェントが次に実行すべきクエリを自動提案

```json
{
  "name": "get_query_suggestions",
  "description": "Get intelligent suggestions for next queries based on current context and results",
  "inputSchema": {
    "type": "object",
    "properties": {
      "file_path": {"type": "string", "description": "Path to PowerPoint file"},
      "current_context": {
        "type": "object",
        "properties": {
          "user_intent": {"type": "string", "description": "Original user request"},
          "previous_results": {"type": "array", "items": {"type": "object"}, "description": "Results from previous queries"},
          "current_step": {"type": "string", "description": "Current step in the workflow"}
        }
      },
      "result_analysis": {
        "type": "object",
        "properties": {
          "result_count": {"type": "integer", "description": "Number of results from last query"},
          "data_quality": {"type": "string", "enum": ["complete", "partial", "insufficient"], "description": "Quality of current results"},
          "missing_elements": {"type": "array", "items": {"type": "string"}, "description": "Elements that seem to be missing"}
        }
      }
    },
    "required": ["file_path", "current_context"]
  }
}
```

### 7. **validate_extraction_results**
**目的**: 抽出結果の妥当性を自動検証

```json
{
  "name": "validate_extraction_results",
  "description": "Validate extraction results against expected patterns and suggest corrections",
  "inputSchema": {
    "type": "object",
    "properties": {
      "results": {"type": "object", "description": "Results to validate"},
      "validation_criteria": {
        "type": "object",
        "properties": {
          "expected_fields": {"type": "array", "items": {"type": "string"}, "description": "Fields that should be present"},
          "expected_count_range": {"type": "object", "properties": {"min": {"type": "integer"}, "max": {"type": "integer"}}},
          "data_type_validation": {"type": "object", "description": "Expected data types for each field"},
          "business_rules": {"type": "array", "items": {"type": "string"}, "description": "Business logic validation rules"}
        }
      },
      "auto_correction": {"type": "boolean", "default": true, "description": "Attempt automatic correction of issues"}
    },
    "required": ["results"]
  }
}
```

## 🔄 自動エラー処理とフォールバック

### エラーシナリオと自動対応

#### 1. **検索結果が0件の場合**
```python
# AIエージェントの自動判断ロジック
if search_results.count == 0:
    # パターン1: タイトル条件を緩和
    fallback_query = modify_title_pattern(original_pattern, "broader")
    
    # パターン2: 全スライドから類似タイトルを検索
    similar_titles = find_similar_titles(file_path, original_pattern)
    
    # パターン3: コンテンツベースの検索に切り替え
    content_search = search_by_content(keywords_from_title)
```

#### 2. **列名が見つからない場合**
```python
# AIエージェントの自動対応
if requested_columns not in table_headers:
    # 類似列名の自動検出
    similar_columns = find_similar_column_names(requested_columns, actual_headers)
    
    # ユーザーに確認せずに最も類似度の高い列名を使用
    auto_mapped_columns = auto_map_columns(requested_columns, similar_columns)
```

#### 3. **フォーマット検出が期待通りでない場合**
```python
# AIエージェントの自動調整
if formatting_results.count < expected_minimum:
    # より広範囲のフォーマット検出
    expanded_formatting = detect_all_formatting_types()
    
    # 代替的な強調表現の検索
    alternative_emphasis = find_alternative_emphasis_patterns()
```

## 🎯 Claude Sonnet 3.5での実行保証

### 1. **明確な実行フロー**
- 各ツールの出力形式を標準化
- 次のステップの判断基準を明確化
- エラー時の代替パスを事前定義

### 2. **コンテキスト保持機能**
- 前回の結果を自動的に次のクエリで参照
- 実行履歴の自動管理
- 中間結果の一時保存

### 3. **自動最適化**
- クエリパフォーマンスの自動調整
- 結果の品質に基づく条件の自動調整
- リソース使用量の自動制御

### 4. **結果の自動検証**
- 期待する結果形式との自動比較
- データの整合性チェック
- ビジネスルールの自動適用

この設計により、Claude Sonnet 3.5相当のAIエージェントは、ユーザーの介入なしに複雑なPowerPoint分析タスクを自動的に実行できます。エラーが発生した場合も、自動的に代替アプローチを試行し、最適な結果を得るまで継続的に調整を行います。## 📚
 AIエージェント向けガイダンス設計

### MCP Resources: 利用可能属性の完全ドキュメント

#### Resource 1: **powerpoint_extraction_capabilities**
```json
{
  "uri": "powerpoint://capabilities",
  "name": "PowerPoint Extraction Capabilities",
  "description": "Complete reference of all extractable attributes and their usage patterns",
  "mimeType": "application/json",
  "content": {
    "slide_attributes": {
      "basic": ["slide_number", "title", "subtitle", "layout_name", "layout_type"],
      "content": ["text_elements", "tables", "images", "shapes", "charts"],
      "metadata": ["object_counts", "slide_size", "position_info"],
      "formatting": ["bold", "italic", "underline", "highlight", "strikethrough", "color", "font_size", "hyperlink"]
    },
    "search_patterns": {
      "title_matching": {
        "exact": "title.equals('exact text')",
        "contains": "title.contains('partial text')",
        "starts_with": "title.starts_with('prefix')",
        "regex": "title.regex('pattern')",
        "multiple": "title.one_of(['pattern1', 'pattern2'])"
      },
      "content_filtering": {
        "has_tables": "content.has_tables = true",
        "has_charts": "content.has_charts = true",
        "text_contains": "content.contains_text = 'search term'",
        "object_count": "content.object_count.min/max = number"
      }
    },
    "table_extraction": {
      "column_selection": {
        "specific": "columns: ['列名1', '列名2']",
        "pattern": "column_patterns: ['.*進捗.*', '.*課題.*']",
        "exclude": "exclude_columns: ['不要列']"
      },
      "formatting_detection": {
        "text_formatting": ["bold", "italic", "highlight", "color"],
        "cell_properties": ["hyperlinks", "merged_cells"],
        "conditional": "detect only cells with specific formatting"
      }
    },
    "common_workflows": {
      "progress_tracking": {
        "steps": ["overview", "search_slides", "extract_tables", "filter_formatting"],
        "typical_columns": ["タスク名", "進捗", "課題", "重要度", "担当者", "期限"],
        "formatting_indicators": ["highlight for urgent", "bold for completed"]
      },
      "data_comparison": {
        "steps": ["search_similar_slides", "extract_consistent_format", "aggregate_results"],
        "grouping_strategies": ["by_slide", "by_project", "by_date"]
      }
    }
  }
}
```

#### Resource 2: **workflow_execution_guide**
```json
{
  "uri": "powerpoint://workflow_guide",
  "name": "Automated Workflow Execution Guide",
  "description": "Step-by-step guide for AI agents to execute complex PowerPoint analysis workflows",
  "mimeType": "application/json",
  "content": {
    "execution_principles": {
      "progressive_refinement": "Start broad, then narrow down based on results",
      "error_recovery": "Always have fallback strategies for each step",
      "context_preservation": "Maintain context between tool calls",
      "result_validation": "Verify results meet user expectations"
    },
    "decision_trees": {
      "search_results_empty": {
        "condition": "search_results.count == 0",
        "actions": [
          "broaden_title_pattern",
          "search_by_content_keywords",
          "check_all_slides_for_similar_patterns"
        ]
      },
      "too_many_results": {
        "condition": "search_results.count > 20",
        "actions": [
          "add_content_filters",
          "narrow_title_pattern",
          "add_slide_range_filter"
        ]
      },
      "column_not_found": {
        "condition": "requested_column not in table_headers",
        "actions": [
          "find_similar_column_names",
          "extract_all_columns_first",
          "use_column_pattern_matching"
        ]
      }
    }
  }
}
```

### MCP Prompts: 自動ラリー対応テンプレート

#### Prompt 1: **complex_data_extraction**
```json
{
  "name": "complex_data_extraction",
  "description": "Template for extracting complex data from PowerPoint presentations with automatic workflow execution",
  "arguments": [
    {
      "name": "file_path",
      "description": "Path to the PowerPoint file",
      "required": true
    },
    {
      "name": "extraction_goal",
      "description": "High-level description of what data to extract",
      "required": true
    },
    {
      "name": "specific_conditions",
      "description": "Specific conditions or filters to apply",
      "required": false
    }
  ],
  "template": "I need to extract data from a PowerPoint presentation. Here's my systematic approach:\n\n1. **EXPLORATION PHASE**\n   - First, I'll get an overview of the presentation structure\n   - Identify slide patterns and content types\n   - Understand the data organization\n\n2. **SEARCH PHASE**\n   - Search for slides matching the criteria: {{extraction_goal}}\n   - If no results: broaden search criteria automatically\n   - If too many results: add more specific filters\n   - Target: 3-15 relevant slides for optimal processing\n\n3. **EXTRACTION PHASE**\n   - Test extraction on one slide first to understand data structure\n   - Apply optimized extraction to all target slides\n   - Handle missing columns or unexpected formats automatically\n\n4. **FILTERING PHASE**\n   - Apply specific conditions: {{specific_conditions}}\n   - Focus on formatting-based filters (bold, highlight, etc.)\n   - Aggregate and organize results logically\n\n5. **VALIDATION PHASE**\n   - Verify results meet the original goal\n   - Check for completeness and accuracy\n   - Provide summary and insights\n\nLet me start with step 1 - getting the presentation overview for: {{file_path}}"
}
```

#### Prompt 2: **progressive_table_analysis**
```json
{
  "name": "progressive_table_analysis",
  "description": "Template for progressive table analysis with automatic error recovery",
  "arguments": [
    {
      "name": "file_path",
      "description": "Path to the PowerPoint file",
      "required": true
    },
    {
      "name": "slide_criteria",
      "description": "Criteria for selecting slides",
      "required": true
    },
    {
      "name": "target_columns",
      "description": "Columns to extract from tables",
      "required": true
    },
    {
      "name": "formatting_focus",
      "description": "Specific formatting to focus on",
      "required": false
    }
  ],
  "template": "I'll analyze tables in PowerPoint slides using a progressive approach:\n\n**STEP 1: SLIDE IDENTIFICATION**\nSearching for slides with criteria: {{slide_criteria}}\n- If 0 results → I'll automatically broaden the search\n- If >20 results → I'll add more specific filters\n- Target: Find 3-15 relevant slides\n\n**STEP 2: TABLE STRUCTURE ANALYSIS**\nAnalyzing one representative slide first:\n- Identify actual column names in tables\n- Map requested columns {{target_columns}} to actual headers\n- Understand table format and data types\n\n**STEP 3: BULK EXTRACTION**\nExtracting from all identified slides:\n- Use optimized column mapping from step 2\n- Apply formatting detection: {{formatting_focus}}\n- Handle variations in table structure automatically\n\n**STEP 4: INTELLIGENT FILTERING**\nApplying smart filters:\n- Focus on cells with specific formatting\n- Remove empty or irrelevant entries\n- Group results logically\n\n**ERROR RECOVERY STRATEGIES:**\n- Column name mismatch → Find similar column names automatically\n- No tables found → Search by content keywords\n- Formatting not detected → Expand formatting detection scope\n\nStarting analysis of: {{file_path}}"
}
```

#### Prompt 3: **adaptive_search_strategy**
```json
{
  "name": "adaptive_search_strategy",
  "description": "Template for adaptive search with automatic strategy adjustment",
  "arguments": [
    {
      "name": "file_path",
      "description": "Path to the PowerPoint file",
      "required": true
    },
    {
      "name": "search_intent",
      "description": "What the user is looking for",
      "required": true
    }
  ],
  "template": "I'll use an adaptive search strategy to find: {{search_intent}}\n\n**ADAPTIVE SEARCH ALGORITHM:**\n\n1. **Initial Broad Search**\n   - Start with generous criteria to understand content landscape\n   - Identify patterns and common structures\n\n2. **Progressive Refinement**\n   - If results < 3: Broaden criteria (remove filters, use partial matches)\n   - If results > 15: Add specific filters (content type, layout, keywords)\n   - If results 3-15: Proceed with extraction\n\n3. **Automatic Fallback Strategies**\n   - Title search fails → Content-based search\n   - Exact match fails → Fuzzy matching\n   - Pattern match fails → Keyword search\n\n4. **Context-Aware Adjustments**\n   - Learn from successful matches\n   - Adapt patterns based on presentation style\n   - Optimize for presentation-specific conventions\n\n**EXECUTION MONITORING:**\n- Track success rate of each strategy\n- Automatically switch to more effective approaches\n- Maintain context across multiple tool calls\n\nBeginning adaptive search for '{{search_intent}}' in: {{file_path}}"
}
```

## 🎯 実装における AIエージェント支援機能

### 1. **自動ワークフロー検出**
```python
# MCPサーバー側で実装
def detect_workflow_pattern(user_request: str) -> str:
    \"\"\"ユーザーリクエストから適切なワークフローパターンを自動検出\"\"\"
    patterns = {
        "progress_extraction": ["進捗", "状況", "課題", "タスク"],
        "data_comparison": ["比較", "変化", "推移", "トレンド"],
        "formatted_content": ["ハイライト", "ボールド", "強調", "重要"]
    }
    # パターンマッチングロジック
    return detected_pattern
```

### 2. **コンテキスト保持機能**
```python
# 実行コンテキストの自動管理
class ExecutionContext:
    def __init__(self):
        self.previous_results = []
        self.current_strategy = None
        self.failed_attempts = []
        self.learned_patterns = {}
    
    def suggest_next_action(self, current_result):
        \"\"\"現在の結果に基づいて次のアクションを提案\"\"\"
        if current_result.count == 0:
            return self.broaden_search_strategy()
        elif current_result.count > 20:
            return self.narrow_search_strategy()
        else:
            return self.proceed_to_extraction()
```

### 3. **エラー予測と事前対策**
```python
# よくあるエラーパターンの事前検出
def predict_potential_issues(query_params):
    \"\"\"クエリパラメータから潜在的な問題を予測\"\"\"
    warnings = []
    if "exact_column_names" in query_params:
        warnings.append("Consider using column_patterns for flexibility")
    if "strict_title_match" in query_params:
        warnings.append("Prepare fallback with partial matching")
    return warnings
```

この設計により、AIエージェントは：
1. **完全なドキュメント**から利用可能な全機能を把握
2. **プロンプトテンプレート**で具体的な実行方法を理解
3. **自動ワークフロー**で効率的な処理を実現
4. **エラー予測**で問題を事前に回避

Claude Sonnet 3.5相当であれば、これらのガイダンスを活用して確実に自動ラリーを実行できます！