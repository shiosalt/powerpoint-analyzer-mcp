# Enhanced PowerPoint MCP Tools Specification

## 🎯 ビジネスユースケース対応

### 想定シナリオ
- 毎週定型で作成するPowerPointファイルから特定情報を取得
- タイトルが特定パターンで始まる複数のスライドから表データを一括抽出
- 複数のPowerPointファイルから同じ形式の情報を一括取得
- AIエージェントによる自動整理・記録

## 🔧 Enhanced MCP Tools

### 1. **search_slides_by_criteria** 
**目的**: 条件に基づくスライド検索

```json
{
  "name": "search_slides_by_criteria",
  "description": "Search slides based on various criteria (title pattern, layout, content, etc.)",
  "inputSchema": {
    "type": "object",
    "properties": {
      "file_path": {"type": "string", "description": "Path to PowerPoint file"},
      "criteria": {
        "type": "object",
        "properties": {
          "title_pattern": {"type": "string", "description": "Regex pattern for slide titles"},
          "title_starts_with": {"type": "string", "description": "Title prefix to match"},
          "title_contains": {"type": "string", "description": "Text that title must contain"},
          "layout_type": {"type": "string", "description": "Specific layout type"},
          "has_tables": {"type": "boolean", "description": "Slides that contain tables"},
          "has_charts": {"type": "boolean", "description": "Slides that contain charts"},
          "slide_numbers": {"type": "array", "items": {"type": "integer"}, "description": "Specific slide numbers"},
          "section_name": {"type": "string", "description": "Section name to search within"},
          "content_contains": {"type": "string", "description": "Text content to search for"}
        }
      },
      "return_content": {"type": "boolean", "default": true, "description": "Whether to return full content or just metadata"}
    },
    "required": ["file_path", "criteria"]
  }
}
```

### 2. **extract_tables_from_slides**
**目的**: 複数スライドからテーブルデータを一括抽出

```json
{
  "name": "extract_tables_from_slides",
  "description": "Extract all tables from specified slides with structured data format",
  "inputSchema": {
    "type": "object",
    "properties": {
      "file_path": {"type": "string", "description": "Path to PowerPoint file"},
      "slide_criteria": {
        "type": "object",
        "description": "Criteria to select slides (same as search_slides_by_criteria)"
      },
      "table_format": {
        "type": "string",
        "enum": ["json", "csv", "structured"],
        "default": "structured",
        "description": "Output format for table data"
      },
      "include_headers": {"type": "boolean", "default": true, "description": "Include table headers"},
      "merge_tables": {"type": "boolean", "default": false, "description": "Merge all tables into one dataset"},
      "filter_columns": {"type": "array", "items": {"type": "string"}, "description": "Specific columns to extract"}
    },
    "required": ["file_path"]
  }
}
```

### 3. **batch_extract_from_files**
**目的**: 複数ファイルから一括データ抽出

```json
{
  "name": "batch_extract_from_files",
  "description": "Extract data from multiple PowerPoint files using the same criteria",
  "inputSchema": {
    "type": "object",
    "properties": {
      "file_paths": {"type": "array", "items": {"type": "string"}, "description": "List of PowerPoint file paths"},
      "extraction_template": {
        "type": "object",
        "properties": {
          "slide_criteria": {"type": "object", "description": "Criteria to select slides"},
          "data_points": {
            "type": "array",
            "items": {
              "type": "object",
              "properties": {
                "name": {"type": "string", "description": "Data point name"},
                "type": {"type": "string", "enum": ["text", "table", "number", "date", "list"]},
                "extraction_rule": {"type": "string", "description": "How to extract this data point"},
                "required": {"type": "boolean", "default": false}
              }
            }
          }
        }
      },
      "output_format": {"type": "string", "enum": ["json", "csv", "summary"], "default": "json"},
      "consolidate_results": {"type": "boolean", "default": true, "description": "Combine results from all files"}
    },
    "required": ["file_paths", "extraction_template"]
  }
}
```

### 4. **extract_progress_data**
**目的**: 進捗データの専用抽出（よくあるユースケース）

```json
{
  "name": "extract_progress_data",
  "description": "Extract progress/status data from slides (specialized for common business use case)",
  "inputSchema": {
    "type": "object",
    "properties": {
      "file_path": {"type": "string", "description": "Path to PowerPoint file"},
      "progress_indicators": {
        "type": "object",
        "properties": {
          "title_patterns": {"type": "array", "items": {"type": "string"}, "description": "Patterns to identify progress slides"},
          "status_keywords": {"type": "array", "items": {"type": "string"}, "description": "Keywords that indicate status"},
          "date_formats": {"type": "array", "items": {"type": "string"}, "description": "Expected date formats"},
          "percentage_extraction": {"type": "boolean", "default": true, "description": "Extract percentage values"},
          "milestone_extraction": {"type": "boolean", "default": true, "description": "Extract milestone information"}
        }
      },
      "output_structure": {
        "type": "string",
        "enum": ["timeline", "summary", "detailed", "dashboard"],
        "default": "summary",
        "description": "How to structure the progress data"
      }
    },
    "required": ["file_path"]
  }
}
```

### 5. **analyze_slide_patterns**
**目的**: スライドパターンの分析と分類

```json
{
  "name": "analyze_slide_patterns",
  "description": "Analyze and categorize slides based on layout patterns and content structure",
  "inputSchema": {
    "type": "object",
    "properties": {
      "file_path": {"type": "string", "description": "Path to PowerPoint file"},
      "analysis_type": {
        "type": "string",
        "enum": ["layout_analysis", "content_classification", "template_detection", "anomaly_detection"],
        "default": "content_classification",
        "description": "Type of pattern analysis to perform"
      },
      "grouping_criteria": {
        "type": "array",
        "items": {"type": "string", "enum": ["layout", "content_type", "object_count", "text_density", "visual_elements"]},
        "description": "Criteria for grouping similar slides"
      },
      "include_recommendations": {"type": "boolean", "default": true, "description": "Include improvement recommendations"}
    },
    "required": ["file_path"]
  }
}
```

### 6. **extract_structured_data**
**目的**: 構造化データの柔軟な抽出

```json
{
  "name": "extract_structured_data",
  "description": "Extract data using flexible field mapping and transformation rules",
  "inputSchema": {
    "type": "object",
    "properties": {
      "file_path": {"type": "string", "description": "Path to PowerPoint file"},
      "data_schema": {
        "type": "object",
        "properties": {
          "fields": {
            "type": "array",
            "items": {
              "type": "object",
              "properties": {
                "field_name": {"type": "string", "description": "Output field name"},
                "source_type": {"type": "string", "enum": ["title", "text", "table_cell", "shape_text", "notes"]},
                "extraction_rule": {"type": "string", "description": "Rule for extracting this field"},
                "data_type": {"type": "string", "enum": ["string", "number", "date", "boolean", "array"]},
                "transformation": {"type": "string", "description": "Optional data transformation rule"},
                "validation": {"type": "string", "description": "Validation rule for extracted data"}
              }
            }
          },
          "grouping": {"type": "string", "description": "How to group the extracted data"},
          "relationships": {"type": "array", "items": {"type": "object"}, "description": "Relationships between fields"}
        }
      },
      "slide_filter": {"type": "object", "description": "Criteria to filter slides"},
      "output_format": {"type": "string", "enum": ["json", "csv", "xml", "database_ready"], "default": "json"}
    },
    "required": ["file_path", "data_schema"]
  }
}
```

### 7. **compare_presentations**
**目的**: 複数プレゼンテーションの比較分析

```json
{
  "name": "compare_presentations",
  "description": "Compare multiple presentations to identify differences, similarities, and trends",
  "inputSchema": {
    "type": "object",
    "properties": {
      "file_paths": {"type": "array", "items": {"type": "string"}, "description": "List of PowerPoint files to compare"},
      "comparison_aspects": {
        "type": "array",
        "items": {"type": "string", "enum": ["content", "structure", "data_trends", "visual_consistency", "template_usage"]},
        "description": "Aspects to compare"
      },
      "baseline_file": {"type": "string", "description": "Optional baseline file for comparison"},
      "generate_report": {"type": "boolean", "default": true, "description": "Generate comparison report"},
      "highlight_changes": {"type": "boolean", "default": true, "description": "Highlight changes between versions"}
    },
    "required": ["file_paths", "comparison_aspects"]
  }
}
```

### 8. **generate_data_summary**
**目的**: 抽出データの要約とインサイト生成

```json
{
  "name": "generate_data_summary",
  "description": "Generate summaries and insights from extracted PowerPoint data",
  "inputSchema": {
    "type": "object",
    "properties": {
      "file_path": {"type": "string", "description": "Path to PowerPoint file"},
      "summary_type": {
        "type": "string",
        "enum": ["executive_summary", "data_insights", "trend_analysis", "key_metrics", "action_items"],
        "default": "executive_summary",
        "description": "Type of summary to generate"
      },
      "focus_areas": {
        "type": "array",
        "items": {"type": "string"},
        "description": "Specific areas to focus the summary on"
      },
      "include_visualizations": {"type": "boolean", "default": false, "description": "Include text-based visualizations"},
      "output_language": {"type": "string", "default": "ja", "description": "Output language (ja/en)"}
    },
    "required": ["file_path"]
  }
}
```

## 🔄 Implementation Priority

### Phase 1: Core Enhanced Tools
1. `search_slides_by_criteria` - 基本的な検索機能
2. `extract_tables_from_slides` - テーブル抽出の強化
3. `extract_progress_data` - 進捗データ専用抽出

### Phase 2: Batch Processing
4. `batch_extract_from_files` - 複数ファイル処理
5. `extract_structured_data` - 柔軟なデータ抽出

### Phase 3: Advanced Analytics
6. `analyze_slide_patterns` - パターン分析
7. `compare_presentations` - 比較分析
8. `generate_data_summary` - 要約生成

## 🎯 Usage Examples

### Example 1: 週次進捗レポートの自動抽出
```python
# 進捗スライドを検索
search_result = search_slides_by_criteria(
    file_path="weekly_report.pptx",
    criteria={
        "title_starts_with": "進捗",
        "has_tables": True
    }
)

# 進捗データを抽出
progress_data = extract_progress_data(
    file_path="weekly_report.pptx",
    progress_indicators={
        "title_patterns": ["進捗.*", ".*状況.*"],
        "status_keywords": ["完了", "進行中", "遅延", "未着手"],
        "percentage_extraction": True
    },
    output_structure="dashboard"
)
```

### Example 2: 複数ファイルからの一括データ抽出
```python
# 複数の月次レポートから同じ形式のデータを抽出
batch_data = batch_extract_from_files(
    file_paths=["2024-01.pptx", "2024-02.pptx", "2024-03.pptx"],
    extraction_template={
        "slide_criteria": {"title_contains": "売上"},
        "data_points": [
            {"name": "month", "type": "text", "extraction_rule": "title_date_extraction"},
            {"name": "revenue", "type": "number", "extraction_rule": "table_column:売上"},
            {"name": "target", "type": "number", "extraction_rule": "table_column:目標"}
        ]
    },
    consolidate_results=True
)
```

## 🔧 Technical Implementation Notes

### Data Processing Pipeline
1. **File Loading & Validation**
2. **Slide Filtering** (based on criteria)
3. **Content Extraction** (targeted extraction)
4. **Data Transformation** (formatting, validation)
5. **Result Aggregation** (consolidation, summarization)

### Performance Considerations
- **Lazy Loading**: 必要なスライドのみを処理
- **Caching**: 検索結果とパターンのキャッシュ
- **Parallel Processing**: 複数ファイル処理の並列化
- **Memory Management**: 大きなファイルの効率的な処理

### Error Handling
- **Graceful Degradation**: 一部のスライドでエラーが発生しても処理を継続
- **Detailed Error Reporting**: どのスライド/ファイルでエラーが発生したかを明確に報告
- **Recovery Suggestions**: エラーの解決方法を提案

この仕様により、PowerPoint MCP Serverは実際のビジネスユースケースに対応できる強力なツールセットを提供できます。