# Coze 变量清单

## 输入变量

```text
task_id
product_name
analysis_type
input_mode
period_start
period_end
baseline_files
chat_files
user_message
uploaded_files
model
```

## 门禁变量

```text
inherit_labels
dedupe_by_buyer_stage_label_period
p0_quality_safety_second_review
excel_acceptance_gate
```

## 中间变量

```text
raw_messages
clean_messages
valid_buyer_messages
need_items
deduped_need_items
duplicate_records
risk_candidates
risk_reviewed_items
label_dictionary
summary_findings
excel_file
markdown_file
manifest_file
```

## 输出变量

```text
status
task_id
message
summary
files.excel
files.markdown
files.manifest
```

## 返回 JSON

```json
{
  "status": "completed",
  "task_id": "{{task_id}}",
  "message": "分析已完成",
  "summary": [],
  "files": {
    "excel": "",
    "markdown": "",
    "manifest": ""
  }
}
```
