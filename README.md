# excel_data_lineage_extractor

基于 Python 的 Excel 指标血缘提取工具，可读取名称管理器中的命名指标，并解析其公式依赖。

## 安装

```bash
pip install -e .
```

## 使用

```bash
excel-lineage /path/to/workbook.xlsx --format json
```

可选输出格式：
- `json`（默认）
- `markdown`
