# Lark-Table Skill Test

这个仓库用于测试 `lark-docs-table-image` skill，把本地 Markdown 表格（含图片）写入飞书 Docx 表格单元格。

## 目录

- `skills/lark-docs-table-image/`: skill 定义与脚本

## 主要脚本

`skills/lark-docs-table-image/scripts/run.py` 支持：

- 新建文档（可指定云盘文件夹或知识库节点）
- 向已有文档指定标题前/后插入横评块
- 自定义小标题
- 设置图片在单元格中的对齐方式（left/center/right）

## 示例

新建到个人文档库：

```bash
uv run "skills/lark-docs-table-image/scripts/run.py" \
  --md "pic2-device-filter-compare.md" \
  --title "新的测试" \
  --wiki-space "my_library" \
  --section-title "2026-0424第一次尝试"
```

插入到已有文档正文最前（通过标题锚点）：

```bash
uv run "skills/lark-docs-table-image/scripts/run.py" \
  --md "device-style-compare.md" \
  --doc "O5KGwjRfniBukVkTrj0c9JiKnND" \
  --insert-mode insert_before \
  --selection-by-title "### 2026-0424第一次尝试" \
  --section-title "第二次横评"
```
