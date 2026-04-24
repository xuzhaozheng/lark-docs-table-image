---
name: lark-docs-table-image
description: 在飞书 Docx 中生成“表格单元格内嵌图片”的对比表。用于用户要求把本地 Markdown 表格（任意行列维度）写入飞书文档，且图片必须出现在表格 cell 内而非文末时。走原生 docx block API（三步法：create image block -> media upload -> replace_image），并避免 markdown replace 导致图片丢失。
---

# lark-docs-table-image

## 适用场景

当用户出现以下意图时使用本 skill：
- “把图片放进表格单元格”
- “还原 Markdown 表格含图效果到飞书文档”
- “不要文末图片索引，要 cell 内图片”
- “精确插入到某个表格位置”

## 核心结论（必须遵守）

1. 不要依赖 `docs +update --markdown` 的 HTML/Markdown 图片语法把图塞进 cell。
2. 要用 docx 原生 block API 操作单元格。
3. 不要在完成后用 `replace_all` 全文替换时间或文本；会触发重解析，可能清掉 cell 图片。
4. 表格图片的稳定链路是：
   - 在目标 `table_cell` 下创建空 `image block`（`block_type=27`）
   - `docs +media-upload` 上传图片到该 image block
   - `PATCH block` 的 `replace_image` 绑定 `file_token`

## 输入约定

最少需要：
- `doc_id`（或可解析为 doc_id 的 URL），或让脚本按标题新建文档
- 本地 Markdown 表格路径（任意“行维度 × 列维度”，图片使用 Markdown 图片语法）
- 图片根目录

建议参数：
- 图片尺寸上限（默认 `400x400`）
- 对齐（默认 `center`）
- 插入模式（`append | insert_after | insert_before`）
- 插入锚点标题（用于 `insert_after/insert_before`）
- 小节标题（默认不再写死“更新时间”）
- 是否清理 cell 内占位文本（默认清理）

## 位置与创建策略（新增）

- 新建文档时支持：
  - 云盘文件夹：`--folder-token`
  - 知识库：`--wiki-space` + `--wiki-node`
  - 文档标题：`--title`
- 更新已有文档时支持：
  - `--doc` 指定目标文档
  - `--insert-mode` + `--selection-by-title` 指定插入位置
- 表格渲染仅处理“本次新增 table block”，不会修改既有表格内容

## 标准流程

### 1) 创建或更新表格骨架

- 用 `docs +update --mode overwrite|append` 写入纯文本表格骨架。
- 图片单元格建议先放短占位（如 `A1`），后续会删除；也可以直接空白。

### 2) 获取表格与单元格 block_id

```bash
lark-cli api GET /open-apis/docx/v1/documents/<doc_id>/blocks \
  --params '{"document_revision_id":-1,"page_size":500}'
```

- 找到 `block_type=31` 的 table block
- 读取 `table.cells`（行优先顺序）

### 3) 对每个目标 cell 写入图片（三步法）

#### 3.1 在 cell 下创建空 image block

```bash
lark-cli api POST /open-apis/docx/v1/documents/<doc_id>/blocks/<table_cell_id>/descendant \
  --data '{
    "index": -1,
    "children_id": ["img_tmp_1"],
    "descendants": [
      {
        "block_id": "img_tmp_1",
        "block_type": 27,
        "image": {},
        "children": []
      }
    ]
  }'
```

记录返回里的真实 `image_block_id`（`block_id_relations`）。

#### 3.2 上传图片到 image block

```bash
lark-cli docs +media-upload \
  --doc-id <doc_id> \
  --parent-type docx_image \
  --parent-node <image_block_id> \
  --file <local_image_path>
```

记录返回 `file_token`。

#### 3.3 绑定 token 到 image block

```bash
lark-cli api PATCH /open-apis/docx/v1/documents/<doc_id>/blocks/<image_block_id> \
  --data '{
    "replace_image": {
      "token": "<file_token>",
      "width": <auto_or_custom_width>,
      "align": 2
    }
  }'
```

尺寸建议：
- 默认使用 `400x400` 作为显示上限。
- 实际写入时按原图比例等比缩放到 `400x400` 内，避免拉伸变形。
- 若某批图片比例特殊，可手动传 `--width` / `--height` 调整上限。

### 4) 清理占位文本（可选但推荐）

若 cell 下存在文本子块（`block_type=2`），删除它们，只保留图片块：

```bash
lark-cli api DELETE /open-apis/docx/v1/documents/<doc_id>/blocks/<table_cell_id>/children/batch_delete \
  --data '{"start_index":0,"end_index":1}'
```

循环删除直到该 cell 不再有文本子块。

### 5) 结果校验

- 再拉一次 blocks，确认目标 cell 的 children 仅包含 `block_type=27`。
- 抽样用 `docs +fetch` 检查结构是否是 `<lark-table> ... <image token="..."/>`。

## 常见错误与处理

- `1770025 operation and block not match`：
  - 对文本块调用了 `replace_image`。必须对 `image block` 调用。
- `1770013 relation mismatch`：
  - image block 与资源关系不对。需先 `media-upload` 到该 block，再 `replace_image`。
- `1061044 parent node not exist`：
  - `parent-node` 传错，必须传真实返回的 `image_block_id`。
- `UNSUPPORTED_HTML_TAG image/table`：
  - 触发了 markdown 解析路径，不是原生 block 更新路径。

## 禁止事项

- 禁止把本地路径直接写进 markdown 图片链接用于 cell 渲染。
- 禁止在最终完成后对整篇做 `replace_all`（尤其会改动 `<image>` 的替换）。
- 禁止把“图片插入文末再期望自动进 cell”作为最终方案。

## 输出给用户的结果格式

- 已处理文档：`doc_id / URL`
- 成功写入单元格数量：`N/N`
- 每个 cell 的结果：`cell_id -> image_block_id -> file_token`
- 若失败：给出具体 cell 与错误码，附重试建议
