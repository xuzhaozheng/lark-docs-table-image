#!/usr/bin/env python3
import argparse
import json
import re
import struct
import subprocess
import sys
from pathlib import Path

IMAGE_RE = re.compile(r"!\[[^\]]*\]\(([^)]+)\)")


def run_json(cmd):
    proc = subprocess.run(cmd, capture_output=True, text=True)
    if proc.returncode != 0:
        print(proc.stdout)
        print(proc.stderr, file=sys.stderr)
        raise RuntimeError("command failed: " + " ".join(cmd))
    return json.loads(proc.stdout)


def split_row(row):
    return [x.strip() for x in row.strip().strip("|").split("|")]


def parse_markdown_table(md_path):
    lines = [ln.strip() for ln in md_path.read_text(encoding="utf-8").splitlines() if ln.strip()]
    if len(lines) < 3:
        raise ValueError("markdown 表格内容不足")
    header = split_row(lines[0])
    rows = [split_row(ln) for ln in lines[2:]]
    return header, rows


def create_doc(title, folder_token=None, wiki_space=None, wiki_node=None):
    cmd = ["lark-cli", "docs", "+create", "--title", title, "--markdown", ""]
    if folder_token:
        cmd.extend(["--folder-token", folder_token])
    if wiki_space:
        cmd.extend(["--wiki-space", wiki_space])
    if wiki_node:
        cmd.extend(["--wiki-node", wiki_node])
    return run_json(cmd)


def update_doc_markdown(doc_id, markdown, mode, selection_by_title=None):
    cmd = [
        "lark-cli",
        "docs",
        "+update",
        "--doc",
        doc_id,
        "--mode",
        mode,
        "--markdown",
        markdown,
    ]
    if selection_by_title:
        cmd.extend(["--selection-by-title", selection_by_title])
    run_json(cmd)


def parse_doc_id(doc):
    doc = doc.strip()
    m = re.search(r"/docx/([A-Za-z0-9]+)", doc)
    if m:
        return m.group(1)
    if " " in doc:
        raise ValueError(f"非法 doc 标识：{doc}")
    return doc


def get_blocks(doc_id):
    return run_json(
        [
            "lark-cli",
            "api",
            "GET",
            f"/open-apis/docx/v1/documents/{doc_id}/blocks",
            "--params",
            '{"document_revision_id":-1,"page_size":500}',
        ]
    )["data"]["items"]


def delete_text_children(doc_id, cell_id):
    while True:
        cell = run_json(
            [
                "lark-cli",
                "api",
                "GET",
                f"/open-apis/docx/v1/documents/{doc_id}/blocks/{cell_id}",
            ]
        )["data"]["block"]
        children = cell.get("children", [])
        if not children:
            break
        first_id = children[0]
        first = run_json(
            [
                "lark-cli",
                "api",
                "GET",
                f"/open-apis/docx/v1/documents/{doc_id}/blocks/{first_id}",
            ]
        )["data"]["block"]
        if first.get("block_type") != 2:
            break
        run_json(
            [
                "lark-cli",
                "api",
                "DELETE",
                f"/open-apis/docx/v1/documents/{doc_id}/blocks/{cell_id}/children/batch_delete",
                "--data",
                '{"start_index":0,"end_index":1}',
            ]
        )


def get_png_size(path: Path):
    with path.open("rb") as f:
        sig = f.read(24)
    if len(sig) < 24 or sig[:8] != b"\x89PNG\r\n\x1a\n":
        return None
    w, h = struct.unpack(">II", sig[16:24])
    return int(w), int(h)


def get_jpeg_size(path: Path):
    with path.open("rb") as f:
        data = f.read()
    if len(data) < 4 or data[:2] != b"\xFF\xD8":
        return None
    i = 2
    while i < len(data) - 9:
        if data[i] != 0xFF:
            i += 1
            continue
        marker = data[i + 1]
        i += 2
        if marker in (0xD8, 0xD9):
            continue
        if i + 2 > len(data):
            break
        seg_len = int.from_bytes(data[i:i + 2], "big")
        if seg_len < 2 or i + seg_len > len(data):
            break
        if marker in (0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF):
            h = int.from_bytes(data[i + 3:i + 5], "big")
            w = int.from_bytes(data[i + 5:i + 7], "big")
            return int(w), int(h)
        i += seg_len
    return None


def get_image_size(path: Path):
    ext = path.suffix.lower()
    if ext == ".png":
        return get_png_size(path)
    if ext in (".jpg", ".jpeg"):
        return get_jpeg_size(path)
    return None


def fit_into_bounds(img_w, img_h, max_w, max_h):
    if img_w <= 0 or img_h <= 0:
        return max_w, max_h
    scale = min(max_w / img_w, max_h / img_h, 1.0)
    return max(1, int(img_w * scale)), max(1, int(img_h * scale))


DEFAULT_WIDTH = 400
DEFAULT_HEIGHT = 400


def main():
    parser = argparse.ArgumentParser(description="Create docx table image document")
    parser.add_argument("--md", required=True, help="markdown table file path")
    parser.add_argument("--doc", default="", help="existing target doc token/url; if empty, create new doc")
    parser.add_argument("--title", default="skill做的表", help="new doc title when --doc is empty")
    parser.add_argument("--folder-token", default="", help="create doc under drive folder token")
    parser.add_argument("--wiki-space", default="", help="create doc under wiki space id")
    parser.add_argument("--wiki-node", default="", help="create doc under wiki node token")
    parser.add_argument("--section-title", default="图片对比表", help="section title above the table")
    parser.add_argument(
        "--insert-mode",
        choices=["append", "insert_after", "insert_before"],
        default="append",
        help="where to place section/table in existing doc",
    )
    parser.add_argument(
        "--selection-by-title",
        default="",
        help="anchor title used with insert_before/insert_after (e.g. '## 原有章节')",
    )
    parser.add_argument("--width", type=int, default=DEFAULT_WIDTH, help="image width in cell (max bound)")
    parser.add_argument("--height", type=int, default=DEFAULT_HEIGHT, help="image height in cell (max bound)")
    parser.add_argument(
        "--image-align",
        choices=["left", "center", "right"],
        default="center",
        help="image alignment inside cell",
    )
    args = parser.parse_args()

    md_path = Path(args.md).resolve()
    if not md_path.exists():
        raise FileNotFoundError(str(md_path))

    header, rows = parse_markdown_table(md_path)
    if len(header) < 2:
        raise ValueError("表头至少包含 1 列行维度 + 1 列图片维度")
    col_count = len(header)
    max_width = args.width if args.width > 0 else DEFAULT_WIDTH
    max_height = args.height if args.height > 0 else DEFAULT_HEIGHT

    table_lines = []
    table_lines.append("| " + " | ".join(header) + " |")
    table_lines.append("|" + "---|" * len(header))
    for row in rows:
        device = row[0]
        blanks = [""] * (col_count - 1)
        table_lines.append("| " + " | ".join([device] + blanks) + " |")
    markdown = "### " + args.section_title + "\n\n" + "\n".join(table_lines)

    if args.doc:
        doc_id = parse_doc_id(args.doc)
    else:
        created = create_doc(
            args.title,
            folder_token=args.folder_token or None,
            wiki_space=args.wiki_space or None,
            wiki_node=args.wiki_node or None,
        )
        data = created.get("data", {})
        doc_id = data.get("doc_id")
        if not doc_id:
            raise RuntimeError("创建文档失败，未返回 doc_id")

    blocks_before = get_blocks(doc_id)
    before_ids = {b.get("block_id") for b in blocks_before}

    if args.insert_mode in ("insert_after", "insert_before"):
        if not args.selection_by_title:
            raise ValueError("insert_before/insert_after 模式必须提供 --selection-by-title")
        update_doc_markdown(
            doc_id,
            markdown,
            mode=args.insert_mode,
            selection_by_title=args.selection_by_title,
        )
    else:
        update_doc_markdown(doc_id, markdown, mode="append")

    blocks_after = get_blocks(doc_id)
    new_tables = [
        b
        for b in blocks_after
        if b.get("block_type") == 31 and b.get("block_id") not in before_ids
    ]
    table = new_tables[0] if new_tables else None
    if not table:
        raise RuntimeError("未找到新插入的 table block")
    cells = table["table"]["cells"]

    image_tasks = []
    for r, row in enumerate(rows, start=1):  # row 0 is header
        for c in range(1, col_count):
            cell_text = row[c] if c < len(row) else ""
            m = IMAGE_RE.search(cell_text)
            if not m:
                continue
            img_path = (md_path.parent / m.group(1)).resolve()
            if not img_path.exists():
                continue
            cell_index = r * col_count + c
            image_tasks.append((cell_index, img_path))

    align_map = {"left": 1, "center": 2, "right": 3}
    align_value = align_map[args.image_align]

    for idx, img_path in image_tasks:
        cell_id = cells[idx]
        temp_id = f"img_{idx}"
        created_child = run_json(
            [
                "lark-cli",
                "api",
                "POST",
                f"/open-apis/docx/v1/documents/{doc_id}/blocks/{cell_id}/descendant",
                "--data",
                json.dumps(
                    {
                        "index": -1,
                        "children_id": [temp_id],
                        "descendants": [
                            {
                                "block_id": temp_id,
                                "block_type": 27,
                                "image": {},
                                "children": [],
                            }
                        ],
                    },
                    ensure_ascii=False,
                ),
            ]
        )
        image_block_id = created_child["data"]["block_id_relations"][0]["block_id"]

        uploaded = run_json(
            [
                "lark-cli",
                "docs",
                "+media-upload",
                "--doc-id",
                doc_id,
                "--parent-type",
                "docx_image",
                "--parent-node",
                image_block_id,
                "--file",
                str(img_path.relative_to(Path.cwd())),
            ]
        )
        token = uploaded["data"]["file_token"]
        raw_size = get_image_size(img_path)
        if raw_size:
            width, height = fit_into_bounds(raw_size[0], raw_size[1], max_width, max_height)
        else:
            width, height = max_width, max_height

        run_json(
            [
                "lark-cli",
                "api",
                "PATCH",
                f"/open-apis/docx/v1/documents/{doc_id}/blocks/{image_block_id}",
                "--data",
                json.dumps(
                    {
                        "replace_image": {
                            "token": token,
                            "width": width,
                            "height": height,
                            "align": align_value,
                        }
                    },
                    ensure_ascii=False,
                ),
            ]
        )
        delete_text_children(doc_id, cell_id)

    print(
        json.dumps(
            {
                "ok": True,
                "doc_id": doc_id,
                "title": args.title,
                "section_title": args.section_title,
                "insert_mode": args.insert_mode,
                "selection_by_title": args.selection_by_title,
                "image_align": args.image_align,
                "max_width": max_width,
                "max_height": max_height,
            },
            ensure_ascii=False,
        )
    )


if __name__ == "__main__":
    main()
