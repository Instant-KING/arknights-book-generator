#!/usr/bin/env python3
"""
明日方舟剧情爬虫 - 从 prts.wiki 抓取所有剧情对话

输出目录结构:
    主线剧情一览/
        特殊/
            隐藏剧情.txt
            采购中心.txt
            ...
        黑暗时代·上/
            序章·上.txt
            0-1 坍塌 行动前.txt
            ...
        黑暗时代·下/
            ...
    活动剧情一览/
        多维合作/
            ...
        危机合约/
            ...

用法:
    python scrape_arknights_story.py                      # 抓取全部
    python scrape_arknights_story.py --list-only          # 仅列出结构
    python scrape_arknights_story.py --limit 5            # 测试前5个
    python scrape_arknights_story.py --page "采购中心/剧情" # 抓取指定页面
    python scrape_arknights_story.py --combine            # 合并为一个文件
"""

import re
import sys
import time
import argparse
import urllib.parse
import urllib.request
from pathlib import Path

BASE_URL = "https://prts.wiki"
TEMPLATE_URL = f"{BASE_URL}/w/Template:%E5%89%A7%E6%83%85%E5%AF%BC%E8%88%AA?action=raw"
RAW_URL = f"{BASE_URL}/w/{{page}}?action=raw"
USER_AGENT = "prts-story-scraper/1.0 (contact@example.com)"
REQUEST_DELAY = 1.2
MAX_RETRIES = 3

OUTPUT_DIR = Path("arknights_dialogue")

GENERIC_SUBS = {"主线", "剧情"}


def safe_print(*args, **kwargs):
    try:
        print(*args, **kwargs)
    except UnicodeEncodeError:
        text = " ".join(str(a) for a in args)
        print(text.encode("ascii", errors="replace").decode("ascii"), **kwargs)


def fetch_text(url: str, retries: int = MAX_RETRIES) -> str:
    for attempt in range(retries):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
            with urllib.request.urlopen(req, timeout=30) as resp:
                return resp.read().decode("utf-8", errors="replace")
        except Exception as e:
            if attempt == retries - 1:
                raise
            time.sleep(2 * (attempt + 1))


def _extract_after_pipe(text: str) -> str:
    """从 wikitext 单元格提取内容，处理 [[link|text]]
    输入格式: ! <attrs>|<内容>  或  ! <内容>（无管道）
    """
    # 去掉开头的 ! 标记
    if text.startswith("!"):
        text = text[1:].strip()
    pipe_pos = text.find("|")
    if pipe_pos >= 0:
        content = text[pipe_pos + 1:]
    else:
        content = text
    # 从链接中提取显示文本: [[page|显示文本]] -> 显示文本
    m = re.search(r"\[\[[^\[\]|]+\|([^\[\]]+)\]\]", content)
    if m:
        return m.group(1).strip()
    # 纯文本，去掉残留的 wikitext 标记
    content = re.sub(r"\[\[|\]\]|\{\{|\}\}", "", content)
    return content.strip()


def _get_rowspan(line: str) -> int:
    """从 wikitext 行提取 rowspan 数值"""
    m = re.search(r'rowspan\s*=\s*"(\d+)"', line)
    return int(m.group(1)) if m else 1


def extract_sections_and_pages(template_text: str) -> list[tuple[str, str, str, str]]:
    """
    解析剧情导航模板
    返回 [(group, section_folder, page_name, display_title), ...]

    模板行结构（每个 row 由 |- 分隔）：
      主线:
        ! <attrs>|<章节名>      ← section
        ! <attrs>|<子分类>      ← sub (通常是"主线")
        | <attrs>|<页面链接>    ← data
      活动 (含 rowspan):
        ! rowspan="N"|<大分类>  ← section (跨 N 行)
        ! <子活动名>            ← sub
        | <页面链接>            ← data
        |-                     ← 下一行，section 沿用
        ! <子活动名2>           ← sub
        | <页面链接2>           ← data
    """
    results = []
    seen = set()

    current_group = ""
    # 带 rowspan 的 section 状态
    pending_section = ""
    pending_section_left = 0

    # 按 |- 拆分行（保留分隔符来判断行边界）
    # 先按行拆分，然后以 |- 为边界分组
    lines = template_text.split("\n")

    # 收集当前行的 ! 单元和 | 单元
    row_bang_cells: list[str] = []  # ! 单元格的文本（已提取 | 后内容）
    row_data_cells: list[str] = []  # | 单元格的文本
    row_bang_rowspans: list[int] = []

    def flush_row():
        """处理累积的一行数据"""
        nonlocal pending_section, pending_section_left

        if not row_data_cells:
            return

        num_bangs = len(row_bang_cells)

        if pending_section_left > 0:
            # 上一行的 rowspan section 仍生效，本行 ! 是 sub
            current_section = pending_section
            current_sub = row_bang_cells[0] if num_bangs >= 1 else ""
            pending_section_left -= 1
            if pending_section_left == 0:
                pending_section = ""
        elif num_bangs >= 1:
            # 新 section
            current_section = row_bang_cells[0]
            current_sub = row_bang_cells[1] if num_bangs >= 2 else ""
            rs = row_bang_rowspans[0] if row_bang_rowspans else 1
            if rs > 1:
                pending_section = current_section
                pending_section_left = rs - 1
        else:
            current_section = ""
            current_sub = ""

        # 确定文件夹名
        if current_sub and current_sub not in GENERIC_SUBS:
            folder = current_sub
        else:
            folder = current_section or current_sub or "Other"

        # 提取所有 data 单元格中的页面链接
        for data_cell in row_data_cells:
            for m in re.finditer(r"\[\[([^\[\]|]+)(?:\|([^\[\]]+))?\]\]", data_cell):
                page = m.group(1).strip()
                title = m.group(2).strip() if m.group(2) else page
                if page not in seen:
                    seen.add(page)
                    results.append((current_group, folder, page, title))

    for line in lines:
        line = line.strip()

        # 跳过空行
        if not line:
            continue

        # 表头：检测 group 标题
        group_match = re.search(r"\[\[剧情一览\|([^\]]+)\]\]", line)
        if group_match:
            current_group = group_match.group(1).strip()
            continue

        # 行分隔符 |- 或 表结束 |}
        if line.startswith("|-"):
            flush_row()
            row_bang_cells = []
            row_data_cells = []
            row_bang_rowspans = []
            continue

        if line.startswith("|}"):
            flush_row()
            row_bang_cells = []
            row_data_cells = []
            row_bang_rowspans = []
            pending_section = ""
            pending_section_left = 0
            continue

        # 跳过表格声明行
        if line.startswith("{|"):
            continue

        # 跳过 colspan 标题行
        if line.startswith("! colspan"):
            continue

        # ! 标题单元格
        if line.startswith("!"):
            rs = _get_rowspan(line)
            text = _extract_after_pipe(line)
            row_bang_cells.append(text)
            row_bang_rowspans.append(rs)
            continue

        # | 数据单元格
        if line.startswith("|"):
            row_data_cells.append(line)
            continue

    # 处理最后一行（模板结束时可能没有 |} 或 |-）
    flush_row()

    return results


def parse_dialogue(raw_text: str) -> list[tuple[str, str]]:
    """从 raw wikitext 中解析对话，返回 [(name, text), ...]"""
    match = re.search(r"\{\{剧情模拟器\|", raw_text)
    if not match:
        return []

    start = match.start()
    depth = 0
    end = start
    i = start
    while i < len(raw_text) - 1:
        if raw_text[i:i+2] == "{{":
            depth += 1
            i += 2
            continue
        elif raw_text[i:i+2] == "}}":
            depth -= 1
            if depth == 0:
                end = i
                break
            i += 2
            continue
        i += 1

    block = raw_text[start:end]

    text_data_match = re.search(r"文本数据=\s*\n?", block)
    if not text_data_match:
        return []

    text_data = block[text_data_match.end():]

    dialogues = []
    for line in text_data.split("\n"):
        line = line.strip()
        if not line:
            continue

        name_match = re.match(r"\[name\s*=\s*\"([^\"]*)\"\]\s*(.*)", line)
        if name_match:
            name = name_match.group(1)
            text = name_match.group(2).strip()
            if text:
                dialogues.append((name, text))
            continue

        sub_match = re.match(r'\[Subtitle\(text\s*=\s*"((?:[^"\\]|\\.)*)"', line)
        if sub_match:
            text = sub_match.group(1)
            if text:
                dialogues.append(("", text))

    return dialogues


def format_dialogue(dialogues: list[tuple[str, str]]) -> str:
    lines = []
    for name, text in dialogues:
        if name:
            lines.append(f"{name}:{text}")
        else:
            lines.append(text)
    return "\n".join(lines)


def scrape_page(page_name: str) -> list[tuple[str, str]]:
    url = RAW_URL.format(page=urllib.parse.quote(page_name, safe="/"))
    raw_text = fetch_text(url)
    return parse_dialogue(raw_text)


def safe_filename(name: str) -> str:
    return re.sub(r'[\\/*?:"<>|]', "_", name)


def main():
    parser = argparse.ArgumentParser(description="明日方舟剧情爬虫")
    parser.add_argument("--list-only", action="store_true", help="仅列出目录结构")
    parser.add_argument("--limit", type=int, default=0, help="只抓取前N个页面")
    parser.add_argument("--page", type=str, default="", help="只抓取指定页面")
    parser.add_argument("--output", type=str, default=str(OUTPUT_DIR), help="输出根目录")
    parser.add_argument("--combine", action="store_true", help="合并所有对话到一个文件")
    parser.add_argument("--no-delay", action="store_true", help="不等待")
    args = parser.parse_args()

    delay = 0 if args.no_delay else REQUEST_DELAY
    output_dir = Path(args.output)
    output_dir.mkdir(parents=True, exist_ok=True)

    safe_print("Fetching story page list...")
    template_text = fetch_text(TEMPLATE_URL)
    all_entries = extract_sections_and_pages(template_text)
    safe_print(f"Found {len(all_entries)} story pages")

    if args.page:
        all_entries = [("", "", args.page, args.page)]

    if args.list_only:
        last_group = ""
        last_section = ""
        for i, (group, section, page, title) in enumerate(all_entries, 1):
            if group != last_group:
                safe_print(f"\n{'='*50}")
                safe_print(f"  {group}")
                safe_print(f"{'='*50}")
                last_group = group
                last_section = ""
            if section != last_section:
                safe_print(f"  [{section}]")
                last_section = section
            safe_print(f"    {i:4d}. {title}")
        return

    if args.limit > 0:
        all_entries = all_entries[:args.limit]

    total = len(all_entries)
    success_count = 0
    empty_count = 0
    error_count = 0

    combine_blocks: dict[str, list[str]] = {}

    for i, (group, section, page_name, title) in enumerate(all_entries, 1):
        safe_print(f"[{i}/{total}] [{group}/{section}] {title} ... ", end="")
        sys.stdout.flush()

        try:
            dialogues = scrape_page(page_name)
            content = format_dialogue(dialogues)
            header = f"# {title}\n# Page: {page_name}\n# Dialogues: {len(dialogues)}\n\n"

            if args.combine:
                key = f"{group}/{section}"
                if key not in combine_blocks:
                    combine_blocks[key] = []
                combine_blocks[key].append(header + content)
            else:
                folder = output_dir / safe_filename(group) / safe_filename(section)
                folder.mkdir(parents=True, exist_ok=True)
                filename = safe_filename(title) + ".txt"
                filepath = folder / filename
                with open(filepath, "w", encoding="utf-8") as f:
                    f.write(header)
                    f.write(content)

            if dialogues:
                safe_print(f"OK ({len(dialogues)} lines)")
                success_count += 1
            else:
                safe_print("empty (header only)")
                empty_count += 1

        except Exception as e:
            safe_print(f"ERROR: {e}")
            error_count += 1

        if delay > 0 and i < total:
            time.sleep(delay)

    if args.combine and combine_blocks:
        combine_path = output_dir / "all_dialogues.txt"
        with open(combine_path, "w", encoding="utf-8") as f:
            for key in sorted(combine_blocks.keys()):
                f.write(f"\n{'='*60}\n")
                f.write(f"  {key}\n")
                f.write(f"{'='*60}\n\n")
                f.write("\n\n".join(combine_blocks[key]))
                f.write("\n\n")
        safe_print(f"\nCombined output: {combine_path}")

    safe_print(f"\nDone! Success: {success_count}, No data: {empty_count}, Error: {error_count}")


if __name__ == "__main__":
    main()
