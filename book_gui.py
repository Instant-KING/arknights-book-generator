"""
明日方舟剧情书籍生成器 —— 傻瓜式GUI
- 从Wiki同步最新剧情
- 选章节预览
- 生成精美docx书籍
"""
import os
import sys
import re
import time
import threading
import tkinter as tk
from tkinter import ttk, messagebox

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)

from pathlib import Path

from generate_docx import batch_generate
import scrape_arknights_story as scraper

_data_dir_a = os.path.join(SCRIPT_DIR, "arknights_dialogue")
_data_dir_b = os.path.join(os.getcwd(), "arknights_dialogue")
DATA_DIR = _data_dir_a if os.path.isdir(_data_dir_a) else _data_dir_b


# =====================================================================
# 同步进度窗口（弹出式）
# =====================================================================

class SyncWindow(tk.Toplevel):
    def __init__(self, parent, data_dir):
        super().__init__(parent)
        self.title("同步最新剧情")
        self.geometry("600x450")
        self.minsize(500, 300)
        self.transient(parent)
        self.grab_set()

        self.data_dir = Path(data_dir)
        self.cancel_event = threading.Event()
        self.is_running = False
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        # 顶部说明
        info = ttk.Label(self, text="正在从 prts.wiki 拉取最新剧情数据...\n新内容会自动保存，已有内容不会丢失。",
                         font=("等线", 10), justify=tk.LEFT)
        info.pack(fill=tk.X, padx=12, pady=(12, 6))

        # 日志文本框
        log_frame = ttk.Frame(self)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 8))

        self.log_text = tk.Text(log_frame, wrap=tk.WORD, font=("Consolas", 9),
                                state=tk.DISABLED, padx=8, pady=8,
                                bg="#1e1e1e", fg="#d4d4d4", insertbackground="white")
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # 底部按钮 + 进度条
        bottom = ttk.Frame(self)
        bottom.pack(fill=tk.X, padx=12, pady=(0, 12))

        self.progress = ttk.Progressbar(bottom, mode="determinate")
        self.progress.pack(fill=tk.X, pady=(0, 8))

        btn_frame = ttk.Frame(bottom)
        btn_frame.pack(fill=tk.X)

        self.status_label = ttk.Label(btn_frame, text="准备中...", font=("等线", 9))
        self.status_label.pack(side=tk.LEFT)

        self.close_btn = ttk.Button(btn_frame, text="关闭", command=self._on_close,
                                    state=tk.DISABLED)
        self.close_btn.pack(side=tk.RIGHT, padx=4)

        self.cancel_btn = ttk.Button(btn_frame, text="取消同步", command=self._cancel)
        self.cancel_btn.pack(side=tk.RIGHT, padx=4)

        # 启动
        self._start()

    def _log(self, msg):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)

    def _start(self):
        self.is_running = True
        self._log("◆ 正在获取剧情页面列表...")
        thread = threading.Thread(target=self._run_sync, daemon=True)
        thread.start()

    def _run_sync(self):
        try:
            # Step 1: 获取模板页面
            self._update_status("获取页面列表...")
            template_text = scraper.fetch_text(scraper.TEMPLATE_URL)
            all_entries = scraper.extract_sections_and_pages(template_text)
            total = len(all_entries)

            # 预先计算每个条目对应的本地文件路径，判断是否已存在
            needed = []  # [(group, section, page_name, title, filepath)]
            skipped = 0
            for group, section, page_name, title in all_entries:
                folder = (self.data_dir / scraper.safe_filename(group)
                          / scraper.safe_filename(section))
                filepath = folder / (scraper.safe_filename(title) + ".txt")
                if filepath.exists():
                    skipped += 1
                else:
                    needed.append((group, section, page_name, title, filepath))

            new_count = len(needed)
            self._log(f"◆ 共 {total} 个页面，已有 {skipped} 个，需拉取 {new_count} 个")
            self._log("")

            if new_count == 0:
                self._log("◆ 所有剧情已是最新，无需同步！")
                return

            self.progress["maximum"] = new_count

            success = 0
            empty = 0
            errors = 0
            delay = scraper.REQUEST_DELAY

            for i, (group, section, page_name, title, filepath) in enumerate(needed):
                if self.cancel_event.is_set():
                    self._log("\n◆ 用户取消同步")
                    break

                self._update_status(f"[{i+1}/{new_count}] {title}")
                self.progress["value"] = i + 1

                try:
                    dialogues = scraper.scrape_page(page_name)
                    content = scraper.format_dialogue(dialogues)
                    header = f"# {title}\n# Page: {page_name}\n# Dialogues: {len(dialogues)}\n\n"

                    filepath.parent.mkdir(parents=True, exist_ok=True)
                    with open(filepath, "w", encoding="utf-8") as f:
                        f.write(header)
                        f.write(content)

                    if dialogues:
                        self._log(f"  OK [{group}/{section}] {title}  ({len(dialogues)}句)")
                        success += 1
                    else:
                        self._log(f"  -- [{group}/{section}] {title}  (空)")
                        empty += 1

                except Exception as e:
                    self._log(f"  ✗ [{group}/{section}] {title}  错误: {e}")
                    errors += 1

                if delay > 0 and i < new_count - 1 and not self.cancel_event.is_set():
                    time.sleep(delay)

            self._log("")
            self._log(f"◆ 同步完成！成功: {success}  空章节: {empty}  失败: {errors}")

        except Exception as e:
            self._log(f"\n◆ 同步出错: {e}")

        finally:
            self.is_running = False
            self.after(0, self._on_done)

    def _update_status(self, text):
        self.after(0, lambda: self.status_label.config(text=text))

    def _on_done(self):
        self.progress.stop()
        self.cancel_btn.config(state=tk.DISABLED)
        self.close_btn.config(state=tk.NORMAL)
        self.status_label.config(text="同步结束")
        # 通知主窗口刷新
        if hasattr(self.master, '_on_sync_finished'):
            self.master._on_sync_finished()

    def _cancel(self):
        if self.is_running:
            self.cancel_event.set()
            self._log("\n◆ 正在取消...（等待当前页面完成）")
            self.cancel_btn.config(state=tk.DISABLED)

    def _on_close(self):
        if self.is_running:
            if messagebox.askyesno("确认", "同步正在进行中，确定要关闭吗？\n关闭后同步会停止。"):
                self.cancel_event.set()
                self.destroy()
        else:
            self.destroy()


# =====================================================================
# 主界面
# =====================================================================

class BookGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("明日方舟剧情书籍生成器")
        self.root.geometry("900x680")
        self.root.minsize(800, 580)

        style = ttk.Style()
        style.theme_use("clam")

        # ---- 顶部标题 ----
        title_frame = ttk.Frame(root)
        title_frame.pack(fill=tk.X, padx=15, pady=(15, 5))
        ttk.Label(title_frame, text="明日方舟剧情书籍生成器",
                  font=("等线", 18, "bold")).pack()
        ttk.Label(title_frame, text="同步Wiki → 选章节 → 一键生成书籍",
                  font=("等线", 10)).pack()

        # ---- 主区域：左右分栏 ----
        main_frame = ttk.Frame(root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)

        # 左侧：类别 + 章节列表
        left_frame = ttk.LabelFrame(main_frame, text="选择章节", padding=5)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=False, padx=(0, 10))

        ttk.Label(left_frame, text="剧情分类:").pack(anchor=tk.W, pady=(5, 0))
        self.category_var = tk.StringVar()
        self.category_combo = ttk.Combobox(left_frame, textvariable=self.category_var,
                                           state="readonly", width=30)
        self.category_combo.pack(fill=tk.X, pady=(0, 5))
        self.category_combo.bind("<<ComboboxSelected>>", self._on_category_select)

        ttk.Label(left_frame, text="章节:").pack(anchor=tk.W)
        self.chapter_listbox = tk.Listbox(left_frame, width=32, height=20,
                                          exportselection=False)
        self.chapter_listbox.pack(fill=tk.BOTH, expand=True, pady=(0, 5))
        self.chapter_listbox.bind("<<ListboxSelect>>", self._on_chapter_select)

        ttk.Button(left_frame, text="刷新列表", command=self._refresh_data).pack(fill=tk.X)

        # 右侧：预览 + 操作
        right_frame = ttk.Frame(main_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # 操作按钮区（放在预览上方）
        action_frame = ttk.Frame(right_frame)
        action_frame.pack(fill=tk.X, pady=(0, 8))

        self.info_label = ttk.Label(action_frame, text="共 0 个剧情文件")
        self.info_label.pack(side=tk.LEFT, padx=5)

        self.sync_btn = ttk.Button(action_frame, text="从Wiki同步最新剧情",
                                   command=self._open_sync_window)
        self.sync_btn.pack(side=tk.RIGHT, padx=5)

        self.generate_btn = ttk.Button(action_frame, text="生成书籍 (docx)",
                                       command=self._generate_book)
        self.generate_btn.pack(side=tk.RIGHT, padx=5)

        self.output_btn = ttk.Button(action_frame, text="打开输出目录",
                                     command=self._open_output_dir)
        self.output_btn.pack(side=tk.RIGHT, padx=5)

        # 预览区
        preview_frame = ttk.LabelFrame(right_frame, text="内容预览", padding=5)
        preview_frame.pack(fill=tk.BOTH, expand=True)

        self.preview_text = tk.Text(preview_frame, wrap=tk.WORD, font=("等线", 10),
                                    state=tk.DISABLED, padx=8, pady=8)
        self.preview_text.pack(fill=tk.BOTH, expand=True)

        # ---- 底部：手动追加内容 ----
        add_frame = ttk.LabelFrame(root, text="手动追加内容（选填）", padding=8)
        add_frame.pack(fill=tk.X, padx=15, pady=(5, 15))

        add_row1 = ttk.Frame(add_frame)
        add_row1.pack(fill=tk.X, pady=(0, 5))

        ttk.Label(add_row1, text="说话人:").pack(side=tk.LEFT)
        self.speaker_entry = ttk.Entry(add_row1, width=18, font=("等线", 10))
        self.speaker_entry.pack(side=tk.LEFT, padx=(5, 20))

        ttk.Label(add_row1, text="对话内容:").pack(side=tk.LEFT)
        self.dialogue_entry = ttk.Entry(add_row1, width=50, font=("等线", 10))
        self.dialogue_entry.pack(side=tk.LEFT, padx=(5, 10))

        self.add_dialogue_btn = ttk.Button(add_row1, text="添加为对话",
                                           command=self._add_dialogue)
        self.add_dialogue_btn.pack(side=tk.LEFT, padx=2)

        add_row2 = ttk.Frame(add_frame)
        add_row2.pack(fill=tk.X)

        ttk.Label(add_row2, text="旁白文本:").pack(side=tk.LEFT)
        self.narrative_entry = ttk.Entry(add_row2, width=50, font=("等线", 10))
        self.narrative_entry.pack(side=tk.LEFT, padx=(5, 10))

        self.add_narrative_btn = ttk.Button(add_row2, text="添加为旁白",
                                            command=self._add_narrative)
        self.add_narrative_btn.pack(side=tk.LEFT, padx=2)

        self.add_status = ttk.Label(add_row2, text="", foreground="green")
        self.add_status.pack(side=tk.LEFT, padx=15)

        # ---- 状态栏 ----
        self.status_var = tk.StringVar(value="就绪")
        status_bar = ttk.Label(root, textvariable=self.status_var,
                               relief=tk.SUNKEN, anchor=tk.W, padding=4)
        status_bar.pack(fill=tk.X, side=tk.BOTTOM)

        # ---- 进度条 ----
        self.progress = ttk.Progressbar(root, mode="indeterminate")

        # ---- 初始化数据 ----
        self.chapters = {}  # {category: [(name, txt_count, path), ...]}
        self.current_chapter_path = None
        self._refresh_data()

    # ================================================================
    # Wiki 同步
    # ================================================================

    def _open_sync_window(self):
        self.sync_btn.config(state=tk.DISABLED, text="同步中...")
        SyncWindow(self.root, DATA_DIR)
        # SyncWindow 完成后会调用 _on_sync_finished

    def _on_sync_finished(self):
        self.sync_btn.config(state=tk.NORMAL, text="从Wiki同步最新剧情")
        self._refresh_data()
        self.status_var.set("Wiki同步完成，列表已刷新")

    # ================================================================
    # 数据刷新
    # ================================================================

    def _refresh_data(self):
        self.chapters.clear()
        if not os.path.isdir(DATA_DIR):
            self.status_var.set("未找到数据目录: " + DATA_DIR)
            return

        for category in os.listdir(DATA_DIR):
            cat_path = os.path.join(DATA_DIR, category)
            if not os.path.isdir(cat_path):
                continue
            chapters = []
            for ch in os.listdir(cat_path):
                ch_path = os.path.join(cat_path, ch)
                if os.path.isdir(ch_path):
                    txt_count = len([f for f in os.listdir(ch_path)
                                     if f.endswith(".txt")])
                    if txt_count > 0:
                        chapters.append((ch, txt_count, ch_path))
            if chapters:
                chapters.sort(key=lambda x: x[0])
                self.chapters[category] = chapters

        categories = sorted(self.chapters.keys())
        self.category_combo["values"] = categories
        if categories:
            self.category_combo.current(0)
            self._on_category_select()

        self.status_var.set(f"已加载 {len(categories)} 个分类, 共 {sum(len(v) for v in self.chapters.values())} 个章节")

    def _on_category_select(self, event=None):
        cat = self.category_var.get()
        self.chapter_listbox.delete(0, tk.END)
        if cat not in self.chapters:
            return
        for ch_name, txt_count, _ch_path in self.chapters[cat]:
            self.chapter_listbox.insert(tk.END, f"{ch_name}  ({txt_count}篇)")

    def _on_chapter_select(self, event=None):
        sel = self.chapter_listbox.curselection()
        if not sel:
            return
        cat = self.category_var.get()
        idx = sel[0]
        ch_name, txt_count, ch_path = self.chapters[cat][idx]
        self.current_chapter_path = ch_path
        self.info_label.config(text=f"共 {txt_count} 个剧情文件")

        lines = []
        total_chars = 0
        txt_files = sorted(
            [f for f in os.listdir(ch_path) if f.endswith(".txt")],
            key=self._sort_key
        )
        for fname in txt_files:
            fpath = os.path.join(ch_path, fname)
            try:
                with open(fpath, "r", encoding="utf-8") as f:
                    content = f.read()
                    lines.append(f"--- {fname} ---")
                    lines.append(content)
                    total_chars += len(content)
                    if total_chars > 8000:
                        lines.append("\n... (内容过长，已截断预览)")
                        break
            except Exception:
                lines.append(f"[无法读取: {fname}]")

        preview = "\n".join(lines)
        self.preview_text.config(state=tk.NORMAL)
        self.preview_text.delete("1.0", tk.END)
        self.preview_text.insert("1.0", preview)
        self.preview_text.config(state=tk.DISABLED)
        self.status_var.set(f"预览: {ch_name} ({txt_count}篇, 约{total_chars}字)")

    @staticmethod
    def _sort_key(filename):
        name = os.path.splitext(filename)[0]
        m = re.match(r"(\d+)-(\d+)\s", name)
        if m:
            major, minor = int(m.group(1)), int(m.group(2))
            has_end = 1 if "后" in name or "END" in name.upper() else 0
            return (major, minor, has_end, name)
        m2 = re.match(r"(\d+)", name)
        if m2:
            return (int(m2.group(1)), 0, 0, name)
        return (9999, 0, 0, name)

    # ================================================================
    # 生成书籍
    # ================================================================

    def _generate_book(self):
        if not self.current_chapter_path:
            messagebox.showwarning("提示", "请先在左侧选择一个章节！")
            return

        ch_path = self.current_chapter_path
        ch_name = os.path.basename(ch_path)
        parent_dir = os.path.dirname(ch_path)
        output_path = os.path.join(parent_dir, f"{ch_name}.docx")

        if os.path.exists(output_path):
            if not messagebox.askyesno("确认", f'"{ch_name}.docx" 已存在，要覆盖吗？'):
                return

        self._set_busy(True)
        self.status_var.set(f"正在生成 {ch_name}.docx ...")

        def _run():
            try:
                result = batch_generate(ch_path, output_path)
                self.root.after(0, lambda: self._on_generate_done(result))
            except Exception as e:
                self.root.after(0, lambda: self._on_generate_error(str(e)))

        threading.Thread(target=_run, daemon=True).start()

    def _on_generate_done(self, result):
        self._set_busy(False)
        if result:
            self.status_var.set(f"生成完成: {result}")
            if messagebox.askyesno("完成", f"书籍已生成！\n\n{result}\n\n要打开所在文件夹吗？"):
                os.startfile(os.path.dirname(result))
        else:
            self.status_var.set("生成失败，请检查文件夹中是否有txt文件")

    def _on_generate_error(self, err_msg):
        self._set_busy(False)
        self.status_var.set("生成出错")
        messagebox.showerror("错误", f"生成过程中出错:\n{err_msg}")

    def _set_busy(self, busy):
        if busy:
            self.generate_btn.config(state=tk.DISABLED, text="生成中...")
            self.progress.pack(fill=tk.X, side=tk.BOTTOM)
            self.progress.start(10)
        else:
            self.generate_btn.config(state=tk.NORMAL, text="生成书籍 (docx)")
            self.progress.stop()
            self.progress.pack_forget()

    # ================================================================
    # 手动追加内容
    # ================================================================

    def _add_dialogue(self):
        speaker = self.speaker_entry.get().strip()
        text = self.dialogue_entry.get().strip()
        if not speaker or not text:
            self.add_status.config(text="请填写说话人和对话内容", foreground="red")
            return
        filepath = self._get_append_filepath()
        if filepath is None:
            return
        with open(filepath, "a", encoding="utf-8") as f:
            f.write(f"\n{speaker}:{text}")
        self.dialogue_entry.delete(0, tk.END)
        self.add_status.config(text='已添加对话！点击「生成书籍」即可更新', foreground="green")
        self._on_chapter_select()

    def _add_narrative(self):
        text = self.narrative_entry.get().strip()
        if not text:
            self.add_status.config(text="请填写旁白文本", foreground="red")
            return
        filepath = self._get_append_filepath()
        if filepath is None:
            return
        with open(filepath, "a", encoding="utf-8") as f:
            f.write(f"\n{text}")
        self.narrative_entry.delete(0, tk.END)
        self.add_status.config(text='已添加旁白！点击「生成书籍」即可更新', foreground="green")
        self._on_chapter_select()

    def _get_append_filepath(self):
        if not self.current_chapter_path:
            messagebox.showwarning("提示", "请先在左侧选择一个章节！")
            return None
        txt_files = sorted(
            [f for f in os.listdir(self.current_chapter_path) if f.endswith(".txt")],
            key=self._sort_key
        )
        if not txt_files:
            messagebox.showwarning("提示", "该章节下暂无txt文件")
            return None
        return os.path.join(self.current_chapter_path, txt_files[-1])

    # ================================================================
    # 工具
    # ================================================================

    def _open_output_dir(self):
        if self.current_chapter_path:
            os.startfile(os.path.dirname(self.current_chapter_path))
        elif os.path.isdir(DATA_DIR):
            os.startfile(DATA_DIR)
        else:
            os.startfile(SCRIPT_DIR)


def main():
    root = tk.Tk()
    BookGeneratorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
