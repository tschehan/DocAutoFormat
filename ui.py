"""Tkinter desktop UI for DocAutoFormat."""

from pathlib import Path
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from pipeline import run_pipeline


BG = "#f4efe7"
CARD = "#fffaf2"
INK = "#1f2937"
SUB = "#6b7280"
LINE = "#d6c9b8"
PRIMARY = "#1f5c4a"
PRIMARY_ACTIVE = "#18493b"
ACCENT = "#b7791f"
SUCCESS = "#21543d"
ERROR = "#8a2d2d"

PROGRESS_STEPS = [
    "后端启动",
    "提取文档内容",
    "文件结构识别",
    "理解修改要求",
    "生成执行计划",
    "修改文档",
]


def build_output_file_path(input_file: str) -> Path:
    """Build the generated filename from the source document name."""
    input_path = Path(input_file)
    return Path.cwd() / f"{input_path.stem}-formatted{input_path.suffix}"


def resource_path(relative_path: str) -> Path:
    """Resolve bundled resource paths for dev and PyInstaller builds."""
    base_path = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base_path / relative_path


root = tk.Tk()
root.title("DocAutoFormat")
root.geometry("920x700")
root.minsize(820, 620)
root.configure(bg=BG)
window_icon = resource_path("option_2_teal_wand.ico")
if window_icon.exists():
    root.iconbitmap(default=str(window_icon))

selected_file = tk.StringVar()
generated_file = tk.StringVar()
result_text = tk.StringVar(value="等待开始")
status_title = tk.StringVar(value="空闲")

progress_step_index = -1
status_lines: list[str] = []

style = ttk.Style()
style.theme_use("clam")
style.configure("App.TFrame", background=BG)
style.configure("Card.TFrame", background=CARD, relief="flat")
style.configure(
    "Title.TLabel",
    background=BG,
    foreground=INK,
    font=("Microsoft YaHei UI", 22, "bold"),
)
style.configure(
    "Subtitle.TLabel",
    background=BG,
    foreground=SUB,
    font=("Microsoft YaHei UI", 10),
)
style.configure(
    "Section.TLabel",
    background=CARD,
    foreground=INK,
    font=("Microsoft YaHei UI", 11, "bold"),
)
style.configure(
    "Hint.TLabel",
    background=CARD,
    foreground=SUB,
    font=("Microsoft YaHei UI", 9),
)
style.configure(
    "Path.TEntry",
    fieldbackground="#f8f4ed",
    foreground=INK,
    bordercolor=LINE,
    lightcolor=LINE,
    darkcolor=LINE,
)
style.configure("Choose.TButton", font=("Microsoft YaHei UI", 10))
style.configure(
    "Primary.TButton",
    font=("Microsoft YaHei UI", 11, "bold"),
    foreground="white",
    background=PRIMARY,
    borderwidth=0,
    focusthickness=3,
    focuscolor=PRIMARY,
    padding=(16, 10),
)
style.map(
    "Primary.TButton",
    background=[("active", PRIMARY_ACTIVE), ("disabled", "#9aa7a2")],
    foreground=[("disabled", "#f2f2f2")],
)


def refresh_status_text() -> None:
    """Render the status list."""
    status_text.configure(state="normal")
    status_text.delete("1.0", tk.END)
    if status_lines:
        status_text.insert(tk.END, "\n".join(f"- {line}" for line in status_lines))
        status_text.insert(tk.END, "\n")
    status_text.see(tk.END)
    status_text.configure(state="disabled")


def append_status(message: str) -> None:
    """Append one status line."""
    status_lines.append(message)
    refresh_status_text()


def replace_last_status(message: str) -> None:
    """Replace the latest status line."""
    if not status_lines:
        status_lines.append(message)
    else:
        status_lines[-1] = message
    refresh_status_text()


def complete_current_step() -> None:
    """Mark the current step as completed."""
    if 0 <= progress_step_index < len(PROGRESS_STEPS):
        replace_last_status(f"{PROGRESS_STEPS[progress_step_index]}完成 ✅")


def choose_file() -> None:
    """Choose a .docx file."""
    file_path = filedialog.askopenfilename(
        title="选择 Word 文件",
        filetypes=[("Word Document", "*.docx")],
    )
    if file_path:
        selected_file.set(file_path)
        generated_file.set("")
        set_result("已选择文件，等待开始处理", SUB)


def set_running_state(is_running: bool) -> None:
    """Enable or disable controls while working."""
    choose_button.configure(state="disabled" if is_running else "normal")
    start_button.configure(state="disabled" if is_running else "normal")


def set_result(message: str, color: str) -> None:
    """Update result panel."""
    result_text.set(message)
    result_value.configure(fg=color)


def on_progress(_message: str) -> None:
    """Receive progress from the backend and show a clean step-based status."""
    global progress_step_index

    next_index = progress_step_index + 1
    if next_index >= len(PROGRESS_STEPS):
        return

    if progress_step_index >= 0:
        complete_current_step()

    progress_step_index = next_index
    current_step = PROGRESS_STEPS[progress_step_index]
    status_title.set(f"{current_step}中...")
    append_status(f"{current_step}中......")


def show_result(result: dict) -> None:
    """Show final result."""
    set_running_state(False)

    if result.get("success"):
        complete_current_step()
        output_file = result.get("output_file", "")
        generated_file.set(output_file)
        set_result(f"任务完成\n输出文件: {output_file}", SUCCESS)
        status_title.set("任务完成 ✅")
        append_status("任务完成 ✅")
        return

    generated_file.set("")
    message = result.get("message", "未知错误")
    set_result(f"处理失败\n错误信息: {message}", ERROR)
    status_title.set("处理失败")
    if 0 <= progress_step_index < len(PROGRESS_STEPS):
        replace_last_status(f"{PROGRESS_STEPS[progress_step_index]}失败")
    append_status(f"任务失败：{message}")


def run_task(input_file: str, requirement_text: str, output_file: str) -> None:
    """Run pipeline in a background thread."""
    result = run_pipeline(
        input_file=input_file,
        requirement_text=requirement_text,
        output_file=output_file,
        progress_callback=on_progress,
    )
    root.after(0, lambda: show_result(result))


def start_processing() -> None:
    """Validate input and start work."""
    global progress_step_index

    input_file = selected_file.get().strip()
    requirement_text = requirement_input.get("1.0", tk.END).strip()

    if not input_file:
        messagebox.showwarning("提示", "请先选择 Word 文件。")
        return

    if not requirement_text:
        messagebox.showwarning("提示", "请输入修改要求。")
        return

    output_file = str(build_output_file_path(input_file))
    generated_file.set("")
    progress_step_index = -1
    status_lines.clear()
    refresh_status_text()

    status_title.set("准备开始")
    append_status("准备开始任务...")
    set_result("正在处理中，请稍等...", ACCENT)
    set_running_state(True)

    worker = threading.Thread(
        target=run_task,
        args=(input_file, requirement_text, output_file),
        daemon=True,
    )
    worker.start()


def on_canvas_configure(event: tk.Event) -> None:
    """Keep the scrollable frame width in sync with the canvas."""
    canvas.itemconfigure(scroll_window, width=event.width)


def on_frame_configure(_event: tk.Event) -> None:
    """Update the scroll region when content size changes."""
    canvas.configure(scrollregion=canvas.bbox("all"))


def on_mousewheel(event: tk.Event) -> str | None:
    """Scroll the window with the mouse wheel."""
    if event.delta == 0:
        return None
    canvas.yview_scroll(int(-event.delta / 120), "units")
    return "break"


outer_frame = ttk.Frame(root, style="App.TFrame")
outer_frame.pack(fill="both", expand=True)
outer_frame.columnconfigure(0, weight=1)
outer_frame.rowconfigure(0, weight=1)

canvas = tk.Canvas(outer_frame, bg=BG, highlightthickness=0, bd=0)
canvas.grid(row=0, column=0, sticky="nsew")

scrollbar = ttk.Scrollbar(outer_frame, orient="vertical", command=canvas.yview)
scrollbar.grid(row=0, column=1, sticky="ns")
canvas.configure(yscrollcommand=scrollbar.set)

main_frame = ttk.Frame(canvas, style="App.TFrame", padding=24)
scroll_window = canvas.create_window((0, 0), window=main_frame, anchor="nw")

canvas.bind("<Configure>", on_canvas_configure)
main_frame.bind("<Configure>", on_frame_configure)
canvas.bind_all("<MouseWheel>", on_mousewheel)

main_frame.columnconfigure(0, weight=1)
main_frame.rowconfigure(4, weight=1)

header = ttk.Frame(main_frame, style="App.TFrame")
header.grid(row=0, column=0, sticky="ew", pady=(0, 18))
header.columnconfigure(0, weight=1)

ttk.Label(header, text="DocAutoFormat", style="Title.TLabel").grid(
    row=0, column=0, sticky="w"
)
ttk.Label(
    header,
    text="选择 Word 文件，输入格式要求，然后生成格式化后的文档。",
    style="Subtitle.TLabel",
).grid(row=1, column=0, sticky="w", pady=(6, 0))

file_card = ttk.Frame(main_frame, style="Card.TFrame", padding=18)
file_card.grid(row=1, column=0, sticky="ew", pady=(0, 14))
file_card.columnconfigure(1, weight=1)

ttk.Label(file_card, text="Word 文件", style="Section.TLabel").grid(
    row=0, column=0, columnspan=2, sticky="w"
)
ttk.Label(
    file_card,
    text="选择一个需要处理的 .docx 文件。",
    style="Hint.TLabel",
).grid(row=1, column=0, columnspan=2, sticky="w", pady=(4, 12))

choose_button = ttk.Button(
    file_card,
    text="选择文件",
    style="Choose.TButton",
    command=choose_file,
)
choose_button.grid(row=2, column=0, sticky="w")

file_entry = ttk.Entry(
    file_card,
    textvariable=selected_file,
    state="readonly",
    style="Path.TEntry",
)
file_entry.grid(row=2, column=1, sticky="ew", padx=(12, 0))

requirement_card = ttk.Frame(main_frame, style="Card.TFrame", padding=18)
requirement_card.grid(row=2, column=0, sticky="nsew", pady=(0, 14))
requirement_card.columnconfigure(0, weight=1)

ttk.Label(requirement_card, text="修改要求", style="Section.TLabel").grid(
    row=0, column=0, sticky="w"
)
ttk.Label(
    requirement_card,
    text="直接输入自然语言要求，例如“标题黑体三号加粗居中，正文宋体小四”。",
    style="Hint.TLabel",
).grid(row=1, column=0, sticky="w", pady=(4, 12))

requirement_input = tk.Text(
    requirement_card,
    height=8,
    wrap="word",
    relief="solid",
    bd=1,
    bg="#fdfbf7",
    fg=INK,
    insertbackground=INK,
    font=("Microsoft YaHei UI", 11),
    highlightthickness=1,
    highlightbackground=LINE,
    highlightcolor=PRIMARY,
    padx=12,
    pady=10,
)
requirement_input.grid(row=2, column=0, sticky="ew")
requirement_input.insert(
    "1.0",
    "标题黑体三号加粗居中，正文宋体小四，1.5 倍行距，两端对齐。",
)

action_bar = ttk.Frame(main_frame, style="App.TFrame")
action_bar.grid(row=3, column=0, sticky="ew", pady=(0, 14))
action_bar.columnconfigure(1, weight=1)

start_button = ttk.Button(
    action_bar,
    text="开始处理",
    style="Primary.TButton",
    command=start_processing,
)
start_button.grid(row=0, column=0, sticky="w")

content_grid = ttk.Frame(main_frame, style="App.TFrame")
content_grid.grid(row=4, column=0, sticky="nsew")
content_grid.columnconfigure(0, weight=3)
content_grid.columnconfigure(1, weight=2)
content_grid.rowconfigure(0, weight=1)

status_card = ttk.Frame(content_grid, style="Card.TFrame", padding=18)
status_card.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
status_card.columnconfigure(0, weight=1)
status_card.rowconfigure(2, weight=1)

ttk.Label(status_card, text="运行状态", style="Section.TLabel").grid(
    row=0, column=0, sticky="w"
)
ttk.Label(status_card, textvariable=status_title, style="Hint.TLabel").grid(
    row=1, column=0, sticky="w", pady=(4, 12)
)

status_text = tk.Text(
    status_card,
    height=14,
    wrap="word",
    relief="solid",
    bd=1,
    state="disabled",
    bg="#fcfaf6",
    fg=INK,
    font=("Consolas", 10),
    highlightthickness=1,
    highlightbackground=LINE,
    padx=12,
    pady=10,
)
status_text.grid(row=2, column=0, sticky="nsew")

result_card = ttk.Frame(content_grid, style="Card.TFrame", padding=18)
result_card.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
result_card.columnconfigure(0, weight=1)

ttk.Label(result_card, text="处理结果", style="Section.TLabel").grid(
    row=0, column=0, sticky="w"
)
ttk.Label(
    result_card,
    text="处理完成后会显示生成文件路径。",
    style="Hint.TLabel",
).grid(row=1, column=0, sticky="w", pady=(4, 12))

result_panel = tk.Frame(
    result_card,
    bg="#f8f4ed",
    bd=1,
    relief="solid",
    highlightthickness=1,
    highlightbackground=LINE,
)
result_panel.grid(row=2, column=0, sticky="ew")

result_value = tk.Label(
    result_panel,
    textvariable=result_text,
    justify="left",
    anchor="w",
    bg="#f8f4ed",
    fg=SUB,
    font=("Microsoft YaHei UI", 11),
    padx=12,
    pady=14,
)
result_value.pack(fill="x")


if __name__ == "__main__":
    root.mainloop()
