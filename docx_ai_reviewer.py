import tkinter as tk
from tkinter import filedialog
import win32com.client
import requests
import os
import time

# -------------------------------
# 原始逻辑函数（增加 url 参数）
# -------------------------------

DEFAULT_SILICONFLOW_URL = "https://api.siliconflow.cn/v1/chat/completions"
DEFAULT_OLLAMA_URL = "http://localhost:11434/api/generate"
DEFAULT_PROMPT = r"你精通错别字修改和文法语法，后面的内容中，如果有错别字，请将错别字按照`a 应为 b`的格式进行罗列，如果有不当的词汇，请将不当用词按照`...看起来奇怪`进行罗列，如果有病句，指明按照`...疑似病句，病因是...`进行罗列，如果没有问题，请回答`no problem at all`，不要做其他回答："

def check_with_siliconflow(prompt, model_name, api_key="", url=DEFAULT_SILICONFLOW_URL):
    payload = {
        "model": model_name,
        "messages": [
            {
                "role": "system",
                "content": "你是一个负责检查论文内容的人工智能，专门检测错别字。"
            },
            {
                "role": "user",
                "content": prompt,
            }
        ],
        "stream": False,
        "max_tokens": 4096,
        "stop": None,
        "temperature": 0.6,
        "top_p": 0.95,
        "top_k": 50,
        "frequency_penalty": 0.5,
        "n": 1,
        "response_format": {"type": "text"},
    }
    headers = {
        "Authorization": "Bearer " + api_key,
        "Content-Type": "application/json"
    }
    response = requests.post(url, json=payload, headers=headers)
    return response.json()['choices'][0]['message']['content'], response.json()['usage']

def check_with_ollama(prompt, model_name, url=DEFAULT_OLLAMA_URL):
    payload = {
        "model": model_name,
        "prompt": prompt,
        "stream": False,
        "options": {
            "num_thread": 32
        }
    }
    response = requests.post(url, json=payload)
    return response.json()['response'], False

def add_comment_to_paragraph(doc, paragraph, comment_text, reviewer='Ai'):
    """
    在指定的段落上添加批注
    """
    range_ = paragraph.Range
    comment = doc.Comments.Add(range_, comment_text)
    comment.Author = reviewer
    comment.Initial = reviewer[:2]

def parse_page_range(range_str):
    """
    解析用户输入的页面范围字符串, 如 '3-5'
    返回 (start_page, end_page) 或 None
    """
    if not range_str.strip():
        return None  # 表示不限制页码
    # 简单只考虑 "start-end" 这种格式，如果需要更复杂的格式，可做进一步扩展
    parts = range_str.split('-')
    try:
        if len(parts) == 1:
            # 只有一个数字，比如 '3'
            page = int(parts[0])
            return (page, page)
        elif len(parts) == 2:
            start = int(parts[0])
            end = int(parts[1])
            if start > end:
                start, end = end, start  # 如果用户不小心写反了
            return (start, end)
        else:
            return None
    except ValueError:
        return None
    
def review_word_document(
    input_path,
    output_path,
    check_function,
    reviewer="Ai",
    model_name="deepseek-ai/DeepSeek-V3",
    prompt="",
    word_visible=True,
    page_range=None,
    log_callback=None
):
    """
    遍历 Word 文档中的所有段落，检测不当表达，并添加批注。
    """
    # 启动 Word
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = word_visible  # 根据用户设置显示/隐藏

    # 打开 Word 文档
    doc = word.Documents.Open(input_path)
    doc.SaveAs(output_path)  # 另存为输出文件

    # 获取 page_range (start, end)
    if page_range is not None:
        (start_page, end_page) = page_range

    total_paras = len(doc.Paragraphs)
    time_start = time.time()
    input_tokens = 0
    output_tokens = 0
    for i, para in enumerate(doc.Paragraphs):
        text = para.Range.Text.strip()

        if text:
            # 判断段落所在页是否在指定范围
            current_page = para.Range.Information(3)
            if page_range is not None:
                if not (start_page <= current_page <= end_page):
                    continue  # 不在范围内则跳过

            review_result, usage_count = check_function(prompt + f'```{text}```', model_name=model_name)

            if not 'no problem at all' in review_result:
                if '</think>' in review_result:
                    review_result = review_result.split('</think>')[-1]
                add_comment_to_paragraph(doc, para, review_result, reviewer=reviewer)
                progress_msg = f"当前页面 {current_page}，总进度 {i+1} / {total_paras}，总耗时 {int(time.time() - time_start)}s，已标注"
            else:
                progress_msg = f"当前页面 {current_page}，总进度 {i+1} / {total_paras}，总耗时 {int(time.time() - time_start)}s，无标注"

            if log_callback:
                log_callback(progress_msg)
            else:
                print(progress_msg)

    if usage_count:
        input_tokens += int(usage_count["prompt_tokens"])
        output_tokens += int(usage_count["completion_tokens"])
        progress_msg = f"总计输入 {input_tokens} tokens，总计输出 {output_tokens} tokens"
        if log_callback:
            log_callback(progress_msg)
        else:
            print(progress_msg)

    doc.SaveAs(output_path)
    doc.Close()
    # 处理结束后可视情况 word.Quit()

# -------------------------------
# 基于 tkinter 的 GUI
# -------------------------------

class ReviewGUI:
    def __init__(self, master):
        self.master = master
        self.master.title("Word 文档 LLM 检测工具 @Michayee")

        # ---------- Part 1: 输入/输出文件 ----------
        frame_file = tk.LabelFrame(master, text="文件选择", padx=10, pady=10)
        frame_file.pack(fill="x", padx=10, pady=5)

        tk.Label(frame_file, text="输入文件：", width=12, anchor="e").grid(row=0, column=0, sticky="w")
        self.input_file_var = tk.StringVar()
        tk.Entry(frame_file, textvariable=self.input_file_var, width=50).grid(row=0, column=1, padx=5)
        tk.Button(frame_file, text="浏览...", command=self.browse_input_file).grid(row=0, column=2, padx=5)

        tk.Label(frame_file, text="输出文件：", width=12, anchor="e").grid(row=1, column=0, sticky="w")
        self.output_file_var = tk.StringVar()
        tk.Entry(frame_file, textvariable=self.output_file_var, width=50).grid(row=1, column=1, padx=5)
        tk.Button(frame_file, text="浏览...", command=self.browse_output_file).grid(row=1, column=2, padx=5)

        frame_page_range = tk.LabelFrame(master, text="页面范围设置(可选)", padx=10, pady=10)
        frame_page_range.pack(fill="x", padx=10, pady=5)

        tk.Label(frame_page_range, text="页面范围：", width=12, anchor="e").grid(row=0, column=0, sticky="w")
        self.page_range_var = tk.StringVar()  # 默认为空表示处理全部
        tk.Entry(frame_page_range, textvariable=self.page_range_var, width=20).grid(row=0, column=1, padx=5)
        tk.Label(frame_page_range, text="例如：3-5 或 2，不填则为全部页面", anchor="w").grid(row=0, column=2, sticky="w")

        # ---------- Part 2: 检测函数选择 & Model Name & URL & API Key ----------
        frame_check = tk.LabelFrame(master, text="检测模型选择", padx=10, pady=10)
        frame_check.pack(fill="x", padx=10, pady=5)

        frame_check0 = tk.Frame(frame_check)
        frame_check0.pack(fill="x", padx=0, pady=5)

        frame_check1 = tk.Frame(frame_check)
        frame_check1.pack(fill="x", padx=0, pady=5)

        # 单选：siliconflow 或 ollama
        self.check_function_var = tk.StringVar(value="siliconflow")
        rb_sf = tk.Radiobutton(frame_check0, text="check_with_siliconflow",
                               variable=self.check_function_var, value="siliconflow",
                               command=self.on_check_function_changed)
        rb_sf.grid(row=0, column=0, sticky="w")

        rb_oll = tk.Radiobutton(frame_check0, text="check_with_ollama",
                                variable=self.check_function_var, value="ollama",
                                command=self.on_check_function_changed)
        rb_oll.grid(row=0, column=1, sticky="w")

        # URL
        tk.Label(frame_check1, text="URL：", width=12, anchor="e").grid(row=1, column=0, sticky="w")
        self.url_var = tk.StringVar(value=DEFAULT_SILICONFLOW_URL)  # 默认先给siliconflow的URL
        self.url_entry = tk.Entry(frame_check1, textvariable=self.url_var, width=50)
        self.url_entry.grid(row=1, column=1, padx=5)

        # Model Name
        tk.Label(frame_check1, text="Model Name：", width=12, anchor="e").grid(row=2, column=0, sticky="w")
        self.model_name_var = tk.StringVar(value="deepseek-ai/DeepSeek-V3")  # siliconflow默认
        self.model_name_entry = tk.Entry(frame_check1, textvariable=self.model_name_var, width=50)
        self.model_name_entry.grid(row=2, column=1, padx=5)

        # SiliconFlow API key（仅对 siliconflow 生效）
        tk.Label(frame_check1, text="API Key：", width=12, anchor="e").grid(row=3, column=0, sticky="w")
        self.api_key_var = tk.StringVar()
        self.api_key_entry = tk.Entry(frame_check1, textvariable=self.api_key_var, width=50)
        self.api_key_entry.grid(row=3, column=1, padx=5)

        # ---------- Part 3: Prompt & Reviewer ----------
        frame_review = tk.LabelFrame(master, text="批注名称 & 提示词", padx=10, pady=10)
        frame_review.pack(fill="x", padx=10, pady=5)

        frame_review0 = tk.Frame(frame_review)
        frame_review0.pack(fill="x", padx=0, pady=5)

        frame_review1 = tk.Frame(frame_review)
        frame_review1.pack(fill="x", padx=0, pady=5)

        # Reviewer 默认 "Ai"
        tk.Label(frame_review0, text="Reviewer：", width=12, anchor="e").grid(row=0, column=0, sticky="w")
        self.reviewer_var = tk.StringVar(value="Ai")
        tk.Entry(frame_review0, textvariable=self.reviewer_var, width=50).grid(row=0, column=1, padx=5)

        # Prompt
        tk.Label(frame_review1, text="Prompt：", width=12, anchor="e").grid(row=1, column=0, sticky="w")
        self.prompt_var = tk.StringVar(value=DEFAULT_PROMPT)
        tk.Entry(frame_review1, textvariable=self.prompt_var, width=50).grid(row=1, column=1, padx=5)

        # ---------- Part 4: 日志输出及开始检测 ----------
        frame_action = tk.Frame(master, padx=10, pady=10)
        frame_action.pack(fill="x", padx=10, pady=5)

        # 勾选项：是否前台显示 Word
        self.word_visible_var = tk.BooleanVar(value=True)
        cb_visible = tk.Checkbutton(frame_action, text="前台运行 Word（可以最小化）", variable=self.word_visible_var)
        cb_visible.pack(side='left', padx=10, pady=5)

        btn_start = tk.Button(frame_action, text="开始检测", command=self.run_review)
        btn_start.pack(side='left', padx=90, pady=5)

        # 日志输出文本框
        frame_output = tk.LabelFrame(master, text="日志输出", padx=10, pady=10)
        frame_output.pack(fill="both", expand=True, padx=10, pady=5)

        self.output_text = tk.Text(frame_output, wrap="word", height=10, width = 65)
        # self.output_text.pack(fill="both", expand=True)
        self.output_text.pack(fill="y", padx=10, pady=5)

        # frame_note = tk.Frame(master, padx=10, pady=10)
        tk.Label(master, text="@Michayee").pack(side = 'right', padx=10, pady=5)

        # 根据选择初始化可见性
        self.on_check_function_changed()

    # -------------- 事件响应 --------------

    def browse_input_file(self):
        file_path = filedialog.askopenfilename(
            title="选择输入 Word 文件",
            filetypes=[("Word 文件", "*.docx;*.doc")]
        )
        if file_path:
            self.input_file_var.set(file_path)

    def browse_output_file(self):
        file_path = filedialog.asksaveasfilename(
            title="选择输出 Word 文件",
            defaultextension=".docx",
            filetypes=[("Word 文件", "*.docx")]
        )
        if file_path:
            self.output_file_var.set(file_path)

    def on_check_function_changed(self):
        """
        根据当前选择的检测函数，更新默认 URL、Model Name，
        并控制 API key 的可编辑状态
        """
        choice = self.check_function_var.get()
        if choice == "siliconflow":
            self.url_var.set(DEFAULT_SILICONFLOW_URL)
            self.model_name_var.set("deepseek-ai/DeepSeek-V3")
            self.api_key_entry.config(state="normal")
        else:  # ollama
            self.url_var.set(DEFAULT_OLLAMA_URL)
            self.model_name_var.set("deepseek:70b")
            self.api_key_var.set("")  # 一般也不需要 API Key
            self.api_key_entry.config(state="disabled")

    def log(self, message):
        """
        往日志区域插入文字
        """
        self.output_text.insert(tk.END, message + "\n")
        self.output_text.see(tk.END)  # 自动滚动到最后

    def run_review(self):
        input_file = self.input_file_var.get()
        output_file = self.output_file_var.get()
        reviewer = self.reviewer_var.get()
        prompt = self.prompt_var.get()
        word_visible = self.word_visible_var.get()

        page_range_str = self.page_range_var.get().strip()
        pr = parse_page_range(page_range_str)  # 返回 (start, end) 或 None

        url = self.url_var.get()
        api_key = self.api_key_var.get()
        model_name = self.model_name_var.get()

        if not input_file or not os.path.isfile(input_file):
            self.log("错误：输入文件不存在或未选择！")
            return
        if not output_file:
            self.log("错误：请指定输出文件！")
            return

        # 根据选择构造检测函数
        if self.check_function_var.get() == "siliconflow":
            def local_check_function(p, model_name):
                return check_with_siliconflow(
                    p,
                    model_name=model_name,
                    api_key=api_key,
                    url=url
                )
        else:
            def local_check_function(p, model_name):
                return check_with_ollama(
                    p,
                    model_name=model_name,
                    url=url
                )

        self.log("开始检测，请稍候...")

        def log_callback(msg):
            self.log(msg)
            self.master.update()

        try:
            review_word_document(
                input_path=input_file,
                output_path=output_file,
                check_function=local_check_function,
                reviewer=reviewer,
                model_name=model_name,
                prompt=prompt,
                word_visible=word_visible,
                page_range=pr,
                log_callback=log_callback
            )
            self.log("检测完成！请在输出文件查看结果。")
        except Exception as e:
            self.log(f"检测过程中出现错误：{e}")

def main():
    root = tk.Tk()

    import sys
        # 获取 PyInstaller 运行时目录
    if getattr(sys, 'frozen', False):  # 检查是否被 PyInstaller 打包
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")

    root.iconbitmap(os.path.join(base_path, "ico_docx_ai_review.ico"))
    ReviewGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
