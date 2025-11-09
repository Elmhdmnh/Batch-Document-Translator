import os
import threading
import queue
import time
import traceback
from tkinter import Tk, Button, Label, Entry, StringVar, filedialog, ttk, Text, END, N, S, E, W
import docx2txt
from docx import Document
import requests


class TranslatorGUI:
    def __init__(self, root):
        self.root = root
        root.title("批量文档翻译器")
        root.geometry("820x520")

        # 状态/配置变量
        self.api_key_var = StringVar()
        self.base_url_var = StringVar(value="https://api.openai.com/v1")
        self.model_var = StringVar(value="gpt-3.5-turbo")
        self.direction_var = StringVar(value="auto")
        self.output_dir_var = StringVar(value=os.getcwd())

        # 文件队列
        self.file_list = []
        self.task_queue = queue.Queue()

        # UI 布局
        self._build_ui()

        # 控制
        self.stop_flag = threading.Event()

    def _build_ui(self):
        # 左侧：配置
        Label(self.root, text="API Key:").grid(row=0, column=0, sticky=W, padx=8, pady=6)
        Entry(self.root, textvariable=self.api_key_var, width=60, show='*').grid(row=0, column=1, columnspan=3, sticky=W)

        Label(self.root, text="Base URL:").grid(row=1, column=0, sticky=W, padx=8, pady=6)
        Entry(self.root, textvariable=self.base_url_var, width=60).grid(row=1, column=1, columnspan=3, sticky=W)

        Label(self.root, text="Model:").grid(row=2, column=0, sticky=W, padx=8, pady=6)
        Entry(self.root, textvariable=self.model_var, width=30).grid(row=2, column=1, sticky=W)

        Label(self.root, text="翻译方向:").grid(row=3, column=0, sticky=W, padx=8, pady=6)
        direction_cb = ttk.Combobox(self.root, textvariable=self.direction_var, values=["auto", "en->zh", "zh->en"], state='readonly', width=10)
        direction_cb.grid(row=3, column=1, sticky=W)

        Label(self.root, text="输出目录:").grid(row=4, column=0, sticky=W, padx=8, pady=6)
        Entry(self.root, textvariable=self.output_dir_var, width=60).grid(row=4, column=1, columnspan=2, sticky=W)
        Button(self.root, text="选择目录", command=self.choose_output_dir).grid(row=4, column=3, sticky=W)

        # 按钮
        Button(self.root, text="选择文件（多选）", command=self.select_files, width=20).grid(row=5, column=0, padx=8, pady=10)
        Button(self.root, text="开始翻译", command=self.start_translation, width=20).grid(row=5, column=1, padx=8)
        Button(self.root, text="停止", command=self.stop_translation, width=10).grid(row=5, column=2)
        Button(self.root, text="清空日志", command=self.clear_log, width=10).grid(row=5, column=3)

        # 进度条和文件列表
        self.progress = ttk.Progressbar(self.root, orient='horizontal', length=600, mode='determinate')
        self.progress.grid(row=6, column=0, columnspan=4, padx=8, pady=6)

        Label(self.root, text="已选择文件:").grid(row=7, column=0, sticky=W, padx=8)
        self.files_text = Text(self.root, height=6, width=95)
        self.files_text.grid(row=8, column=0, columnspan=4, padx=8)

        Label(self.root, text="日志:").grid(row=9, column=0, sticky=W, padx=8)
        self.log_text = Text(self.root, height=8, width=95)
        self.log_text.grid(row=10, column=0, columnspan=4, padx=8, pady=8)

    def choose_output_dir(self):
        d = filedialog.askdirectory()
        if d:
            self.output_dir_var.set(d)

    def select_files(self):
        paths = filedialog.askopenfilenames(title="选择 Word 文档", filetypes=[("Word 文档", "*.docx *.doc")])
        if not paths:
            return
        self.file_list = list(paths)
        self.files_text.delete(1.0, END)
        for p in self.file_list:
            self.files_text.insert(END, p + "\n")
        self.log(f"选择了 {len(self.file_list)} 个文件")

    def log(self, *parts):
        t = time.strftime('%Y-%m-%d %H:%M:%S')
        self.log_text.insert(END, f"[{t}] " + " ".join(map(str, parts)) + "\n")
        self.log_text.see(END)

    def clear_log(self):
        self.log_text.delete(1.0, END)

    def stop_translation(self):
        self.stop_flag.set()
        self.log("用户请求停止。等待当前任务完成...")

    def start_translation(self):
        if not self.file_list:
            self.log("错误：未选择文件")
            return
        api_key = self.api_key_var.get().strip()
        if not api_key:
            self.log("警告：未填写 API Key（仍可继续但大部分 API 会返回 401）")

        # reset flags
        self.stop_flag.clear()
        self.progress['maximum'] = len(self.file_list)
        self.progress['value'] = 0

        # 启动后台线程
        t = threading.Thread(target=self._worker_thread, daemon=True)
        t.start()

    def _worker_thread(self):
        total = len(self.file_list)
        for idx, path in enumerate(self.file_list, start=1):
            if self.stop_flag.is_set():
                self.log("已停止，退出任务循环。")
                break
            try:
                self.log(f"开始翻译 {os.path.basename(path)} ({idx}/{total})")
                self._translate_file(path)
                self.log(f"完成：{os.path.basename(path)}")
            except Exception as e:
                self.log(f"翻译失败：{os.path.basename(path)} -> {e}")
                traceback.print_exc()
            finally:
                self.progress['value'] = idx
        self.log("所有任务处理完成或已停止。")

    def _translate_file(self, filepath):
        # 检查文件扩展名
        ext = os.path.splitext(filepath)[1].lower()
        
        try:
            if ext == '.docx':
                # 使用 docx2txt 处理 .docx 文件
                self.log(f"使用 docx2txt 处理 .docx 文件: {os.path.basename(filepath)}")
                raw_text = docx2txt.process(filepath)
            elif ext == '.doc':
                # 对于 .doc 文件，使用 win32com (仅限 Windows)
                self.log(f"使用 win32com 处理 .doc 文件: {os.path.basename(filepath)}")
                try:
                    import win32com.client
                    word_app = win32com.client.Dispatch("Word.Application")
                    word_app.Visible = False
                    doc = word_app.Documents.Open(filepath)
                    raw_text = doc.Content.Text
                    doc.Close()
                    word_app.Quit()
                except ImportError:
                    raise ValueError("处理 .doc 文件需要安装 pywin32 库：pip install pywin32")
                except Exception as e:
                    raise ValueError(f"使用 Word 处理 .doc 文件失败: {str(e)}")
            else:
                raise ValueError(f"不支持的文件格式: {ext}")
                    
        except Exception as e:
            raise ValueError(f"读取文件失败: {str(e)}")
        
        if not raw_text or raw_text.strip() == "":
            raise ValueError("读取到的文件为空或无法提取文本。")

        # 构造 prompt
        direction = self.direction_var.get()
        if direction == 'auto':
            system_prompt = "你是一个专业的翻译人员，翻译时遵守信达雅原则，自动识别源语言并翻译为目标语言（如果源语言是英文，翻译为中文；如果源语言是中文，翻译为英文）。保留原文段落结构。"
        elif direction == 'en->zh':
            system_prompt = "你是一个专业的翻译人员，请把下面的英文文本翻译为中文，翻译要信、达、雅，保留段落结构。"
        elif direction == 'zh->en':
            system_prompt = "你是一个专业的翻译人员，请把下面的中文文本翻译为英文，翻译要信、达、雅，保留段落结构。"
        else:
            system_prompt = "你是翻译人员，按要求翻译文本。"

        # Prepare API call
        api_key = self.api_key_var.get().strip()
        base_url = self.base_url_var.get().strip().rstrip('/')
        model = self.model_var.get().strip()

        # Some endpoints expect /v1/chat/completions, others /chat/completions
        possible_endpoints = [f"{base_url}/v1/chat/completions", f"{base_url}/chat/completions", base_url]

        headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"} if api_key else {"Content-Type": "application/json"}

        messages = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": raw_text}
        ]

        payload = {"model": model, "messages": messages}

        # Try endpoints until one works or fail
        resp_json = None
        last_exc = None
        for ep in possible_endpoints:
            try:
                self.log(f"调用 API: {ep} (model={model})")
                r = requests.post(ep, json=payload, headers=headers, timeout=120)
                if r.status_code == 200:
                    resp_json = r.json()
                    break
                else:
                    # 记录错误并继续尝试
                    self.log(f"HTTP {r.status_code}: {r.text[:400]}")
            except Exception as e:
                last_exc = e
                self.log(f"调用失败：{e}")

        if resp_json is None:
            raise RuntimeError(f"API 请求失败，最后一次异常：{last_exc}")

        # 解析响应（兼容 OpenAI 格式和类似格式）
        translated = None
        try:
            # OpenAI style: resp_json['choices'][0]['message']['content']
            translated = resp_json['choices'][0]['message']['content']
        except Exception:
            # 其它可能的字段
            try:
                translated = resp_json['choices'][0]['text']
            except Exception:
                # 如果没有 choices，则尝试直接查找 'content'
                translated = resp_json.get('content') or resp_json.get('translation')

        if not translated:
            raise RuntimeError('无法从 API 响应中提取翻译内容。响应摘要：' + str(list(resp_json.keys())[:10]))

        # 保存结果：TXT + docx
        out_dir = self.output_dir_var.get() or os.path.dirname(filepath)
        os.makedirs(out_dir, exist_ok=True)
        base = os.path.splitext(os.path.basename(filepath))[0]

        txt_path = os.path.join(out_dir, base + '_translated.txt')
        with open(txt_path, 'w', encoding='utf-8') as f:
            f.write(translated)

        # 将翻译文本按段落拆分并写入 docx
        doc = Document()
        # 简单根据换行来分段，保持段落结构的一致性
        for para in translated.split('\n'):
            doc.add_paragraph(para)
        docx_out = os.path.join(out_dir, base + '_translated.docx')
        doc.save(docx_out)

        self.log(f"保存：{txt_path}")
        self.log(f"保存：{docx_out}")


if __name__ == '__main__':
    root = Tk()
    app = TranslatorGUI(root)
    root.mainloop()