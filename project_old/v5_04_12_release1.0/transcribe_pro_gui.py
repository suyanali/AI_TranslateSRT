
# transcribe_pro_gui.py (v1.7 - Refactored)

import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox
from tkinter import font as tkfont
import os
import sys
import subprocess
import threading
import queue
import multiprocessing
import runpy
from datetime import datetime

# ==============================================================================
#  PyInstaller 路徑解決方案
# ==============================================================================

def resource_path(relative_path):
    """ 獲取打包後資源的絕對路徑 """
    try:
        # PyInstaller 建立一個臨時資料夾並將路徑儲存在 _MEIPASS 中
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# ==============================================================================
#  智慧字體選擇
# ==============================================================================

def get_preferred_font(root_window):
    """ 從預設列表中選擇一個可用的字體 """
    font_priority = [
        'PingFang TC', 'Microsoft JhengHei', 'Noto Sans CJK TC', 
        'Source Han Sans TC', 'Heiti TC', 'PingFang SC', 
        'Microsoft YaHei', 'sans-serif'
    ]
    available_fonts = set(tkfont.families(root_window))
    for font in font_priority:
        if font in available_fonts:
            return font
    return 'sans-serif'

# ==============================================================================
# 全域設定
# ==============================================================================

# *** 修改: 直接指向新的整合腳本 ***
TRANSCRIPTION_SCRIPT_NAME = "transcribe_pro_v5_branch_04_branch_12.py"

DEFAULT_PROMPT_TEMPLATE = """你是一位頂級的 AI 語音轉文字專家，專精於生成和格式化完全符合行業標準的 SRT 字幕檔案。你的輸出必須精確無誤。

SRT 結構的關鍵規則（必須毫無例外地嚴格遵守）：

序列號：一個從 1 開始並嚴格遞增的整數。

時間碼：
格式絕對必須為 hh:mm:ss,xxx (小時:分鐘:秒,毫秒)。這是不可協商的。
小時部分 (hh) 即使為零，也必須顯示為 00。例如，1 分 5 秒 9 毫秒應表示為 00:01:05,009，絕不能省略小時部分而寫成 01:05,009。此規則適用於檔案中的每一個時間碼。
分鐘 (mm) 和秒 (ss) 若不足兩位數，必須以 0 在前面補齊（例如 00:01:05,009）。
毫秒 (xxx) 必須為三位數，若不足三位數，必須以 0 在尾部補齊（例如 00:00:01,050 而不是 00:00:01,50）。
開始時間和結束時間之間必須使用 --> (一個空格，兩個減號，一個大於號，一個空格) 分隔。
時間碼行本身前後不得有任何多餘空格或字元。

字幕文字：
核心原則：在每一個字幕塊中，此字幕文字內容必須且只能顯示為單獨的一行。嚴禁在單個字幕塊的字幕文字部分內部產生換行符或顯示為多行。

遵循下述「字幕生成規則」。
空行：每個字幕塊 (包含序列號、時間碼、字幕文字) 之後，必須有一個且只有一個完整的空行 將其與下一個字幕塊分隔開。這是SRT格式的基礎。
換行符：序列號行、時間碼行、以及字幕文字行，這三者各自作為獨立的行，他們之間必須使用標準換行符分隔。

字幕生成規則（請按優先級順序執行）：

第一優先：單行顯示與長度限制
重申核心原則：每一字幕塊的字幕文字部分，必須且只能是一行。
為確保此單行字幕易于閱讀，其文字長度絕對不能超過 15 個{language}字元。
如果原始口語語句的自然長度超過 15 個字元，或者其語義停頓點暗示需要分行，則該原始語句必須被切分為 數個新的、獨立的字幕塊。每一個新切分出的字幕塊都將擁有全新的序列號和對應的時間碼，並且其自身的字幕文字部分嚴格遵守單行和 15 字元內的長度限制。

第二優先：語意分段與新字幕塊的創建
當因上述「單行顯示與長度限制」原則需要切分原始口語語句時，切分點應選擇在最自然的語意停頓處（例如，在一個短語或子句結束後）。
每一次有效的切分都意味著結束當前字幕塊，並為切分後的下一段文字創建一個全新的字幕塊。 這樣，原本可能導致在同一字幕塊內換行的內容，會被合理分配到連續的多個單行字幕塊中。

第三優先：時間長度
每一段字幕（即每一個單行字幕塊）顯示的目標時長應在 3 到 5 秒 之間。
請根據實際的語速和語氣停頓來微調，確保字幕與聲音同步。
如果為了滿足單行和 15 字長度限制而切分出的字幕塊文字非常短，其顯示時間可以略短於 3 秒（例如 1-2 秒），但應盡量避免過於頻繁的極短字幕。

第四優先：內容淨化與精簡
自動省略沒有意義的語助詞（如：嗯、啊、呃）、口吃或不影響語意的重複詞語（如：那個、那個）。
字幕行中不包含任何標點符號（如：，。？！）。

第五優先：不需要(音樂)(歌曲)這種備註，有對白或歌詞都需提取成srt
{terms_section}
最終指令：
請根據提供的音檔，直接生成一份完全符合上述所有 SRT 結構和字幕生成規則的{language}字幕內容。再次強調，SRT 格式的精確性至關重要，特別是 確保每個時間碼都包含 hh: 部分 (即使是 00:)、每個字幕塊的字幕文字部分只有一行、時間碼的逗號、毫秒補零以及字幕塊之間的空行。最後輸出為 srt 檔案下載。"""
TERMS_TEMPLATE = """
第六優先：角色名稱對應翻譯與性別
請嚴格按照對應翻譯使用角色名稱，不得任意更改或混用。根據性別使用合適的敬稱（如先生、小姐），記住人稱代詞「你」、「妳」→ 全部翻為「你」；「他」、「她」→ 全部翻為「他」：
{terms_list}
"""

# ==============================================================================
# GUI 主應用程式類別
# ==============================================================================

class TranscriptionApp:
    def __init__(self, master):
        self.master = master
        self.master.title(f"字幕轉錄工具 v1.7 (核心: {TRANSCRIPTION_SCRIPT_NAME})")
        self.master.geometry("850x800")

        self.is_running = False
        self.process = None
        self.log_queue = queue.Queue()
        self.temp_prompt_file = None

        self._setup_styles_and_fonts()

        main_frame = ttk.Frame(master, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        self._create_widgets(main_frame)
        self._check_dependencies()

    def _setup_styles_and_fonts(self):
        self.preferred_font = get_preferred_font(self.master)
        style = ttk.Style(self.master)
        style.configure("TButton", padding=6, relief="flat", font=(self.preferred_font, 10))
        style.configure("TLabel", font=(self.preferred_font, 10))
        style.configure("TEntry", font=(self.preferred_font, 10))
        style.configure("TNotebook.Tab", font=(self.preferred_font, 10, 'bold'))
        style.configure("TCheckbutton", font=(self.preferred_font, 10))

    def _create_widgets(self, parent):
        # --- 1. 檔案選擇區 ---
        file_frame = ttk.LabelFrame(parent, text=" 1. 選擇來源檔案 ", padding="10")
        file_frame.pack(fill=tk.X, padx=5, pady=5)
        self.file_path_var = tk.StringVar()
        self.file_path_entry = ttk.Entry(file_frame, textvariable=self.file_path_var, state="readonly", width=80)
        self.file_path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        self.browse_button = ttk.Button(file_frame, text="選擇檔案...", command=self._select_file)
        self.browse_button.pack(side=tk.LEFT)

        # --- 2. Prompt 設定區 ---
        prompt_frame = ttk.LabelFrame(parent, text=" 2. 設定轉錄與翻譯規則 (將合併成 Prompt) ", padding="10")
        prompt_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        notebook = ttk.Notebook(prompt_frame)
        notebook.pack(fill=tk.BOTH, expand=True)
        
        tab1 = ttk.Frame(notebook, padding="5")
        notebook.add(tab1, text='主要規則')
        self.main_rules_text = scrolledtext.ScrolledText(tab1, wrap=tk.WORD, height=10, font=(self.preferred_font, 10))
        self.main_rules_text.pack(fill=tk.BOTH, expand=True)
        self.main_rules_text.insert(tk.END, DEFAULT_PROMPT_TEMPLATE)
        
        tab2 = ttk.Frame(notebook, padding="5")
        notebook.add(tab2, text='進階設定')
        lang_frame = ttk.Frame(tab2)
        lang_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(lang_frame, text="目標語言 (必填):").pack(side=tk.LEFT, padx=(0, 5))
        self.language_var = tk.StringVar(value="繁體中文")
        self.language_entry = ttk.Entry(lang_frame, textvariable=self.language_var, width=20)
        self.language_entry.pack(side=tk.LEFT)
        terms_frame = ttk.LabelFrame(tab2, text="人名或術語對應表 (選填，每行一組，例如: Gemini = 雙子星)", padding="5")
        terms_frame.pack(fill=tk.BOTH, expand=True)
        self.terms_text = scrolledtext.ScrolledText(terms_frame, wrap=tk.WORD, height=5, font=(self.preferred_font, 10))
        self.terms_text.pack(fill=tk.BOTH, expand=True)
        self.terms_text.insert(tk.END, """William = 威廉
Sherlock = 夏洛克""")

        # --- 3. 參數設定區 ---
        params_frame = ttk.LabelFrame(parent, text=" 3. 執行參數設定 ", padding="10")
        params_frame.pack(fill=tk.X, padx=5, pady=5)
        params_frame.columnconfigure(1, weight=1)
        params_frame.columnconfigure(3, weight=1)
        
        ttk.Label(params_frame, text="API Key (可選):").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.api_key_var = tk.StringVar()
        self.api_key_entry = ttk.Entry(params_frame, textvariable=self.api_key_var, width=40)
        self.api_key_entry.grid(row=0, column=1, sticky="ew", padx=5, pady=2)
        
        ttk.Label(params_frame, text="模型名稱:").grid(row=0, column=2, sticky="w", padx=5, pady=2)
        self.model_name_var = tk.StringVar(value="models/gemini-2.5-pro")
        self.model_name_entry = ttk.Entry(params_frame, textvariable=self.model_name_var, width=30)
        self.model_name_entry.grid(row=0, column=3, sticky="ew", padx=5, pady=2)
        
        ttk.Label(params_frame, text="分段時長 (秒):").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        self.chunk_duration_var = tk.StringVar(value="600")
        self.chunk_duration_entry = ttk.Entry(params_frame, textvariable=self.chunk_duration_var, width=15)
        self.chunk_duration_entry.grid(row=1, column=1, sticky="w", padx=5, pady=2)
        
        ttk.Label(params_frame, text="暫存資料夾:").grid(row=1, column=2, sticky="w", padx=5, pady=2)
        self.temp_dir_var = tk.StringVar(value=os.path.join(os.getcwd(), "temp"))
        self.temp_dir_entry = ttk.Entry(params_frame, textvariable=self.temp_dir_var)
        self.temp_dir_entry.grid(row=1, column=3, sticky="ew", padx=5, pady=2)
        
        ttk.Label(params_frame, text="修正閾值:").grid(row=2, column=0, sticky="w", padx=5, pady=2)
        self.correction_threshold_var = tk.StringVar(value="5")
        self.correction_threshold_entry = ttk.Entry(params_frame, textvariable=self.correction_threshold_var, width=15)
        self.correction_threshold_entry.grid(row=2, column=1, sticky="w", padx=5, pady=2)
        
        ttk.Label(params_frame, text="重疊容忍 (秒):").grid(row=2, column=2, sticky="w", padx=5, pady=2)
        self.overlap_tolerance_var = tk.StringVar(value="0.5")
        self.overlap_tolerance_entry = ttk.Entry(params_frame, textvariable=self.overlap_tolerance_var, width=15)
        self.overlap_tolerance_entry.grid(row=2, column=3, sticky="w", padx=5, pady=2)

        # --- 4. 執行與日誌區 ---
        action_frame = ttk.Frame(parent, padding="5")
        action_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.start_button = ttk.Button(action_frame, text="開始轉錄", command=self._start_transcription)
        self.start_button.pack(side=tk.LEFT, padx=(0, 10))
        
        # *** 新增: AI 摘要核取方塊 ***
        self.ai_summary_var = tk.BooleanVar(value=True)
        self.ai_summary_check = ttk.Checkbutton(action_frame, text="啟用 AI 智慧摘要", variable=self.ai_summary_var)
        self.ai_summary_check.pack(side=tk.LEFT, padx=10)
        
        self.status_var = tk.StringVar(value="狀態: 準備就緒")
        self.status_label = ttk.Label(action_frame, textvariable=self.status_var)
        self.status_label.pack(side=tk.LEFT)
        
        log_frame = ttk.LabelFrame(parent, text=" 即時日誌 ", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, state="disabled", font=('Consolas', 9))
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def _select_file(self):
        file_path = filedialog.askopenfilename(
            title="選擇影音檔案",
            filetypes=(("影音檔案", "*.mp4 *.mkv *.mp3 *.flac *.wav"), ("所有檔案", "*.*" ))
        )
        if file_path:
            self.file_path_var.set(file_path)
            self.status_var.set(f"狀態: 已選擇檔案 - {os.path.basename(file_path)}")

    def _check_dependencies(self):
        # *** 修改: 簡化依賴性檢查 ***
        if not os.path.exists(resource_path(TRANSCRIPTION_SCRIPT_NAME)):
            messagebox.showerror("依賴性錯誤", f"找不到核心腳本檔案:\n\n{TRANSCRIPTION_SCRIPT_NAME}\n\n請確保它和本程式放在同一個資料夾中。")
            self.master.destroy()
            return
        if sys.platform == "win32" and not os.path.exists(resource_path("ffmpeg.exe")):
            self._log_message("警告: 在程式目錄中未找到 'ffmpeg.exe'。請確保它與主程式在同一資料夾，或在系統 PATH 中。\n")

    def _log_message(self, message):
        self.log_text.config(state="normal")
        self.log_text.insert(tk.END, message)
        self.log_text.see(tk.END)
        self.log_text.config(state="disabled")

    def _build_prompt(self):
        base_template = self.main_rules_text.get("1.0", tk.END)
        language = self.language_var.get().strip() or "繁體中文"
        prompt = base_template.replace("{language}", language)
        terms_content = self.terms_text.get("1.0", tk.END).strip()
        if terms_content:
            terms_section = TERMS_TEMPLATE.format(terms_list=terms_content)
            prompt = prompt.replace("{terms_section}", terms_section)
        else:
            prompt = prompt.replace("{terms_section}", "")
        return prompt

    def _build_command(self):
        input_file = self.file_path_var.get()
        if not input_file:
            messagebox.showerror("錯誤", "請先選擇一個要處理的影音檔案。")
            return None
        try:
            prompt_content = self._build_prompt()
            # 使用 os.path.join 確保路徑正確
            temp_prompt_path = os.path.join(self.temp_dir_var.get(), f"temp_prompt_{datetime.now().strftime('%Y%m%d%H%M%S')}.md")
            
            # 確保暫存目錄存在
            os.makedirs(os.path.dirname(temp_prompt_path), exist_ok=True)

            with open(temp_prompt_path, 'w', encoding='utf-8') as f:
                f.write(prompt_content)
            self.temp_prompt_file = temp_prompt_path
            self._log_message(f"資訊: 已生成臨時 Prompt 檔案: {self.temp_prompt_file}\n")
        except Exception as e:
            messagebox.showerror("錯誤", f"無法建立臨時 Prompt 檔案於 '{self.temp_dir_var.get()}' 目錄: {e}")
            return None
        
        # *** 修改: 直接建立執行核心腳本的命令 ***
        command = [
            f'\"{sys.executable}\"',
            f'\"{resource_path(TRANSCRIPTION_SCRIPT_NAME)}\"',
            f'\"{input_file}\"',
            "--prompt_file", f'\"{self.temp_prompt_file}\"',
            "--chunk_duration", self.chunk_duration_var.get(),
            "--temp_dir", f'\"{self.temp_dir_var.get()}\"',
            "--model_name", self.model_name_var.get(),
            "--correction_threshold", self.correction_threshold_var.get(),
            "--overlap_tolerance", self.overlap_tolerance_var.get()
        ]
        api_key = self.api_key_var.get().strip()
        if api_key:
            command.extend(["--api_key", api_key])
        
        # *** 新增: 根據核取方塊決定是否加入 --enable_ai_summary ***
        if self.ai_summary_var.get():
            command.append("--enable_ai_summary")
            
        return command

    def _start_transcription(self):
        if self.is_running:
            messagebox.showwarning("提示", "目前已有一個轉錄任務正在執行中。")
            return
        command = self._build_command()
        if not command:
            return
        self.is_running = True
        self._set_ui_state(tk.DISABLED)
        self.log_text.config(state="normal")
        self.log_text.delete('1.0', tk.END)
        self.log_text.config(state="disabled")
        self.status_var.set("狀態: 轉錄中...請稍候...")
        self._log_message("="*80 + "\n")
        self._log_message(f"執行指令: {' '.join(command)}\n")
        self._log_message("="*80 + "\n\n")
        self.thread = threading.Thread(target=self._run_process, args=(command,), daemon=True)
        self.thread.start()
        self.master.after(100, self._process_log_queue)

    def _run_process(self, command):
        try:
            env = os.environ.copy()
            # 確保打包後的環境能找到 ffmpeg
            if hasattr(sys, '_MEIPASS'):
                env["PATH"] = sys._MEIPASS + os.pathsep + env.get('PATH', '')
            
            # 強制子程序使用 UTF-8 編碼，解決 Windows cp950 問題
            env["PYTHONUTF8"] = "1"
            
            command_str = " ".join(command)
            self.process = subprocess.Popen(
                command_str, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                text=True, encoding='utf-8', errors='replace', bufsize=1, shell=True,
                # 在 Windows 上隱藏子程序的控制台窗口
                creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0,
                env=env
            )
            for line in iter(self.process.stdout.readline, ''):
                self.log_queue.put(line)
            self.process.wait()
        except Exception as e:
            self.log_queue.put(f"\n\n[GUI 錯誤] 執行子程序時發生嚴重錯誤: {e}\n")
        finally:
            self.log_queue.put(None) # 發送結束信號

    def _process_log_queue(self):
        try:
            while True:
                line = self.log_queue.get_nowait()
                if line is None:
                    self._on_transcription_finish()
                    return
                self._log_message(line)
        except queue.Empty:
            pass
        if self.is_running:
            self.master.after(100, self._process_log_queue)

    def _on_transcription_finish(self):
        self.is_running = False
        return_code = self.process.returncode if self.process else -1
        if return_code == 0:
            self.status_var.set("狀態: 轉錄成功完成！")
            messagebox.showinfo("完成", "字幕轉錄與日誌流程已成功完成！")
        else:
            self.status_var.set(f"狀態: 轉錄失敗或已中止 (返回碼: {return_code})")
            messagebox.showerror("失敗", f"轉錄過程出錯或被中止.\n請檢查日誌以獲取詳細資訊。")
        self._set_ui_state(tk.NORMAL)
        # 刪除臨時檔案的邏輯保持不變
        if self.temp_prompt_file and os.path.exists(self.temp_prompt_file):
            try:
                os.remove(self.temp_prompt_file)
                self._log_message(f"\n資訊: 已刪除臨時 Prompt 檔案: {self.temp_prompt_file}\n")
            except Exception as e:
                self._log_message(f"\n警告: 無法刪除臨時 Prompt 檔案: {e}\n")
        self.temp_prompt_file = None

    def _set_ui_state(self, state):
        self.browse_button.config(state=state)
        self.start_button.config(state=state)
        self.ai_summary_check.config(state=state) # 控制核取方塊狀態
        for entry in [self.api_key_entry, self.model_name_entry, self.chunk_duration_entry, self.temp_dir_entry, self.language_entry, self.correction_threshold_entry, self.overlap_tolerance_entry]:
            entry.config(state='readonly' if state == tk.DISABLED else 'normal')
        for text_widget in [self.main_rules_text, self.terms_text]:
            text_widget.config(state=tk.DISABLED if state == tk.DISABLED else tk.NORMAL)

# ==============================================================================
#  主程式進入點
# ==============================================================================

def main_gui_launch():
    """ 啟動主 GUI 應用程式 """
    try:
        # 為了解決高 DPI 螢幕下的模糊問題
        if sys.platform == "win32":
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(1)
        
        root = tk.Tk()
        app = TranscriptionApp(root)
        
        # 將關閉按鈕綁定到自訂函式
        root.protocol("WM_DELETE_WINDOW", lambda: on_closing(root, app))
        
        root.mainloop()
    except Exception as e:
        import traceback
        error_message = f"GUI 啟動時發生致命錯誤:\n\n錯誤類型: {type(e).__name__}\n錯誤訊息: {str(e)}\n\n詳細追蹤資訊:\n{traceback.format_exc()}"
        # 備用錯誤顯示
        try:
            error_root = tk.Tk()
            error_root.withdraw()
            messagebox.showerror("致命錯誤", error_message)
        except:
            print(error_message, file=sys.stderr)

def on_closing(root, app):
    """ 處理關閉視窗事件 """
    if app.is_running:
        if messagebox.askokcancel("警告", "轉錄仍在進行中，確定要強制結束嗎？\n這可能會導致處理中斷和資料遺失。"):
            if app.process:
                app.process.terminate() # 嘗試終止子程序
            root.destroy()
    else:
        root.destroy()

if __name__ == "__main__":
    # 為了讓 PyInstaller 打包後的 exe 能正常運作
    multiprocessing.freeze_support()
    
    # 這段 runpy 邏輯是為了解決舊版打包時的特定問題，
    # 在簡化後的架構中，它不再是必要的，但保留它不會造成負面影響。
    # 為了穩定性，我們暫時保留它。
    if len(sys.argv) > 1 and sys.argv[1].endswith('.py'):
        script_to_run = sys.argv[1]
        sys.argv = sys.argv[1:] # 調整 sys.argv 以免混淆被呼叫的腳本
        runpy.run_path(resource_path(script_to_run), run_name="__main__")
    else:
        # 正常啟動 GUI
        main_gui_launch()
