
# run_and_log.py (v3.2 - The Final Encoding Fix)

import sys
import os
import io # <-- 導入 io 模組

# ==============================================================================
#  核心修正: 強制重設標準輸出的編碼為 UTF-8
# ==============================================================================
# 這是在 Windows 上打包成 .exe 後，解決 'cp950' UnicodeEncodeError 的終極方案。
# 它不再依賴環境變數，而是直接在程式碼層面重新配置 stdout 和 stderr。
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')
    except Exception as e:
        # 在極端情況下，如果重設失敗，至少讓程式能繼續，而不是崩潰
        print(f"[緊急警告] 重設 stdout/stderr 編碼失敗: {e}")

# ==============================================================================
#  核心修正: 打破 PyInstaller 的「部門牆」
# ==============================================================================
if hasattr(sys, '_MEIPASS'):
    sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# ==============================================================================
#  核心模組導入 (現在可以安全導入)
# ==============================================================================
try:
    import transcribe_pro_v5_banch_04_banch_09 as transcriber
except ImportError as e:
    print(f"[啟動器錯誤] 無法導入核心轉錄腳本: {e}")
    print("請確保 transcribe_pro_v5_banch_04_banch_09.py 與本程式在同一資料夾中，或已正確打包。")
    sys.exit(1)

try:
    from google import genai
except ImportError:
    genai = None
    print("[可選功能警告] 未找到 google-genai 套件，AI 智慧摘要功能將被停用。")

from datetime import datetime
import argparse

# ==============================================================================
#  後續程式碼與 v3.1 基本相同
# ==============================================================================

class Logger:
    def __init__(self, terminal, logfile):
        self.terminal = terminal
        self.log = logfile

    def write(self, message):
        self.terminal.write(message)
        self.log.write(message)
        self.log.flush()

    def flush(self):
        self.terminal.flush()
        self.log.flush()

def create_summary_prompt(log_content):
    return f"""你是一位專業的日誌分析助理。請分析以下這份從頭到尾的 Python 腳本執行日誌。

你的任務是：
1.  找出所有執行失敗或出現嚴重錯誤的音訊區塊 (例如 `船_chunk_000.mp3`)。
2.  對於每個有問題的區塊，條列出具體的錯誤訊息（例如 `ServerError`）或 SRT 糾錯記錄（例如 `時間軸重疊`、`無法解析時間戳`）。
3.  判斷對於出錯的區塊，後續的重試（`嘗試 2/3`）是否最終成功解決了問題（標誌是出現了 `成功！已將修正後的字幕儲存至...` 的訊息）。
4.  整理所有偵測到的伺服器錯誤 (例如 `HTTP/1.1 500 Internal Server Error`)。
5.  最後，用繁體中文給我一份清晰、分點的摘要報告，報告的標題是「日誌分析摘要報告」。

日誌內容如下：
---
{log_content}
---
"""

def run_ai_summary(log_filename, api_key, model_name):
    if not genai:
        print("AI 摘要功能已停用，因為缺少 google-genai 套件。")
        return

    print("\n" + "="*25 + " 自動 API 摘要階段 " + "="*25)
    try:
        if not os.path.exists(log_filename):
            print(f"錯誤：找不到日誌檔案 '{log_filename}' 進行摘要。")
            return

        final_api_key = api_key or os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
        client = genai.Client(api_key=final_api_key) if final_api_key else genai.Client()
        print("摘要器：成功建立 API 用戶端。")

        print(f"正在讀取日誌檔案: {log_filename}")
        with open(log_filename, 'r', encoding='utf-8') as f:
            log_content = f.read()

        if not log_content.strip():
            print("日誌檔案為空，無需摘要。")
            return

        prompt = create_summary_prompt(log_content)
        
        print(f"正在向模型 '{model_name}' 發送請求以生成摘要...")
        response = client.models.generate_content(model=model_name, contents=prompt)
        summary_content = response.text

        summary_filename = os.path.splitext(log_filename)[0] + "_summary.txt"
        with open(summary_filename, 'w', encoding='utf-8') as f:
            f.write(summary_content)
        
        print(f"摘要已成功生成並儲存至: {summary_filename}")

    except Exception as e:
        print("="*80)
        print(f"[嚴重警告] AI 智慧摘要過程中發生錯誤，但主轉錄流程不受影響。")
        print(f"錯誤類型: {type(e).__name__}")
        print(f"錯誤訊息: {e}")
        print("請檢查您的 API 金鑰是否有效、網路連線是否正常，或 Google 伺服器狀態。")
        print("="*80)

def extract_summary(log_lines, summary_filename):
    errors, corrections = [], []
    error_keywords = ['error', '錯誤', '失敗', 'exception', 'traceback', 'abnormal', '異常']
    correction_keywords = ["檢測到", "校正為", "修正時長", "縮短時長", "強制觸發安全回退", "時間軸重疊", "時間倒流", "無法解析時間戳"]
    for line in log_lines:
        line_lower = line.lower()
        if any(keyword in line_lower for keyword in error_keywords):
            errors.append(line.strip())
        elif any(keyword in line for keyword in correction_keywords):
            corrections.append(line.strip())
    with open(summary_filename, 'w', encoding='utf-8') as f:
        f.write("="*59 + "\n          手動關鍵字摘要 (錯誤與嚴重問題)\n" + "="*59 + "\n\n")
        f.write("\n".join(sorted(list(set(errors)), key=errors.index)) if errors else "太棒了！在執行過程中未偵測到任何錯誤 (ERROR) 或嚴重問題。\n")
        f.write("\n\n" + "="*59 + "\n             手動關鍵字摘要 (SRT 糾錯操作)\n" + "="*59 + "\n\n")
        f.write("\n".join(sorted(list(set(corrections)), key=corrections.index)) if corrections else "在執行過程中未偵測到任何 SRT 糾錯操作。\n")

def main():
    parser = argparse.ArgumentParser(description="日誌記錄與轉錄啟動器 (v3.2 - 整合版)")
    parser.add_argument("input_file", nargs='?', help="要處理的影音檔案路徑。")
    parser.add_argument("--api_key")
    parser.add_argument("--model_name", default="models/gemini-2.5-flash")
    args, unknown = parser.parse_known_args()

    input_file_path = next((arg for arg in sys.argv[1:] if not arg.startswith('--')), None)
    base_name = os.path.splitext(os.path.basename(input_file_path))[0] if input_file_path else "script_run"
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    log_filename = f"{base_name}_日誌_{timestamp}.txt"
    summary_filename = f"{base_name}_摘要_{timestamp}.txt"

    original_stdout, original_stderr = sys.stdout, sys.stderr
    return_code = 0

    try:
        with open(log_filename, 'w', encoding='utf-8') as log_file:
            sys.stdout = sys.stderr = Logger(original_stdout, log_file)
            print("="*80 + f"\n日誌記錄器已啟動 (v3.2)。執行的核心邏輯來自: {transcriber.__name__}.py\n" + f"完整日誌將記錄到: {log_filename}\n" + "="*80)
            
            transcriber.main()

    except SystemExit as e:
        return_code = e.code if isinstance(e.code, int) else 1
    except Exception as e:
        return_code = 1
        print(f"\n[啟動器致命錯誤] 執行過程中發生未預期的錯誤: {e}")
        import traceback
        print(traceback.format_exc())
    finally:
        sys.stdout, sys.stderr = original_stdout, original_stderr

    print("="*80 + f"\n主程式邏輯執行完畢。返回碼: {return_code}\n" + "="*80)

    if os.path.exists(log_filename):
        with open(log_filename, 'r', encoding='utf-8') as f:
            log_lines = f.readlines()
        print(f"正在生成手動摘要檔案: {summary_filename}")
        extract_summary(log_lines, summary_filename)
        print("手動摘要檔案已成功生成。")

    run_ai_summary(log_filename, args.api_key, args.model_name)
    
    print("="*80 + "\n所有流程已執行完畢。" + "\n" + "="*80)

if __name__ == "__main__":
    main()
