# transcribe_pro_v5_branch_04_branch_12.py
# 版本號: v5_4_12_20250803
# 修改內容簡述:
# 1. 【終極編碼修正】在 main() 函式和主執行區塊中，都加入了強制設定 sys.stdout 和 sys.stderr
#    編碼為 UTF-8 的邏輯。
# 2. 此修正旨在徹底解決在 Windows 系統上將程式打包成 .exe 後，因系統預設編碼 (cp950)
#    與日誌中的寬字元 (如中文) 衝突而導致的 UnicodeEncodeError。
# 3. 確保腳本無論是透過 GUI 間接呼叫，還是直接在命令列執行，都能穩定地處理各種字元。

import os
import sys
import subprocess
import re
from datetime import datetime, timedelta
import argparse
import logging
import time
import io

# ==============================================================================
# 核心編碼修正函式
# ==============================================================================
def force_utf8_encoding():
    """
    強制將 sys.stdout 和 sys.stderr 的編碼設定為 UTF-8。
    這是解決 Windows cp950 編碼錯誤的關鍵。
    """
    if sys.stdout.encoding != 'utf-8':
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    if sys.stderr.encoding != 'utf-8':
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

# ==============================================================================
# 全域日誌設定
# ==============================================================================
log_string_io = io.StringIO()

class Tee(object):
    def __init__(self, *files):
        self.files = files
    def write(self, obj):
        for f in self.files:
            f.write(obj)
            f.flush()
    def flush(self):
        for f in self.files:
            f.flush()

def setup_logging(log_filename):
    global log_string_io
    log_string_io = io.StringIO()

    for handler in logging.root.handlers[:]:
        logging.root.removeHandler(handler)

    log_format = '%(asctime)s - %(levelname)s - %(message)s'
    
    # Tee 會將日誌同時發送到原始的 stdout 和我們的記憶體 StringIO
    tee_stdout = Tee(sys.stdout, log_string_io)
    
    logging.basicConfig(
        level=logging.INFO,
        format=log_format,
        stream=tee_stdout,
        force=True
    )
    
    # 注意：這裡不再重定向 sys.stdout/stderr，因為我們在 main() 的開頭已經處理了編碼
    print(f"日誌將被記錄到: {log_filename}")

def save_log_file(log_filename):
    with open(log_filename, 'w', encoding='utf-8') as f:
        f.write(log_string_io.getvalue())
    print(f"日誌檔案已儲存: {log_filename}")

# --- Gemini API 相關 ---
try:
    from google import genai
except ImportError:
    print("="*80)
    print("【重要環境配置錯誤】")
    print("程式無法執行 'from google import genai'。")
    print("請檢查您是否已安裝了支援此導入方式的正確 Google GenAI SDK 套件，例如：")
    print(">>> pip install -U -q google-genai")
    print("="*80)
    exit(1)

# ==============================================================================
# SRT 糾錯邏輯 (v14)
# ==============================================================================

def parse_time_v9(time_str):
    ts = time_str.strip()
    if ":_" in ts:
        ts = ts.replace(':_', ':')
        logging.warning(f"檢測到並修正了異常的時間戳格式 'HH:_MM' -> '{ts}'")

    if re.match(r'^\d+:\d+:\d+:\d+[,.]\d+$', ts):
        logging.warning(f"檢測到含影格(frame)的四段式時間戳，此格式無法安全轉換，已跳過: {ts}"); return None
    ts = re.sub(r'(\d+:\d+:\d+)[.:](\d+)$', r'\1,\2', ts)
    ts = re.sub(r'^(\d+:\d+)[.,:](\d+)$', r'00:\1,\2', ts)
    match = re.match(r'^(\d+):(\d+):(\d+),(\d+)$', ts)
    if not match:
        logging.warning(f"無法解析時間戳格式: {ts}"); return None
    try:
        h, m, s, ms = (int(g) for g in match.groups())
        if s >= 60: logging.warning(f"檢測到無效的秒數 ({s})，將其校正為 59。原始: '{time_str}'"); s = 59
        if m >= 60: logging.warning(f"檢測到無效的分鐘數 ({m})，將其校正為 59。原始: '{time_str}'"); m = 59
        return timedelta(hours=h, minutes=m, seconds=s, milliseconds=ms)
    except ValueError:
        logging.error(f"時間戳中的數字無法轉換: {ts}"); return None

def format_timedelta_v7(td):
    if not isinstance(td, timedelta): return "00:00:00,000"
    total_seconds = td.total_seconds()
    if total_seconds < 0: total_seconds = 0
    hours, remainder = divmod(total_seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    milliseconds = td.microseconds // 1000
    return f"{int(hours):02}:{int(minutes):02}:{int(seconds):02},{milliseconds:03}"

def format_srt_from_text_v14(srt_content, audio_filename, overlap_tolerance_td):
    if srt_content.startswith("```srt"): srt_content = srt_content[srt_content.find('\n') + 1:]
    if srt_content.endswith("\n```"): srt_content = srt_content[:-4]
    srt_content = srt_content.strip().replace('\r\n', '\n')

    pattern = re.compile(r'(\n*\d+\n[\d:,._]+\s*-->\s*[\d:,._]+)')
    processed_content = pattern.sub(r'<SRT_BLOCK_SEPARATOR>\1', srt_content)
    if processed_content.startswith('<SRT_BLOCK_SEPARATOR>'):
        processed_content = processed_content[len('<SRT_BLOCK_SEPARATOR>'):]
    blocks = [block.strip() for block in processed_content.split('<SRT_BLOCK_SEPARATOR>') if block.strip()]

    corrected_blocks = []
    entry_counter = 1
    last_correct_end_td = timedelta(0)
    severe_correction_count = 0
    SAFE_FALLBACK_DURATION = timedelta(seconds=5)
    MAX_DURATION = timedelta(minutes=3)

    for i, block in enumerate(blocks):
        lines = [line.strip() for line in block.split('\n') if line.strip()]
        if len(lines) < 2: continue
        time_line = lines[1]
        text_lines = lines[2:]
        full_text = '\n'.join(text_lines)
        if not full_text: continue

        try:
            start_raw, end_raw = [t.strip() for t in time_line.split('-->')]
            start_td, end_td = parse_time_v9(start_raw), parse_time_v9(end_raw)
            
            is_unparsable = start_td is None or end_td is None
            is_overlap_violation = not is_unparsable and start_td < (last_correct_end_td - overlap_tolerance_td)

            if is_unparsable or is_overlap_violation:
                severe_correction_count += 1
                if is_unparsable:
                    logging.warning(f"塊 {i+1} 無法解析時間戳: '{time_line}'。強制觸發安全回退。 (嚴重修正計數: {severe_correction_count})")
                else:
                    logging.warning(f"塊 {i+1} 檢測到嚴重時間軸重疊 (> {overlap_tolerance_td.total_seconds()}s): 上一條結束於 {format_timedelta_v7(last_correct_end_td)}, 此條卻開始於 {format_timedelta_v7(start_td)}。強制觸發安全回退。 (嚴重修正計數: {severe_correction_count})")
                start_td = last_correct_end_td + timedelta(milliseconds=100)
                end_td = start_td + SAFE_FALLBACK_DURATION
            
            if end_td <= start_td:
                logging.warning(f"塊 {i+1} 檢測到時間倒流或零時長。修正時長。")
                end_td = start_td + SAFE_FALLBACK_DURATION
            
            if (end_td - start_td) > MAX_DURATION:
                logging.warning(f"塊 {i+1} 檢測到超長持續時間。縮短時長。")
                end_td = start_td + SAFE_FALLBACK_DURATION
            
            last_correct_end_td = end_td
            corrected_blocks.append(f"{entry_counter}\n{format_timedelta_v7(start_td)} --> {format_timedelta_v7(end_td)}\n{full_text}")
            entry_counter += 1
        except Exception as e:
            logging.error(f"處理字幕塊 {i+1} 時發生嚴重錯誤，已跳過: '{block[:50]}...' - 錯誤: {e}")

    if not corrected_blocks:
        logging.warning(f"在 {audio_filename} 的回應中最終未能解析出任何有效的字幕塊。")
        return "", severe_correction_count

    return "\n\n".join(corrected_blocks) + "\n\n", severe_correction_count

# ==============================================================================
# 主要工作流程函式
# ==============================================================================

def transcribe_audio(client, audio_path, prompt_text, model_name, correction_threshold, overlap_tolerance, max_retries=3):
    srt_path = os.path.splitext(audio_path)[0] + ".srt"
    file_basename = os.path.basename(audio_path)
    tokens_used = 0
    retry_delays = [65, 130, 250]
    overlap_tolerance_td = timedelta(seconds=overlap_tolerance)

    for attempt in range(max_retries):
        uploaded_file = None
        try:
            logging.info(f"[{file_basename} | 嘗試 {attempt+1}/{max_retries}] 正在上傳檔案...")
            uploaded_file = client.files.upload(file=audio_path)
            logging.info(f"檔案已上傳。正在向模型 '{model_name}' 發送轉錄請求...")
            response = client.models.generate_content(model=model_name, contents=[prompt_text, uploaded_file])
            if hasattr(response, 'usage_metadata') and response.usage_metadata:
                tokens_used = response.usage_metadata.total_token_count
            if uploaded_file:
                logging.info(f"請求完成，正在刪除已上傳的遠端檔案 '{uploaded_file.name}'...")
                client.files.delete(name=uploaded_file.name)
            
            if response.text:
                raw_srt_path = os.path.splitext(srt_path)[0] + ".raw.txt"
                with open(raw_srt_path, 'w', encoding='utf-8') as f: f.write(response.text)
                
                corrected_srt, severe_correction_count = format_srt_from_text_v14(response.text, file_basename, overlap_tolerance_td)
                
                if severe_correction_count > correction_threshold:
                    logging.warning(f"[{file_basename}] 偵測到 {severe_correction_count} 次嚴重修正，已超過閾值 {correction_threshold}。")
                    logging.warning("目前的 SRT 結果品質不佳，將觸發自動重試以獲取更好的結果。")
                    raise ValueError(f"Correction threshold exceeded: {severe_correction_count} > {correction_threshold}")
                
                with open(srt_path, 'w', encoding='utf-8') as f: f.write(corrected_srt)
                logging.info(f"成功！已將修正後的字幕儲存至: {os.path.basename(srt_path)}")
                return srt_path, tokens_used
            else:
                logging.error(f"無法從 '{file_basename}' 的回應中獲取文字。")
                return None, tokens_used

        except Exception as e:
            logging.warning(f"處理 '{file_basename}' 時捕獲到異常 (類型: {type(e).__name__})。訊息: {e}")
            if uploaded_file:
                try: client.files.delete(name=uploaded_file.name)
                except Exception as del_e: logging.warning(f"在錯誤處理中刪除遠端檔案失敗: {del_e}")
            if attempt < max_retries -1:
                wait_time = retry_delays[attempt]
                logging.info(f"將在 {wait_time} 秒後重試...")
                time.sleep(wait_time)
            else:
                logging.error(f"已達到最大重試次數 ({max_retries})，轉錄 '{file_basename}' 失敗。")
                return None, tokens_used
    return None, tokens_used

def split_audio(input_file, temp_dir, chunk_duration_seconds=600):
    if not os.path.exists(temp_dir): os.makedirs(temp_dir)
    base_name = os.path.splitext(os.path.basename(input_file))[0]
    output_pattern = os.path.join(temp_dir, f"{base_name}_chunk_%03d.mp3")
    existing_chunks = sorted([os.path.join(temp_dir, f) for f in os.listdir(temp_dir) if f.startswith(f"{base_name}_chunk_") and f.endswith(".mp3")])
    if existing_chunks:
        logging.info(f"偵測到 {len(existing_chunks)} 個已存在的音訊區塊。跳過 FFmpeg 分割。")
        return existing_chunks
    
    command = ['ffmpeg', '-i', input_file, '-vn', '-f', 'segment', '-segment_time', str(chunk_duration_seconds), '-acodec', 'libmp3lame', '-b:a', '192k', output_pattern]
    try:
        logging.info(f"開始使用 FFmpeg 分割: {os.path.basename(input_file)}")
        startupinfo = None
        if os.name == 'nt':
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE

        process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, startupinfo=startupinfo, creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0)
        stdout, stderr = process.communicate()

        if process.returncode != 0:
            logging.error(f"FFmpeg 錯誤: {stderr.decode('utf-8', errors='ignore')}")
            return []
            
    except FileNotFoundError: 
        logging.error("找不到 FFmpeg。請確保它已安裝並在系統 PATH 中，或者 ffmpeg.exe 與本程式位於同一目錄。"); return []
    except Exception as e:
        logging.error(f"執行 FFmpeg 時發生未知錯誤: {e}"); return []
        
    return sorted([os.path.join(temp_dir, f) for f in os.listdir(temp_dir) if f.startswith(f"{base_name}_chunk_") and f.endswith(".mp3")])


def merge_srts(srt_files, final_srt_path, chunk_duration_seconds):
    logging.info(f"正在將 {len(srt_files)} 個 SRT 檔案合併至 {os.path.basename(final_srt_path)}")
    global_offset, entry_counter = timedelta(0), 1
    chunk_duration_td = timedelta(seconds=chunk_duration_seconds)
    with open(final_srt_path, 'w', encoding='utf-8') as outfile:
        sorted_srts = sorted(srt_files)
        for i, srt_file in enumerate(sorted_srts):
            try:
                with open(srt_file, 'r', encoding='utf-8') as infile:
                    content = infile.read().strip()
                    if not content:
                        if i < len(sorted_srts) - 1: global_offset += chunk_duration_td
                        continue
                    for entry in content.split('\n\n'):
                        if not entry.strip(): continue
                        lines = entry.split('\n')
                        if len(lines) < 2 or '-->' not in lines[1]: continue
                        start_str, end_str = [t.strip() for t in lines[1].split('-->')]
                        start_td, end_td = parse_time_v9(start_str), parse_time_v9(end_str)
                        if start_td is None or end_td is None: continue
                        adjusted_start, adjusted_end = start_td + global_offset, end_td + global_offset
                        outfile.write(f"{entry_counter}\n{format_timedelta_v7(adjusted_start)} --> {format_timedelta_v7(adjusted_end)}\n" + '\n'.join(lines[2:]) + '\n\n')
                        entry_counter += 1
                if i < len(sorted_srts) - 1: global_offset += chunk_duration_td
            except FileNotFoundError:
                logging.warning(f"找不到要合併的 SRT 檔案: {srt_file}。時間偏移量將基於區塊長度繼續計算。")
                if i < len(sorted_srts) - 1: global_offset += chunk_duration_td
            except Exception as e:
                logging.error(f"合併 SRT '{os.path.basename(srt_file)}' 時發生錯誤: {e}")
                if i < len(sorted_srts) - 1: global_offset += chunk_duration_td

# ==============================================================================
# 日誌摘要功能
# ==============================================================================

def create_manual_summary(log_content, summary_filename):
    keywords = [
        "錯誤", "失敗", "警告", "Error", "Failed", "Warning", "Exception",
        "重試", "Retry", "閾值", "Threshold", "修正", "Correction",
        "無法解析", "Unparsable", "時間軸重疊", "Overlap", "安全回退",
        "總共使用了", "token", "最終 SRT"
    ]
    summary_lines = [line for line in log_content.splitlines() if any(keyword in line for keyword in keywords)]
    summary_text = "\n".join(summary_lines)
    with open(summary_filename, 'w', encoding='utf-8') as f:
        f.write("="*40 + "\n手動關鍵字摘要\n" + "="*40 + "\n\n" + summary_text)
    logging.info(f"已生成手動摘要檔案: {summary_filename}")

def create_ai_summary(log_content, summary_filename, client, model_name):
    logging.info("偵測到 AI 摘要請求。正在準備呼叫 API...")
    if not client:
        logging.error("AI 摘要失敗：找不到 API 用戶端。")
        return

    prompt = f"""
    你是一位專業的軟體測試工程師，你的任務是分析以下這份由 Python 字幕轉錄腳本產生的日誌檔案。
    請從中提取關鍵資訊，並以清晰、有條理的方式總結出來。

    你的分析應包含以下幾點：
    1.  **整體狀況總結**：簡要說明這次轉錄任務是成功還是失敗，以及是否發生了任何嚴重問題。
    2.  **音訊區塊處理**：列出每個音訊區塊 (chunk) 的處理情況。對於每個區塊，請註明：
        *   是否轉錄成功。
        *   是否觸發了重試機制？如果是，重試了幾次？最終是否成功？
        *   是否發生了大量的「嚴重修正」？（例如，時間軸重疊、無法解析時間戳等）
    3.  **主要錯誤與警告**：如果日誌中出現了任何 Python 錯誤 (Exception)、FFmpeg 錯誤或 API 錯誤，請明確指出它們的內容和發生的時間點。
    4.  **資源使用**：記錄總共使用了多少 token。
    5.  **最終結果**：說明最終的 SRT 檔案是否成功生成。

    請使用繁體中文、Markdown 格式來呈現你的報告，使其易於閱讀。

    --- 以下是日誌內容 ---
    {log_content}
    --- 日誌內容結束 ---
    """
    
    try:
        logging.info(f"正在向模型 '{model_name}' 發送 AI 摘要請求...")
        response = client.models.generate_content(model=model_name, contents=[prompt])
        ai_summary = response.text
        with open(summary_filename, 'a', encoding='utf-8') as f:
            f.write("\n\n" + "="*40 + "\nAI 智慧摘要\n" + "="*40 + "\n\n" + ai_summary)
        logging.info(f"已將 AI 智慧摘要附加到檔案: {summary_filename}")
    except Exception as e:
        logging.error(f"生成 AI 智慧摘要時發生錯誤: {e}")
        error_message = f"\n\n--- AI 摘要失敗 ---\n錯誤詳情: {e}\n"
        with open(summary_filename, 'a', encoding='utf-8') as f:
            f.write(error_message)

# ==============================================================================
# 主函式
# ==============================================================================

def main():
    # 【關鍵修正】在程式邏輯開始前，立即強制設定編碼
    force_utf8_encoding()
    
    start_time = time.time()
    parser = argparse.ArgumentParser(description="自動化音訊/視訊轉錄工作流程 (v12 - 強制UTF-8編碼)。")
    parser.add_argument("input_file", help="要處理的大型音訊或視訊檔案路徑。")
    parser.add_argument("--temp_dir", default=os.path.join(os.getcwd(), "temp"), help="儲存臨時音訊區塊和 SRT 檔案的目錄。")
    parser.add_argument("--chunk_duration", type=int, default=600, help="每個音訊區塊的持續時間（秒）。")
    parser.add_argument("--prompt_file", default="rule.md", help="提示檔案的路徑。")
    parser.add_argument("--api_key", help="您的 API 金鑰 (手動指定，優先級最高)。")
    parser.add_argument("--model_name", default="models/gemini-2.5-pro", help="要使用的 Gemini 模型名稱。")
    parser.add_argument("--correction_threshold", type=int, default=5, help="單一區塊內，觸發自動重跑的嚴重修正次數閾值。")
    parser.add_argument("--overlap_tolerance", type=float, default=0.5, help="允許的字幕時間軸重疊容忍秒數。")
    parser.add_argument("--enable_ai_summary", action='store_true', help="處理完成後，使用 API 對日誌進行智慧摘要。")
    
    args = parser.parse_args()

    file_basename = os.path.splitext(os.path.basename(args.input_file))[0]
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_filename = f"{file_basename}_日誌_{timestamp}.txt"
    summary_filename = f"{file_basename}_摘要_{timestamp}.txt"
    
    # 在設定日誌系統前，編碼必須已被修正
    setup_logging(log_filename)

    logging.info("="*50)
    logging.info(f"開始處理檔案: {args.input_file}")
    logging.info(f"版本: transcribe_pro_v5_branch_04_branch_12")
    logging.info("="*50)

    api_key = args.api_key or os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
    client = None
    try:
        client = genai.Client(api_key=api_key) if api_key else genai.Client()
        logging.info(f"成功建立 API 用戶端。將使用模型: {args.model_name}")
    except Exception as e:
        logging.error(f"建立 API 用戶端失敗: {e}");
        if not api_key: logging.error("請確認 GEMINI_API_KEY 或 GOOGLE_API_KEY 環境變數已正確設定，或使用 --api_key 參數。")
        save_log_file(log_filename)
        return
    
    prompt_text = ""
    if os.path.exists(args.prompt_file):
        with open(args.prompt_file, 'r', encoding='utf-8') as f: prompt_text = f.read()
    else:
        logging.warning(f"找不到提示檔案: {args.prompt_file}。將使用空提示。")

    if not os.path.exists(args.temp_dir): os.makedirs(args.temp_dir)
    chunk_mp3_files = split_audio(args.input_file, args.temp_dir, args.chunk_duration)
    
    if chunk_mp3_files:
        all_chunk_srts, total_tokens_used = [], 0
        for chunk_mp3_path in chunk_mp3_files:
            chunk_srt_path = os.path.splitext(chunk_mp3_path)[0] + ".srt"
            all_chunk_srts.append(chunk_srt_path)
            if os.path.exists(chunk_srt_path) and os.path.getsize(chunk_srt_path) > 0:
                logging.info(f"偵測到已存在的 SRT，跳過轉錄: {os.path.basename(chunk_srt_path)}")
            else:
                _, tokens = transcribe_audio(client, chunk_mp3_path, prompt_text, args.model_name, args.correction_threshold, args.overlap_tolerance)
                total_tokens_used += tokens if tokens else 0

        logging.info(f"所有音訊區塊處理完成。總共使用了約 {total_tokens_used} 個 token。")
        final_srt_path = f"{file_basename}.srt"
        merge_srts(all_chunk_srts, final_srt_path, args.chunk_duration)
        logging.info(f"工作流程完成。最終 SRT 檔案位於: {final_srt_path}")

    end_time = time.time()
    logging.info(f"總共耗時: {end_time - start_time:.2f} 秒。")
    
    log_content = log_string_io.getvalue()
    save_log_file(log_filename)
    create_manual_summary(log_content, summary_filename)
    
    if args.enable_ai_summary:
        create_ai_summary(log_content, summary_filename, client, args.model_name)

if __name__ == "__main__":
    # 【關鍵修正】在主執行區塊也強制設定編碼，確保萬無一失
    force_utf8_encoding()
    try:
        main()
    except Exception as e:
        # 現在 logging.error 應該可以安全地處理任何字元了
        logging.error(f"主程式發生未預期的嚴重錯誤: {e}", exc_info=True)
    finally:
        print("\n程式執行完畢。")
