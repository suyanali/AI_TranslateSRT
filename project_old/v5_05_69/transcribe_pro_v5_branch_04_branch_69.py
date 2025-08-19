# transcribe_pro_v5_branch_04_branch_69.py
# 版本號: v5_6_69_20250813
# 修改內容簡述:
# 1. 【SRT校正強化】: 根據使用者回饋的11點校正規則，對 `parse_time_v10` 函式進行全面強化。
# 2. 【規則 6 修正】: 新增正規表示式邏輯，能夠正確處理以冒號分隔毫秒的非標準時間格式 (例如 '00:00:190' -> '00:00,190')。
# 3. 【規則 1 防呆】: 新增對過多冒號 (例如 'hh:mm:ss:ff,xxx') 的無效格式檢查，提前攔截並記錄錯誤。
# 4. 【程式碼審查】: 再次確認現有程式碼已涵蓋其他校正規則，如小數點轉換、時間單位補零、秒/分鐘數超限校正、時間倒流/重疊等問題。
# 5. 【功能新增】: 新增 `run_partial_transcription_task` 函式，實現對音訊指定時間區段進行局部轉錄的功能。
# 6. 【核心邏輯】: `run_partial_transcription_task` 會：
#    a. 使用 FFmpeg 根據使用者提供的開始與結束時間，精準切割一小段音訊。
#    b. 呼叫現有的 `transcribe_audio` 模組進行轉錄。
#    c. 【時間軸校正】: 自動將轉錄結果 (時間軸從 0 開始) 加上開始時間的偏移量，產生與原始音訊時間軸對齊的 SRT 內容。
#    d. 將校正後的 SRT 儲存為獨立檔案 (檔名包含時間戳)，不覆蓋任何現有檔案。
#    e. 提供選項，允許使用者決定是否保留切割出的暫存音訊檔。
# 7. 【CLI 擴充】: `main_cli` 新增 `--partial_only`, `--start_time`, `--end_time`, `--keep_partial_audio` 等命令列參數以支援新功能。
# 8. 【功能變更】腳本的重做機制會更加嚴格，當偵測到字幕持續時間異常過長時，也會將其視為嚴重錯誤並觸發重試，默認修正閾值(「嚴重修正」)改為6。嚴重修正標準：無法解析時間戳或時間倒流、時間軸重疊、字幕持續時間異常過長(預設18秒[MAX_DURATION = timedelta(seconds=18)])
# 9. 【功能變更】生成 SRT轉錄情況報告將「API 回應為空值 (empty response) 」視為一個正式的錯誤進而觸發重試流程
# 10.【新增】併發處理：引入 ThreadPoolExecutor，透過 --workers 參數可設定多執行緒同時處理音訊分塊，加快處理速度。
# 11.【新增】智慧速率限制：實作 MinuteRateLimiter，透過 --rpm 參數控制每分鐘 API 請求總數，降低 429 錯誤機率。GUI介面可填整
# 12.【強化重試機制】將原固定延遲 [65, 130, 250] 改為「指數退避 + 全抖動」(sleep_with_full_jitter)，在錯誤重試時更智慧地分散等待時間，避免所有執行緒同時重試。
# 13.【新增】命令列參數：加入 --workers、--rpm、--max_retries、--retry_base、--retry_cap 等進階設定，方便調整效能
# 14.【新增】連續空回應中止：新增 --empty_abort_threshold 參數，當連續收到 API 空回應達到設定次數時，自動中止整個轉錄流程。
# 15.【新增】SRTContentParseError 異常類別，用於表示 SRT 內容解析失敗的情況。
# 16.【重構】`format_srt_from_text_v16` 函式中的 SRT 內容解析邏輯，使其採用逐行解析的方式，以提高對不規則 SRT 格式的容錯性。現在，即使文本為空，只要時間戳有效，也會將其視為一個塊。當無法解析出任何有效字幕塊時，會拋出 `SRTContentParseError`。調整 `transcribe_audio` 函式的錯誤處理，使其能夠捕獲 `EmptyResponseError` 和 `SRTContentParseError`，並觸發內部的指數退避重試邏輯。
# 17.【重構】因應多works導致的日誌不清，更改最後的SRT轉錄日誌(ai摘要)prompt以及在所有警告訊息前，加入了 `audio_filename` 變數，以明確指出是哪個音訊檔案的 SRT 塊出現問題。在 `run_transcription_task` 函式內的 `_job` 函式中，當任務產生例外時，在錯誤訊息中加入了 `os.path.basename(path)`，以顯示是哪個音訊區塊檔案導致了例外。
# 18.【重構】parse_time_v10 函式中的毫秒補零規則修改為在前面補處理已更新為「補在百位數」的需求（例如 99 變成 099）。
import os
import sys
import subprocess
import re
from datetime import datetime, timedelta
import argparse
import logging
import time
import io
import math
from types import SimpleNamespace

# NEW: 併發與限速所需 import
import threading
import random
from collections import deque
from concurrent.futures import ThreadPoolExecutor, as_completed
from threading import Lock

# NEW: 自訂例外（共用）
class EmptyResponseError(Exception):
    pass

# NEW: 自訂例外（SRT內容解析失敗）
class SRTContentParseError(Exception):
    pass

# NEW: 單程序共用的滑動視窗 RPM 限速器
class MinuteRateLimiter:
    """單程序滑動視窗 RPM 限速器：所有執行緒共用。"""
    def __init__(self, rpm: int):
        self.rpm = max(1, int(rpm))
        self.ts = deque()
        self.lock = threading.Lock()

    def wait(self):
        window = 60.0
        with self.lock:
            now = time.time()
            # 清掉 60 秒之前的請求時間戳
            while self.ts and now - self.ts[0] > window:
                self.ts.popleft()
            if len(self.ts) >= self.rpm:
                sleep_for = window - (now - self.ts[0]) + 0.01
            else:
                sleep_for = 0.0
        if sleep_for > 0:
            time.sleep(sleep_for)
        with self.lock:
            self.ts.append(time.time())

# NEW: 指數退避 + 全抖動
def sleep_with_full_jitter(base=65, attempt=1, cap=250):
    """
    Exponential Backoff with Full Jitter
    base: 第一次重試的最大等待上限（秒）
    attempt: 第 N 次重試（從 1 開始）
    cap: 單次延遲最大上限（秒）
    """
    t = min(cap, base * (2 ** (attempt - 1)))
    time.sleep(random.uniform(0, t))

def get_application_path():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

APP_PATH = get_application_path()

def force_utf8_encoding():
    if sys.stdout.encoding != 'utf-8':
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    if sys.stderr.encoding != 'utf-8':
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

class QueueHandler(logging.Handler):
    def __init__(self, log_queue):
        super().__init__()
        self.log_queue = log_queue
    def emit(self, record):
        self.log_queue.put(self.format(record))

def setup_logging(log_filename, verbose=False, log_queue=None):
    log_format = '%(asctime)s - %(levelname)s - %(message)s'
    level = logging.DEBUG if verbose else logging.INFO
    for handler in logging.root.handlers[:]:
        logging.root.removeHandler(handler)
    handlers = [logging.FileHandler(log_filename, 'a', 'utf-8')]
    if log_queue:
        handlers.append(QueueHandler(log_queue))
    else:
        handlers.append(logging.StreamHandler(sys.stdout))
    logging.basicConfig(level=level, format=log_format, handlers=handlers)
    
    logging.getLogger("google.generativeai").setLevel(logging.INFO)
    logging.getLogger("urllib3").setLevel(logging.WARNING)
    logging.getLogger("google.api_core").setLevel(logging.WARNING)

try:
    from google import genai
except ImportError:
    pass

def is_final_srt_valid(srt_text):
    """
    檢查最終生成的 SRT 字串結構是否完整。
    判斷標準：序列號的數量必須嚴格等於時間戳的數量。
    """
    if not srt_text or not srt_text.strip():
        return True # 空的 SRT 內容視為有效，不觸發重試

    # 計算時間戳 (-->) 的數量
    arrow_count = srt_text.count('-->')

    # 使用正則表達式計算序列號行的數量 (一行中只有數字)
    # re.MULTILINE 讓 ^ 和 $ 能匹配每一行的開頭和結尾
    sequence_number_count = len(re.findall(r'^\d+$', srt_text, re.MULTILINE))

    if arrow_count != sequence_number_count:
        logging.warning(f"最終 SRT 結構驗證失敗：時間戳數量 ({arrow_count}) 與序列號數量 ({sequence_number_count}) 不匹配。")
        return False
        
    return True

def parse_time_v10(time_str):
    """
    將多種格式的時間字串 (HH:MM:SS,ms 或 HH:MM:SS.ms) 解析為 timedelta 物件。
    這個版本增強了對毫秒的處理，並能處理不規則的格式。
    """
    ts = time_str.strip().replace(':_', ':').replace('：', ':')
    
    # 處理潛在的格林威治時間格式 (HH:MM:SS.msZ)
    if ts.upper().endswith('Z'):
        ts = ts[:-1]

    # 【規則 6 修正】: 處理 hh:mm:ss:ms 或 mm:ss:ms 這種以冒號分隔毫秒的格式
    # 優先修正，避免被後續的防呆機制誤擋
    ts = re.sub(r'(.*):(\d{1,3})$', r'\1,\2', ts)

    # 偵測並標準化毫秒分隔符
    if ',' in ts:
        parts = ts.split(',')
        time_part = parts[0]
        ms_part = parts[1]
    elif '.' in ts:
        parts = ts.split('.')
        # 避免錯誤地將 HH.MM.SS 中的點視為毫秒分隔符
        if len(parts) > 2 and ':' in parts[-2]:
             time_part = ".".join(parts[:-1])
             ms_part = parts[-1]
        else:
             time_part = parts[0]
             ms_part = parts[1] if len(parts) > 1 else "0"
    else:
        time_part = ts
        ms_part = "0"

    # 【規則 1 防呆】: 檢查是否存在過多的冒號，這是一個無效的格式
    if time_part.count(':') > 2:
        logging.warning(f"無法解析時間戳格式: '{time_str}' (原因: 時間部分 '{time_part}' 包含超過2個冒號)。")
        return None

    # 補全時間部分
    time_parts = time_part.split(':')
    if len(time_parts) == 1: # SS
        time_part = f"00:00:{time_parts[0]}"
    elif len(time_parts) == 2: # MM:SS
        time_part = f"00:{time_parts[0]}:{time_parts[1]}"

    # 組合回標準格式
    standard_ts = f"{time_part},{ms_part}"

    match = re.match(r'^(\d+):(\d+):(\d+),(\d+)$', standard_ts)
    if not match:
        logging.warning(f"無法解析時間戳格式: '{time_str}' (標準化後為: '{standard_ts}')")
        return None
    
    try:
        h, m, s, ms = (int(g) for g in match.groups())
        
        # 對毫秒進行位數補齊 (例如 5 -> 500, 50 -> 500)
        ms_str = match.group(4)
        if len(ms_str) < 3:
            ms = int(ms_str.rjust(3, '0'))

        if s >= 60:
            logging.warning(f"秒數無效 ({s})，校正為 59。原始: '{time_str}'")
            s = 59
        if m >= 60:
            logging.warning(f"分鐘數無效 ({m})，校正為 59。原始: '{time_str}'")
            m = 59
        return timedelta(hours=h, minutes=m, seconds=s, milliseconds=ms)
    except (ValueError, IndexError):
        logging.error(f"時間戳中的數字無法轉換: '{time_str}'")
        return None

def format_timedelta_v7(td):
    if not isinstance(td, timedelta): return "00:00:00,000"
    total_seconds = max(0, td.total_seconds())
    hours, remainder = divmod(total_seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    milliseconds = td.microseconds // 1000
    return f"{int(hours):02}:{int(minutes):02}:{int(seconds):02},{milliseconds:03}"

def format_srt_from_text_v16(srt_content, audio_filename, overlap_tolerance_td, chunk_duration_td, max_silence_seconds=10.0):
    srt_content = srt_content.strip().replace('\r\n', '\n').replace('\r', '\n')
    if srt_content.startswith("```srt"): srt_content = srt_content[srt_content.find('\n') + 1:]
    if srt_content.endswith("\n```"): srt_content = srt_content[:-4]

    lines = srt_content.split('\n')
    parsed_blocks = []
    current_block = {}
    
    for i, line in enumerate(lines):
        line = line.strip()
        
        if not line: # 空行通常是塊之間的間隔
            if current_block and "time_line" in current_block: # 如果當前塊已收集時間行，則表示一個塊結束
                parsed_blocks.append(current_block)
                current_block = {}
            continue
        
        if line.isdigit() and not current_block: # 序列號，且是新塊的開始
            current_block["original_index"] = int(line) - 1 # 轉換為0-based index
            current_block["text_lines"] = []
        elif "original_index" in current_block and "time_line" not in current_block and '-->' in line: # 時間戳行
            current_block["time_line"] = line
        elif "original_index" in current_block: # 文本行
            current_block["text_lines"].append(line)
        else: # 無法識別的行，可能是開頭的雜訊或錯誤格式
            logging.warning(f"[{audio_filename}] 無法識別的行 (行 {i+1}): '{line}'。已跳過。")

    if current_block and "time_line" in current_block: # 添加最後一個塊
        parsed_blocks.append(current_block)

    final_parsed_blocks = []
    for i, block_data in enumerate(parsed_blocks):
        original_index = block_data.get("original_index", i) # 如果沒有序列號，使用遍歷索引
        time_line = block_data.get("time_line")
        full_text = '\n'.join(block_data.get("text_lines", []))

        if not time_line:
            logging.warning(f"[{audio_filename}] 塊 {original_index+1} 缺少時間戳行，已跳過。")
            continue

        start_raw, end_raw = [t.strip() for t in time_line.split('-->')]
        start_td, end_td = parse_time_v10(start_raw), parse_time_v10(end_raw)
        
        is_valid = start_td is not None and end_td is not None and end_td > start_td
        
        # 即使文本為空，只要時間戳有效，也將其視為一個塊
        if not full_text:
            logging.warning(f"[{audio_filename}] 塊 {original_index+1} 的 API 回應文本為空。")
            # 不再 'continue'，而是將其添加到 parsed_blocks 中，以便後續處理
            
        final_parsed_blocks.append({
            "original_index": original_index,
            "start_td": start_td,
            "end_td": end_td,
            "text": full_text,
            "is_valid": is_valid,
            "time_line": time_line
        })

    # 使用 final_parsed_blocks 進行後續處理
    parsed_blocks = final_parsed_blocks
    
    corrected_blocks = []
    entry_counter = 1
    last_correct_end_td = timedelta(0)
    severe_correction_count = 0
    MAX_DURATION = timedelta(seconds=18)
    MAX_SILENCE_TD = timedelta(seconds=max_silence_seconds)
    
    i = 0
    while i < len(parsed_blocks):
        block = parsed_blocks[i]
        
        is_unparsable = not block["is_valid"]
        
        is_overlap_violation = False
        if overlap_tolerance_td.total_seconds() >= 0:
            is_overlap_violation = block["is_valid"] and block["start_td"] < (last_correct_end_td - overlap_tolerance_td)

        if is_unparsable or is_overlap_violation:
            severe_correction_count += 1
            log_msg = f"[{audio_filename}] SRT修正: 塊 {block['original_index']+1} " + (f"無法解析時間戳 '{block['time_line']}' 或時間倒流" if is_unparsable else f"檢測到時間軸重疊")
            logging.warning(log_msg)

            next_good_start_td = None
            for next_block in parsed_blocks[i+1:]:
                if next_block["is_valid"] and next_block["start_td"] >= last_correct_end_td:
                    next_good_start_td = next_block["start_td"]
                    break
            
            MAX_REASONABLE_GAP = timedelta(seconds=30)
            
            use_smart_logic = (
                next_good_start_td is not None and
                (next_good_start_td - last_correct_end_td) <= MAX_REASONABLE_GAP
            )

            dynamic_safe_duration = timedelta(seconds=2) if len(block["text"]) >= 8 else timedelta(seconds=1)

            if use_smart_logic:
                silence_duration = next_good_start_td - last_correct_end_td
                if silence_duration > MAX_SILENCE_TD:
                    logging.warning(f"[{audio_filename}]   -> 檢測到超長靜音 ({silence_duration.total_seconds():.1f}s)，採用「安全後貼」策略。")
                    end_td = next_good_start_td - timedelta(milliseconds=200)
                    start_td = end_td - dynamic_safe_duration
                else:
                    logging.warning(f"[{audio_filename}]   -> 檢測到常規靜音 ({silence_duration.total_seconds():.1f}s)，採用「智慧置中」策略。")
                    remaining_silence = silence_duration - dynamic_safe_duration
                    start_offset = max(timedelta(milliseconds=100), remaining_silence / 2)
                    start_td = last_correct_end_td + start_offset
                    end_td = start_td + dynamic_safe_duration
            else:
                if next_good_start_td:
                     logging.warning(f"[{audio_filename}]   -> 探測到過於遙遠的下個時間點，退回標準修正策略。")
                logging.warning(f"[{audio_filename}]   -> 採用標準向前修正策略。")
                dynamic_safe_duration = timedelta(seconds=3) if len(block["text"]) >= 8 else timedelta(seconds=1.5)
                start_td = last_correct_end_td + timedelta(milliseconds=100)
                end_td = start_td + dynamic_safe_duration
            
            if start_td < last_correct_end_td:
                start_td = last_correct_end_td + timedelta(milliseconds=50)
            if end_td <= start_td:
                end_td = start_td + dynamic_safe_duration

            block["start_td"] = start_td
            block["end_td"] = end_td
            block["is_valid"] = True

        if (block["end_td"] - block["start_td"]) > MAX_DURATION:
            logging.warning(f"[{audio_filename}] SRT修正: 塊 {block['original_index']+1} 檢測到超長持續時間 ({ (block['end_td'] - block['start_td']).total_seconds():.1f}s > {MAX_DURATION.total_seconds()}s)，已自動校正。")
            severe_correction_count += 1 # 將超長持續時間視為嚴重修正
            block["end_td"] = block["start_td"] + timedelta(seconds=5)

        if block["end_td"] > chunk_duration_td:
            logging.warning(f"[{audio_filename}] SRT修正: 塊 {block['original_index']+1} 結束時間 ({format_timedelta_v7(block['end_td'])}) 超過分段時長 ({format_timedelta_v7(chunk_duration_td)})，已校正至邊界。")
            block["end_td"] = chunk_duration_td
            if block["start_td"] >= block["end_td"]:
                block["start_td"] = block["end_td"] - timedelta(seconds=1)
                if block["start_td"] < timedelta(0):
                    block["start_td"] = timedelta(0)

        last_correct_end_td = block["end_td"]
        
        corrected_blocks.append(f"{entry_counter}\n{format_timedelta_v7(block['start_td'])} --> {format_timedelta_v7(block['end_td'])}\n{block['text']}")
        entry_counter += 1
        i += 1

    if not corrected_blocks:
        logging.warning(f"在 {audio_filename} 的回應中未能解析出任何有效的字幕塊。")
        raise SRTContentParseError(f"在 {audio_filename} 的回應中未能解析出任何有效的字幕塊。")
        
    return "\n\n".join(corrected_blocks) + "\n\n", severe_correction_count, last_correct_end_td

# CHANGED: 加入 max_retries、rate_limiter、retry_base、retry_cap 參數
def transcribe_audio(client, audio_path, prompt_text, model_name,
                     correction_threshold, overlap_tolerance, chunk_duration,
                     truncation_threshold, ffmpeg_executable, is_last_chunk=False,
                     max_retries=3, rate_limiter=None, retry_base=65, retry_cap=250):
    srt_path = os.path.splitext(audio_path)[0] + ".srt"
    file_basename = os.path.basename(audio_path)
    tokens_used, uploaded_file = 0, None
    overlap_tolerance_td = timedelta(seconds=overlap_tolerance)

    for attempt in range(max_retries):
        try:
            logging.info(f"[{file_basename} | 嘗試 {attempt+1}/{max_retries}] 正在上傳檔案...")
            if rate_limiter: rate_limiter.wait()  # NEW: 上傳前限速
            uploaded_file = client.files.upload(file=audio_path)

            logging.info(f"檔案已上傳。正在向模型 '{model_name}' 發送轉錄請求...")
            if rate_limiter: rate_limiter.wait()  # NEW: 呼叫前限速
            response = client.models.generate_content(model=model_name, contents=[prompt_text, uploaded_file])

            if hasattr(response, 'usage_metadata') and response.usage_metadata:
                tokens_used = response.usage_metadata.total_token_count
            if not response.text:
                raise EmptyResponseError("API 回應為空值 (empty response)。")

            with open(os.path.splitext(srt_path)[0] + ".raw.txt", 'w', encoding='utf-8') as f: f.write(response.text)
            
            chunk_duration_td = timedelta(seconds=chunk_duration)
            # NEW: format_srt_from_text_v16 現在會拋出 SRTContentParseError
            corrected_srt, severe_correction_count, last_subtitle_end_td = format_srt_from_text_v16(response.text, file_basename, overlap_tolerance_td, chunk_duration_td)
            
            if not is_final_srt_valid(corrected_srt):
                raise ValueError("校正後的 SRT 檔案結構驗證失敗 (序列號與時間戳數量不匹配)，觸發重試。")

            if corrected_srt and truncation_threshold > 0:
                effective_duration_td = timedelta(seconds=chunk_duration)
                duration_source_msg = f"標準分段時長 {chunk_duration}s"
                if is_last_chunk:
                    actual_duration_seconds = get_media_duration(audio_path, ffmpeg_executable)
                    if actual_duration_seconds is not None:
                        logging.info(f"正在為最後一個區塊 '{file_basename}' 獲取精確音訊時長: {actual_duration_seconds:.2f}s")
                        effective_duration_td = timedelta(seconds=actual_duration_seconds)
                        duration_source_msg = f"音訊實際長度 {actual_duration_seconds:.2f}s"
                    else:
                        logging.warning(f"無法獲取最後一個區塊 '{file_basename}' 的精確時長，將退回使用標準分段時長。")
                
                end_gap_seconds = (effective_duration_td - last_subtitle_end_td).total_seconds()
                
                if end_gap_seconds > truncation_threshold:
                    log_msg = (
                        f"SRT截斷警告: 區塊 '{file_basename}' 的結尾偵測到超過 {truncation_threshold} 秒的空白 "
                        f"({end_gap_seconds:.1f}s)。(基於 {duration_source_msg})。"
                        " 回應可能不完整，請手動檢查。"
                    )
                    logging.warning(log_msg)

            if severe_correction_count > correction_threshold:
                raise ValueError(f"SRT嚴重錯誤: 偵測到 {severe_correction_count} 次嚴重修正，超過閾值 {correction_threshold}。")
            with open(srt_path, 'w', encoding='utf-8') as f: f.write(corrected_srt)
            logging.info(f"成功！已將修正後的字幕儲存至: {os.path.basename(srt_path)}")
            return srt_path, tokens_used

        except Exception as e:
            # NEW: 捕獲到空回應或SRT解析異常，觸發重試
            if isinstance(e, (EmptyResponseError, SRTContentParseError)):
                logging.warning(f"處理 '{file_basename}' 時捕獲到轉錄或解析異常，將觸發重試: {e}")
            else:
                logging.warning(f"處理 '{file_basename}' 時捕獲到未預期異常，將觸發重試: {e}")

            if attempt < max_retries - 1:
                # NEW: 優先尊重 Retry-After
                retry_after = None
                try:
                    retry_after = getattr(getattr(e, "response", None), "headers", {}).get("Retry-After")
                except Exception:
                    pass
                if retry_after and str(retry_after).isdigit():
                    delay = int(retry_after)
                    logging.info(f"偵測到 Retry-After: {delay}s，暫停後重試...")
                    time.sleep(delay)
                else:
                    # NEW: 指數退避 + 全抖動
                    logging.info(f"將使用指數退避+全抖動策略進行重試 (第 {attempt+1} 次)...")
                    sleep_with_full_jitter(base=retry_base, attempt=attempt+1, cap=retry_cap)
                continue
            else:
                logging.error(f"已達最大重試次數，轉錄 '{file_basename}' 失敗。")
                return None, tokens_used

        finally:
            if uploaded_file:
                try:
                    client.files.delete(name=uploaded_file.name)
                except Exception as del_e:
                    logging.warning(f"刪除遠端檔案 '{uploaded_file.name}' 失敗: {del_e}")
    return None, tokens_used

def get_media_duration(file_path, ffmpeg_executable):
    command = [ffmpeg_executable, '-i', file_path]
    try:
        result = subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=False, creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == 'win32' else 0)
        output = result.stdout.decode('utf-8', errors='ignore') + result.stderr.decode('utf-8', errors='ignore')
        duration_match = re.search(r"Duration: (\d{2}):(\d{2}):(\d{2})\.(\d{2})", output)
        if duration_match:
            h, m, s, c = (int(g) for g in duration_match.groups())
            return (h * 3600) + (m * 60) + s + (c / 100.0)
        logging.warning("在 FFmpeg 輸出中找不到時長資訊。")
        return None
    except Exception as e:
        logging.error(f"使用 FFmpeg 獲取檔案時長失敗: {e}")
        return None

def get_chunk_file_regex(base_name, chunk_duration_seconds, extension):
    escaped_base = re.escape(base_name)
    escaped_ext = re.escape(f".{extension.lstrip('.')}")
    return re.compile(rf"^{escaped_base}_{chunk_duration_seconds}s_chunk_\d{{3}}{escaped_ext}$")

def split_audio(input_file, temp_dir, chunk_duration_seconds, ffmpeg_executable, recreate=False):
    logging.info("[STATUS] 正在檢查音訊區塊...")
    os.makedirs(temp_dir, exist_ok=True)
    base_name = os.path.splitext(os.path.basename(input_file))[0]
    chunk_file_regex_mp3 = get_chunk_file_regex(base_name, chunk_duration_seconds, "mp3")
    duration = get_media_duration(input_file, ffmpeg_executable)
    if duration is None: return []
    theoretical_chunks_count = math.ceil(duration / chunk_duration_seconds)
    chunk_base_name_prefix = f"{base_name}_{chunk_duration_seconds}s_chunk_"
    theoretical_chunk_names = {f"{chunk_base_name_prefix}{i:03d}.mp3" for i in range(theoretical_chunks_count)}
    if recreate:
        logging.info("[STATUS] 使用者選擇重新開始，將安全刪除所有符合當前設定的舊音訊與字幕區塊...")
        chunk_file_regex_srt = get_chunk_file_regex(base_name, chunk_duration_seconds, "srt")
        for f in os.listdir(temp_dir):
            if chunk_file_regex_mp3.match(f) or chunk_file_regex_srt.match(f):
                try: os.remove(os.path.join(temp_dir, f))
                except OSError as e: logging.error(f"刪除檔案 {f} 失敗: {e}")
    existing_chunks = {f for f in os.listdir(temp_dir) if chunk_file_regex_mp3.match(f)}
    missing_chunks = theoretical_chunk_names - existing_chunks
    if not missing_chunks:
        logging.info(f"所有 {theoretical_chunks_count} 個音訊區塊均已存在且設定相符。跳過分割。")
        return sorted([os.path.join(temp_dir, f) for f in theoretical_chunk_names])
    logging.info(f"[STATUS] 偵測到 {len(existing_chunks)} 個有效區塊，將僅補切 {len(missing_chunks)} 個缺失的區塊...")
    for i, chunk_name in enumerate(sorted(list(missing_chunks))):
        chunk_index = int(re.search(r'_chunk_(\d+)', chunk_name).group(1))
        start_time = chunk_index * chunk_duration_seconds
        output_path = os.path.join(temp_dir, chunk_name)
        command = [ffmpeg_executable, '-i', input_file, '-ss', str(start_time), '-t', str(chunk_duration_seconds), '-vn', '-acodec', 'libmp3lame', '-b:a', '192k', '-y', output_path]
        try:
            logging.info(f"正在補切區塊 {i+1}/{len(missing_chunks)}: {chunk_name}...")
            subprocess.run(command, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == 'win32' else 0)
        except Exception as e: logging.error(f"補切區塊 {chunk_name} 失敗: {e.stderr.decode() if hasattr(e, 'stderr') else e}")
    return sorted([os.path.join(temp_dir, f) for f in theoretical_chunk_names])

def get_safe_path(base_path):
    if not os.path.exists(base_path): return base_path
    directory, filename = os.path.split(base_path)
    name, ext = os.path.splitext(filename)
    counter = 1
    while True:
        new_path = os.path.join(directory, f"{name}_{counter}{ext}")
        if not os.path.exists(new_path): return new_path
        counter += 1

def merge_srts(srt_files, final_srt_path, chunk_duration_seconds):
    logging.info(f"[STATUS] 正在合併 {len(srt_files)} 個 SRT 檔案...")
    global_offset, entry_counter = timedelta(0), 1
    chunk_duration_td = timedelta(seconds=chunk_duration_seconds)
    with open(final_srt_path, 'w', encoding='utf-8') as outfile:
        sorted_srts = sorted(srt_files)
        for i, srt_file in enumerate(sorted(sorted_srts)):
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
                        start_td, end_td = parse_time_v10(start_str), parse_time_v10(end_str)
                        if start_td is None or end_td is None: continue
                        adjusted_start, adjusted_end = start_td + global_offset, end_td + global_offset
                        outfile.write(f"{entry_counter}\n{format_timedelta_v7(adjusted_start)} --> {format_timedelta_v7(adjusted_end)}\n" + '\n'.join(lines[2:]) + '\n\n')
                        entry_counter += 1
                if i < len(sorted_srts) - 1: global_offset += chunk_duration_td
            except FileNotFoundError:
                logging.warning(f"找不到要合併的 SRT 檔案: {srt_file}，將以空白時段取代。")
                if i < len(sorted_srts) - 1: global_offset += chunk_duration_td
            except Exception as e:
                logging.error(f"合併 SRT '{os.path.basename(srt_file)}' 時發生錯誤: {e}")
                if i < len(sorted_srts) - 1: global_offset += chunk_duration_td

def create_transcription_report(log_filepath, client, model_name, log_queue=None):
    report_filename = log_filepath.replace('_日誌_', '_SRT轉錄情況_')
    logging.info("[STATUS] 正在生成 SRT轉錄情況報告...")
    if not client:
        logging.error("報告生成失敗：找不到 API 用戶端。")
        return
    try:
        with open(log_filepath, 'r', encoding='utf-8') as f: log_content = f.read()
    except Exception as e:
        logging.error(f"無法讀取日誌檔案以生成報告: {e}")
        return
    prompt = f"""你是一位專業的軟體測試工程師，你的任務是分析以下這份由 Python 字幕轉錄腳本產生的日誌檔案。
本日誌可能包含多個獨立的轉錄工作流程（多軌處理）。請針對每個工作流程依序生成「SRT轉錄情況報告」，並將所有報告整合為同一份文件，依工作流程順序排列。
你的報告應包含以下幾點：
1.  **整體狀況總結**：簡要說明這次轉錄任務是成功還是失敗，以及是否發生了任何嚴重問題(例如是否有重大錯誤、嚴重警告、API回應空白或SRT截斷警告)。
2.  **音訊區塊處理**：列出每個音訊區塊 (chunk) 的處理情況。對於每個區塊，請註明：
    *   是否轉錄成功。
    *   是否觸發了重試機制？如果是，重試了幾次？最終是否成功？
    *   是否發生了大量的「嚴重修正」？
    *   API回應空白或SRT截斷警告
3.  **主要錯誤與警告**：如果日誌中出現了任何 Python 錯誤 (Exception)、FFmpeg 錯誤或 API 錯誤，請明確指出它們的內容和發生的時間點。請特別檢查日誌中是否包含「SRT截斷警告」。如果有的話，必須在報告中顯著地列出是哪個區塊發生了問題，並提醒使用者手動檢查該區塊的字幕完整性。其他明顯異常（如 API 回應文本為空)
4.  **資源使用**：記錄總共使用了多少 token。
5.  **最終結果**：說明最終的 SRT 檔案是否成功生成。
請使用繁體中文、Markdown 格式來呈現你的報告，使其易於閱讀。
---
以下是日誌內容 ---
{log_content}
--- 日誌內容結束 ---
"""
    try:
        logging.info(f"正在向模型 '{model_name}' 發送報告生成請求...")
        response = client.models.generate_content(model=model_name, contents=[prompt])
        
        # 檢查 API 回應是否有效
        if not response or not hasattr(response, 'text') or not response.text:
            raise ValueError("模型未返回有效的報告內容 (API response was empty or invalid)。")
            
        with open(report_filename, 'w', encoding='utf-8') as f: f.write("="*40 + "\nSRT轉錄情況報告 (AI 生成)\n" + "="*40 + "\n\n" + response.text)
        logging.info(f"已成功生成報告檔案: {report_filename}")
    except Exception as e:
        logging.error(f"生成 SRT轉錄情況報告時發生錯誤: {e}")
        if log_queue: log_queue.put(f"[RETRY_REPORT]{log_filepath}")
        else: print(f"[RETRY_REPORT]{log_filepath}")

def run_transcription_task(config, log_queue=None):
    exit_code = 0
    prompt_filepath = None

    # NEW: 連續空回應計數器
    empty_lock = Lock()
    empty_consecutive = {"n": 0}  # 使用 dict 以便在閉包內修改

    def _reset_empty_counter():
        """成功時重置計數器"""
        with empty_lock:
            if empty_consecutive["n"] > 0:
                logging.info(f"[EMPTY] API 回應恢復正常，連續空回應計數已歸零。")
                empty_consecutive["n"] = 0

    def _mark_empty_and_maybe_abort():
        """標記一次空回應，並檢查是否達到中止門檻"""
        with empty_lock:
            empty_consecutive["n"] += 1
            n = empty_consecutive["n"]
        
        # 從 config 物件安全地獲取閾值，若無則使用預設值 5
        thr = getattr(config, "empty_abort_threshold", 5)
        
        logging.warning(f"[EMPTY] 連續空回應次數 = {n} / {thr if thr > 0 else '∞'}")

        if thr > 0 and n >= thr:
            abort_msg = f"連續空回應已達門檻 {n} 次，強制終止整個流程。"
            logging.critical(f"[ABORT] {abort_msg}")
            # 拋出一個致命的 RuntimeError 來中斷 ThreadPoolExecutor
            raise RuntimeError(abort_msg)

    try:
        file_basename = os.path.splitext(os.path.basename(config.input_file))[0]
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_filename = os.path.join(APP_PATH, f"{file_basename}_日誌_{timestamp}.txt")
        setup_logging(log_filename, config.verbose, log_queue)
        logging.info("="*80 + f"\n開始執行任務: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n版本: {os.path.basename(__file__)}\n處理檔案: {config.input_file}\n" + "="*80)
        os.makedirs(config.temp_dir, exist_ok=True)
        
        prompt_text = config.prompt_text if hasattr(config, 'prompt_text') and config.prompt_text else ""
        if not prompt_text: 
            logging.warning("提示為空。")
        else:
            prompt_filename = f"{file_basename}_prompt_{timestamp}.txt"
            prompt_filepath = os.path.join(config.temp_dir, prompt_filename)
            try:
                with open(prompt_filepath, 'w', encoding='utf-8') as f:
                    f.write(prompt_text)
                logging.info(f"已將本次執行的 Prompt 儲存至: {prompt_filepath}")
            except Exception as e:
                logging.error(f"儲存 Prompt 檔案失敗: {e}")
                prompt_filepath = None

        if config.merge_only:
            logging.info("【僅合併模式】啟動...")
            chunk_file_regex_srt = get_chunk_file_regex(file_basename, config.chunk_duration, "srt")
            all_chunk_srts = sorted([os.path.join(config.temp_dir, f) for f in os.listdir(config.temp_dir) if chunk_file_regex_srt.match(f)])
            if not all_chunk_srts:
                logging.error(f"在 '{config.temp_dir}' 中找不到任何符合模式的 .srt 區塊檔案。")
                raise SystemExit(1)
            final_srt_path = get_safe_path(os.path.join(APP_PATH, f"{file_basename}.srt"))
            merge_srts(all_chunk_srts, final_srt_path, config.chunk_duration)
            logging.info(f"僅合併模式完成。最終 SRT 檔案位於: {final_srt_path}")
            return 0
        client = None
        try:
            api_key = config.api_key or os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
            client = genai.Client(api_key=api_key)
            logging.info(f"成功建立 API 用戶端。將使用模型: {config.model_name}")
        except Exception as e:
            logging.error(f"建立 API 用戶端失敗: {e}")
            raise SystemExit(1)
        
        # CHANGED: 在 split 完、檢查 resume 後，組裝待處理清單
        chunk_mp3_files = split_audio(config.input_file, config.temp_dir, config.chunk_duration, config.ffmpeg_path, config.recreate)
        transcription_was_performed = False
        if chunk_mp3_files:
            total_tokens_used, all_tasks_were_skipped = 0, True
            to_process = []
            for i, chunk_mp3_path in enumerate(chunk_mp3_files):
                logging.info(f"[STATUS] 正在檢查區塊 {i+1}/{len(chunk_mp3_files)}: {os.path.basename(chunk_mp3_path)}")
                chunk_srt_path = os.path.splitext(chunk_mp3_path)[0] + ".srt"
                if config.resume and os.path.exists(chunk_srt_path) and os.path.getsize(chunk_srt_path) > 0:
                    logging.info(f"偵測到已存在的有效 SRT 檔案，跳過轉錄區塊: {os.path.basename(chunk_mp3_path)}")
                    continue
                all_tasks_were_skipped = False
                to_process.append((i, chunk_mp3_path))

            # NEW: 單程序共用的 RPM 閘門
            rate_limiter = MinuteRateLimiter(getattr(config, "rpm", 3))
            workers = max(1, getattr(config, "workers", 2))

            if to_process:
                transcription_was_performed = True
                last_index = len(chunk_mp3_files) - 1
                logging.info(f"啟動併發處理：workers={workers}, rpm={rate_limiter.rpm}（單程序共用）")

                def _job(i, path):
                    is_last = (i == last_index)
                    try:
                        truncation_threshold_value = getattr(config, 'truncation_threshold', 60)
                        srt_path, tokens = transcribe_audio(
                            client, path, prompt_text, config.model_name,
                            config.correction_threshold, config.overlap_tolerance, config.chunk_duration,
                            truncation_threshold_value, config.ffmpeg_path, is_last_chunk=is_last,
                            max_retries=getattr(config, "max_retries", 3),
                            rate_limiter=rate_limiter,
                            retry_base=getattr(config, "retry_base", 65),
                            retry_cap=getattr(config, "retry_cap", 250),
                        )
                        # 任何一次成功都重置計數器
                        _reset_empty_counter()
                        return (i, srt_path, tokens or 0)
                    except EmptyResponseError:
                        # 捕獲到空回應，標記並可能中止
                        _mark_empty_and_maybe_abort()
                        # 返回 None 表示此單塊失敗，讓主流程繼續
                        return (i, None, 0)
                    except SRTContentParseError: # NEW: 捕獲SRT內容解析錯誤
                        _mark_empty_and_maybe_abort() # 視為類似空回應的嚴重錯誤
                        return (i, None, 0)

                try:
                    with ThreadPoolExecutor(max_workers=workers) as ex:
                        futures = [ex.submit(_job, i, p) for i, p in to_process]
                        for fut in as_completed(futures):
                            try:
                                # fut.result() 會重新拋出 _job 中未被捕獲的異常，例如 RuntimeError
                                i, srt_path, tokens = fut.result()
                                total_tokens_used += tokens
                            except Exception as exc:
                                # 這個 except 主要捕捉 result() 拋出的非 RuntimeError 的其他問題
                                # NEW: 包含檔案路徑
                                logging.error(f'任務 {i} ({os.path.basename(path)}) 在取得結果時產生例外: {exc}')
                except RuntimeError as fatal:
                    # 這是捕捉 _mark_empty_and_maybe_abort 拋出的致命錯誤的地方
                    logging.critical(f"任務因致命錯誤而中止: {fatal}")
                    # 可以在此處添加額外的清理邏輯
                    raise SystemExit(1) # 直接中止後續流程


            if all_tasks_were_skipped:
                logging.info("所有區塊轉錄都已完成並跳過。")

            logging.info(f"所有音訊區塊處理完成。總共使用了約 {total_tokens_used} 個 token。")
            
            all_chunk_srts = [os.path.splitext(p)[0] + ".srt" for p in chunk_mp3_files]
            valid_srts = [s for s in all_chunk_srts if os.path.exists(s) and os.path.getsize(s) > 0]
            if not valid_srts and transcription_was_performed:
                logging.error("所有區塊轉錄均失敗或找不到有效的 SRT 檔案。任務中止。")
                raise SystemExit(1)
            final_srt_path = get_safe_path(os.path.join(APP_PATH, f"{file_basename}.srt"))
            merge_srts(all_chunk_srts, final_srt_path, config.chunk_duration)
            logging.info(f"工作流程完成。最終 SRT 檔案位於: {final_srt_path}")
        if config.enable_report:
            if transcription_was_performed:
                create_transcription_report(log_filename, client, config.model_name, log_queue)
            else:
                logging.info("沒有執行新的轉錄，跳過 SRT轉錄情況報告的生成。")
    except SystemExit as e:
        exit_code = e.code if e.code is not None else 1
        logging.error(f"任務因 SystemExit 中止 (退出碼: {exit_code})")
    except Exception as e:
        exit_code = 1
        logging.error(f"任務發生未預期的嚴重錯誤: {e}", exc_info=True)
    finally:
        logging.info(f"任務執行完畢。退出碼: {exit_code}")
        if prompt_filepath and hasattr(config, 'keep_prompt_file') and not config.keep_prompt_file:
            try:
                os.remove(prompt_filepath)
                logging.info(f"已自動刪除本次執行的 Prompt 檔案: {prompt_filepath}")
            except OSError as e:
                logging.warning(f"自動刪除 Prompt 檔案失敗: {e}")
        return exit_code

def run_partial_transcription_task(config, log_queue=None):
    """
    執行局部轉錄任務：切割指定時間區段的音訊，進行轉錄，並將結果的時間軸校正後儲存。
    """
    exit_code = 0
    temp_audio_path = None
    try:
        file_basename = os.path.splitext(os.path.basename(config.input_file))[0]
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_filename = os.path.join(APP_PATH, f"{file_basename}_日誌_partial_{timestamp}.txt")
        setup_logging(log_filename, config.verbose, log_queue)
        
        logging.info("="*80 + f"\n開始執行局部轉錄任務: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n版本: {os.path.basename(__file__)}\n" + "="*80)
        logging.info(f"原始檔案: {config.input_file}")
        logging.info(f"轉錄區段: {config.start_time} --> {config.end_time}")

        start_td = parse_time_v10(config.start_time)
        end_td = parse_time_v10(config.end_time)

        if start_td is None or end_td is None or end_td <= start_td:
            logging.error(f"提供的時間範圍無效: {config.start_time} -> {config.end_time}。請檢查格式 (HH:MM:SS,ms) 且結束時間需大於開始時間。")
            return 1

        duration_td = end_td - start_td
        
        # 建立唯一的暫存音訊檔名
        time_str_for_filename = config.start_time.replace(":", "-").replace(",", "_")
        temp_audio_filename = f"{file_basename}_partial_{time_str_for_filename}.mp3"
        temp_audio_path = os.path.join(config.temp_dir, temp_audio_filename)
        
        # 使用 FFmpeg 切割音訊
        logging.info(f"正在使用 FFmpeg 從 '{config.input_file}' 切割音訊片段...")
        command = [
            config.ffmpeg_path,
            '-i', config.input_file,
            '-ss', str(start_td.total_seconds()),
            '-t', str(duration_td.total_seconds()),
            '-vn', '-acodec', 'libmp3lame', '-b:a', '192k', '-y',
            temp_audio_path
        ]
        try:
            subprocess.run(command, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == 'win32' else 0)
            logging.info(f"成功切割音訊片段至: {temp_audio_path}")
        except Exception as e:
            logging.error(f"使用 FFmpeg 切割音訊失敗: {e.stderr.decode() if hasattr(e, 'stderr') else e}")
            return 1

        # 建立 API 用戶端
        client = None
        try:
            api_key = config.api_key or os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
            client = genai.Client(api_key=api_key)
            logging.info(f"成功建立 API 用戶端，將使用模型: {config.model_name}")
        except Exception as e:
            logging.error(f"建立 API 用戶端失敗: {e}")
            return 1

        # 執行轉錄 (注意：此處的 chunk_duration 應為片段自身的長度，truncation_threshold 應停用)
        partial_srt_path, _ = transcribe_audio(
            client,
            temp_audio_path,
            config.prompt_text,
            config.model_name,
            config.correction_threshold,
            config.overlap_tolerance,
            chunk_duration=duration_td.total_seconds(),
            truncation_threshold=0, # 局部轉錄不檢查結尾空白
            ffmpeg_executable=config.ffmpeg_path,
            is_last_chunk=True, # 視為單一的最後區塊
            max_retries=config.max_retries if hasattr(config, 'max_retries') else 3
        )

        if not partial_srt_path or not os.path.exists(partial_srt_path):
            logging.error("局部轉錄失敗，未能生成 SRT 檔案。")
            exit_code = 1
            return exit_code

        # --- 時間軸校正 ---
        logging.info("轉錄完成，正在進行時間軸校正...")
        with open(partial_srt_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        adjusted_content = ""
        for entry in content.strip().split('\n\n'):
            if not entry.strip(): continue
            lines = entry.split('\n')
            if len(lines) < 2 or '-->' not in lines[1]:
                adjusted_content += entry + '\n\n'
                continue
            
            try:
                seq_num = lines[0]
                time_line = lines[1]
                text_lines = '\n'.join(lines[2:])
                
                start_str, end_str = [t.strip() for t in time_line.split('-->')]
                entry_start_td = parse_time_v10(start_str)
                entry_end_td = parse_time_v10(end_str)

                if entry_start_td is None or entry_end_td is None:
                    raise ValueError("無法解析時間戳")

                # 加上使用者輸入的起始時間作為偏移量
                adjusted_start = entry_start_td + start_td
                adjusted_end = entry_end_td + start_td
                
                adjusted_content += f"{seq_num}\n{format_timedelta_v7(adjusted_start)} --> {format_timedelta_v7(adjusted_end)}\n{text_lines}\n\n"
            except Exception as e:
                logging.warning(f"處理 SRT 條目時出錯，將保留原始條目: {e}\n原始條目:\n{entry}")
                adjusted_content += entry + '\n\n'

        # 儲存校正後的 SRT 檔案
        final_srt_filename = f"{file_basename}_partial_{time_str_for_filename}.srt"
        final_srt_path = get_safe_path(os.path.join(APP_PATH, final_srt_filename))
        
        with open(final_srt_path, 'w', encoding='utf-8') as f:
            f.write(adjusted_content)
        
        logging.info(f"時間軸校正完成！最終局部 SRT 檔案儲存於: {final_srt_path}")

    except Exception as e:
        exit_code = 1
        logging.error(f"局部轉錄任務發生未預期的嚴重錯誤: {e}", exc_info=True)
    finally:
        if temp_audio_path and os.path.exists(temp_audio_path) and not getattr(config, 'keep_partial_audio', False):
            try:
                os.remove(temp_audio_path)
                logging.info(f"已自動刪除暫存音訊檔: {temp_audio_path}")
            except OSError as e:
                logging.warning(f"自動刪除暫存音訊檔失敗: {e}")
        logging.info(f"局部轉錄任務執行完畢。退出碼: {exit_code}")
        return exit_code

def run_summarize_only_task(config, log_queue=None):
    exit_code = 0
    try:
        setup_logging(config.log_file, config.verbose, log_queue)
        logging.info("【僅摘要模式】啟動...")
        api_key = config.api_key or os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
        client = genai.Client(api_key=api_key)
        logging.info(f"成功建立 API 用戶端。將使用模型: {config.model_name}")
        create_transcription_report(config.log_file, client, config.model_name, log_queue)
        logging.info("僅摘要模式完成。")
    except Exception as e:
        exit_code = 1
        logging.error(f"在僅摘要模式下建立 API 用戶端或生成報告時失敗: {e}")
    finally:
        return exit_code

def main_cli():
    force_utf8_encoding()
    start_time = time.time()
    parser = argparse.ArgumentParser(description="自動化音訊/視訊轉錄工作流程。")
    
    # --- 通用參數 ---
    parser.add_argument("--file", dest="input_file", help="要處理的大型音訊或視訊檔案路徑。")
    parser.add_argument("--ffmpeg_path", help="FFmpeg 執行檔的完整路徑。")
    parser.add_argument("--temp_dir", default=os.path.join(APP_PATH, "temp"), help="儲存臨時音訊區塊和 SRT 檔案的目錄。")
    parser.add_argument("--api_key", help="您的 API 金鑰。")
    parser.add_argument("--model_name", default="models/gemini-2.5-pro", help="要使用的 Gemini 模型名稱。")
    parser.add_argument("--prompt_file", help="包含主要指令的提示檔案路徑。")
    parser.add_argument("--verbose", action='store_true', help="啟用更詳細的日誌記錄。")
    parser.add_argument('--keep_prompt_file', action='store_true', help='保留為本次執行建立的 Prompt 檔案以供偵錯。')

    # --- 完整轉錄模式參數 ---
    parser.add_argument("--chunk_duration", type=int, default=600, help="每個音訊區塊的持續時間（秒）。")
    parser.add_argument("--correction_threshold", type=int, default=6, help="觸發自動重跑的嚴重修正次數閾值。")
    parser.add_argument("--overlap_tolerance", type=float, default=0.5, help="允許的字幕時間軸重疊容忍秒數。設為負數 (例如 -1) 可完全關閉重疊偵測。")
    parser.add_argument("--truncation_threshold", type=int, default=60, help="可疑的結尾空白秒數閾值，用於偵測可能被截斷的回應。設為 0 可停用此檢查。")
    parser.add_argument("--report", dest="enable_report", action='store_true', help="處理完成後，生成 SRT 轉錄情況報告。")
    parser.add_argument("--resume", action='store_true', help="從上次的中斷處繼續任務。")
    parser.add_argument("--recreate", action='store_true', help="強制重新建立所有區塊。")

    # NEW: 併發與限速、重試策略參數
    parser.add_argument("--workers", type=int, default=1, help="併發處理的工作執行緒數（建議 2~4）。")
    parser.add_argument("--rpm", type=int, default=3, help="單程序每分鐘允許的最大請求數。")
    parser.add_argument("--max_retries", type=int, default=3, help="單個區塊的最大重試次數。")
    parser.add_argument("--retry_base", type=int, default=65, help="指數退避的基準秒數（第一次重試的最大等待上限）。")
    parser.add_argument("--retry_cap", type=int, default=250, help="單次退避的最大秒數上限。")
    parser.add_argument("--empty_abort_threshold", type=int, default=5, help="連續空回應達到此次數就終止整個流程；0=關閉。")

    # --- 局部轉錄模式參數 ---
    parser.add_argument("--partial_only", action='store_true', help="僅執行局部轉錄操作。")
    parser.add_argument("--start_time", help="局部轉錄的開始時間 (格式: HH:MM:SS,ms)。")
    parser.add_argument("--end_time", help="局部轉錄的結束時間 (格式: HH:MM:SS,ms)。")
    parser.add_argument('--keep_partial_audio', action='store_true', help='保留為局部轉錄切割出的暫存音訊檔以供偵錯。')

    # --- 其他獨立模式 ---
    parser.add_argument("--merge_only", action='store_true', help="僅執行合併 SRT 檔案的操作。")
    parser.add_argument("--summarize_only", action='store_true', help="僅重新生成 AI 報告。需要提供 --log_file。")
    parser.add_argument("--log_file", help="用於生成 AI 報告的日誌檔案路徑。")

    args = parser.parse_args()
    config = SimpleNamespace(**vars(args))

    # 讀取 Prompt 檔案內容
    if hasattr(config, 'prompt_file') and config.prompt_file and os.path.exists(config.prompt_file):
        with open(config.prompt_file, 'r', encoding='utf-8') as f:
            config.prompt_text = f.read()
    else:
        config.prompt_text = ""

    # --- 根據模式執行對應的任務 ---
    exit_code = 0
    if config.summarize_only:
        if not config.log_file or not os.path.exists(config.log_file):
            sys.exit("錯誤: 使用 --summarize_only 模式時，必須提供一個有效的 --log_file 路徑。")
        exit_code = run_summarize_only_task(config)
    elif config.partial_only:
        if not all([config.input_file, config.ffmpeg_path, config.start_time, config.end_time]):
            parser.error("使用 --partial_only 模式時，必須提供 '--file', '--ffmpeg_path', '--start_time' 和 '--end_time' 參數。")
        exit_code = run_partial_transcription_task(config)
    elif not config.input_file or not config.ffmpeg_path:
        parser.error("在標準轉錄或合併模式下，必須提供 '--file' 和 '--ffmpeg_path' 參數。")
    else:
        # 預設執行完整轉錄流程
        exit_code = run_transcription_task(config)

    print(f"\n總共耗時: {time.time() - start_time:.2f} 秒。")
    sys.exit(exit_code)

if __name__ == "__main__":
    try:
        import google.genai
    except ImportError:
        print("="*80 + "\n【重要環境配置錯誤】\n...")
        sys.exit(1)
    main_cli()
