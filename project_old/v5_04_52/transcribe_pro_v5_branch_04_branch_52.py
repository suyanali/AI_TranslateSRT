# transcribe_pro_v5_branch_04_branch_52.py
# 版本號: v5_4_52_20250809
# 修改內容簡述:
# 1. 【日誌過濾】: 根據使用者提供的日誌範例，精準調整日誌過濾。將 `google.generativeai` 層級設為 INFO (以顯示 HTTP Request)，同時將 `urllib3` 等更底層的函式庫設為 WARNING，達到最佳的日誌清晰度。
# 2. 【路徑修正】: 延續 v51 的修正，確保所有輸出檔案都在主程式目錄。

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
    
    # 根據使用者日誌範例精準過濾
    logging.getLogger("google.generativeai").setLevel(logging.INFO)
    logging.getLogger("urllib3").setLevel(logging.WARNING)
    logging.getLogger("google.api_core").setLevel(logging.WARNING)

try:
    from google import genai
except ImportError:
    pass

def parse_time_v9(time_str):
    ts = time_str.strip().replace(':_', ':')
    if re.match(r'^\d+:\d+:\d+:\d+[,.]\d+$', ts): return None
    ts = re.sub(r'(\d+:\d+:\d+)[.:](\d+)$', r'\1,\2', ts)
    ts = re.sub(r'^(\d+:\d+)[.,:](\d+)$', r'00:\1,\2', ts)
    match = re.match(r'^(\d+):(\d+):(\d+),(\d+)$', ts)
    if not match:
        logging.warning(f"無法解析時間戳格式: {ts}")
        return None
    try:
        h, m, s, ms = (int(g) for g in match.groups())
        if s >= 60:
            logging.warning(f"秒數無效 ({s})，校正為 59。原始: '{time_str}'")
            s = 59
        if m >= 60:
            logging.warning(f"分鐘數無效 ({m})，校正為 59。原始: '{time_str}'")
            m = 59
        return timedelta(hours=h, minutes=m, seconds=s, milliseconds=ms)
    except ValueError:
        logging.error(f"時間戳中的數字無法轉換: {ts}")
        return None

def format_timedelta_v7(td):
    if not isinstance(td, timedelta): return "00:00:00,000"
    total_seconds = max(0, td.total_seconds())
    hours, remainder = divmod(total_seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    milliseconds = td.microseconds // 1000
    return f"{int(hours):02}:{int(minutes):02}:{int(seconds):02},{milliseconds:03}"

def format_srt_from_text_v15(srt_content, audio_filename, overlap_tolerance_td):
    srt_content = srt_content.strip().replace('\r\n', '\n').replace('\r', '\n')
    if srt_content.startswith("```srt"): srt_content = srt_content[srt_content.find('\n') + 1:]
    if srt_content.endswith("\n```"): srt_content = srt_content[:-4]
    pattern = re.compile(r'(\n*\d+\n[\d:,._]+\s*-->\s*[\d:,._]+)')
    processed_content = pattern.sub(r'<SRT_BLOCK_SEPARATOR>\1', srt_content)
    if processed_content.startswith('<SRT_BLOCK_SEPARATOR>'): processed_content = processed_content[len('<SRT_BLOCK_SEPARATOR>'):]
    blocks = [block.strip() for block in processed_content.split('<SRT_BLOCK_SEPARATOR>') if block.strip()]
    corrected_blocks, entry_counter, last_correct_end_td, severe_correction_count = [], 1, timedelta(0), 0
    MAX_DURATION = timedelta(minutes=3)
    for i, block in enumerate(blocks):
        lines = [line.strip() for line in block.split('\n') if line.strip()]
        if len(lines) < 2: continue
        time_line, full_text = lines[1], '\n'.join(lines[2:])
        if not full_text:
            logging.warning(f"塊 {i+1} 的 API 回應文本為空，已跳過。")
            continue
        try:
            dynamic_safe_duration = timedelta(seconds=2) if len(full_text) >= 8 else timedelta(seconds=1)
            start_raw, end_raw = [t.strip() for t in time_line.split('-->')]
            start_td, end_td = parse_time_v9(start_raw), parse_time_v9(end_raw)
            is_unparsable = start_td is None or end_td is None
            is_overlap_violation = not is_unparsable and start_td < (last_correct_end_td - overlap_tolerance_td)
            if is_unparsable or is_overlap_violation:
                severe_correction_count += 1
                log_msg = f"SRT修正: 塊 {i+1} " + (f"無法解析時間戳 '{time_line}'" if is_unparsable else f"檢測到時間軸重疊") + "，已自動校正。"
                logging.warning(log_msg)
                start_td = last_correct_end_td + timedelta(milliseconds=100)
                end_td = start_td + dynamic_safe_duration
            if end_td <= start_td:
                logging.warning(f"SRT修正: 塊 {i+1} 檢測到時間倒流，已自動校正。")
                end_td = start_td + dynamic_safe_duration
            if (end_td - start_td) > MAX_DURATION:
                logging.warning(f"SRT修正: 塊 {i+1} 檢測到超長持續時間，已自動校正。")
                end_td = start_td + dynamic_safe_duration
            last_correct_end_td = end_td
            corrected_blocks.append(f"{entry_counter}\n{format_timedelta_v7(start_td)} --> {format_timedelta_v7(end_td)}\n{full_text}")
            entry_counter += 1
        except Exception as e:
            logging.error(f"處理字幕塊 {i+1} 時發生嚴重錯誤: {e}")
    if not corrected_blocks:
        logging.warning(f"在 {audio_filename} 的回應中未能解析出任何有效的字幕塊。")
        return "", severe_correction_count
    return "\n\n".join(corrected_blocks) + "\n\n", severe_correction_count

def transcribe_audio(client, audio_path, prompt_text, model_name, correction_threshold, overlap_tolerance, max_retries=3):
    srt_path = os.path.splitext(audio_path)[0] + ".srt"
    file_basename = os.path.basename(audio_path)
    tokens_used, retry_delays, overlap_tolerance_td, uploaded_file = 0, [65, 130, 250], timedelta(seconds=overlap_tolerance), None
    for attempt in range(max_retries):
        try:
            logging.info(f"[{file_basename} | 嘗試 {attempt+1}/{max_retries}] 正在上傳檔案...")
            uploaded_file = client.files.upload(file=audio_path)
            logging.info(f"檔案已上傳。正在向模型 '{model_name}' 發送轉錄請求...")
            response = client.models.generate_content(model=model_name, contents=[prompt_text, uploaded_file])
            if hasattr(response, 'usage_metadata') and response.usage_metadata: tokens_used = response.usage_metadata.total_token_count
            if not response.text: raise ValueError("API 回應為空值 (empty response)。")
            with open(os.path.splitext(srt_path)[0] + ".raw.txt", 'w', encoding='utf-8') as f: f.write(response.text)
            corrected_srt, severe_correction_count = format_srt_from_text_v15(response.text, file_basename, overlap_tolerance_td)
            if severe_correction_count > correction_threshold: raise ValueError(f"SRT嚴重錯誤: 偵測到 {severe_correction_count} 次嚴重修正，超過閾值 {correction_threshold}。")
            with open(srt_path, 'w', encoding='utf-8') as f: f.write(corrected_srt)
            logging.info(f"成功！已將修正後的字幕儲存至: {os.path.basename(srt_path)}")
            return srt_path, tokens_used
        except Exception as e:
            logging.warning(f"處理 '{file_basename}' 時捕獲到異常: {e}")
            if attempt < max_retries - 1:
                logging.info(f"將在 {retry_delays[attempt]} 秒後重試...")
                time.sleep(retry_delays[attempt])
            else:
                logging.error(f"已達最大重試次數，轉錄 '{file_basename}' 失敗。" )
                return None, tokens_used
        finally:
            if uploaded_file:
                try: client.files.delete(name=uploaded_file.name)
                except Exception as del_e: logging.warning(f"刪除遠端檔案 '{uploaded_file.name}' 失敗: {del_e}")
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
        logging.info(f"所有 {len(theoretical_chunk_names)} 個音訊區塊均已存在且設定相符。跳過分割。")
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
                        start_td, end_td = parse_time_v9(start_str), parse_time_v9(end_str)
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
請從中提取關鍵資訊，並以清晰、有條理的方式，生成一份「SRT轉錄情況報告」。
你的報告應包含以下幾點：
1.  **整體狀況總結**：簡要說明這次轉錄任務是成功還是失敗，以及是否發生了任何嚴重問題。
2.  **音訊區塊處理**：列出每個音訊區塊 (chunk) 的處理情況。對於每個區塊，請註明：
    *   是否轉錄成功。
    *   是否觸發了重試機制？如果是，重試了幾次？最終是否成功？
    *   是否發生了大量的「嚴重修正」？
3.  **主要錯誤與警告**：如果日誌中出現了任何 Python 錯誤 (Exception)、FFmpeg 錯誤或 API 錯誤，請明確指出它們的內容和發生的時間點。
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
        with open(report_filename, 'w', encoding='utf-8') as f: f.write("="*40 + "\nSRT轉錄情況報告 (AI 生成)\n" + "="*40 + "\n\n" + response.text)
        logging.info(f"已成功生成報告檔案: {report_filename}")
    except Exception as e:
        logging.error(f"生成 SRT轉錄情況報告時發生錯誤: {e}")
        if log_queue: log_queue.put(f"[RETRY_REPORT]{log_filepath}")
        else: print(f"[RETRY_REPORT]{log_filepath}")

def run_transcription_task(config, log_queue=None):
    exit_code = 0
    try:
        file_basename = os.path.splitext(os.path.basename(config.input_file))[0]
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_filename = os.path.join(APP_PATH, f"{file_basename}_日誌_{timestamp}.txt")
        setup_logging(log_filename, config.verbose, log_queue)
        logging.info("="*80 + f"\n開始執行任務: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n版本: {os.path.basename(__file__)}\n處理檔案: {config.input_file}\n" + "="*80)
        os.makedirs(config.temp_dir, exist_ok=True)
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
        prompt_text = config.prompt_text if hasattr(config, 'prompt_text') and config.prompt_text else ""
        if not prompt_text: logging.warning("提示為空。")
        chunk_mp3_files = split_audio(config.input_file, config.temp_dir, config.chunk_duration, config.ffmpeg_path, config.recreate)
        transcription_was_performed = False
        if chunk_mp3_files:
            total_tokens_used, all_tasks_were_skipped = 0, True
            for i, chunk_mp3_path in enumerate(chunk_mp3_files):
                logging.info(f"[STATUS] 正在檢查區塊 {i+1}/{len(chunk_mp3_files)}: {os.path.basename(chunk_mp3_path)}")
                chunk_srt_path = os.path.splitext(chunk_mp3_path)[0] + ".srt"
                if config.resume and os.path.exists(chunk_srt_path) and os.path.getsize(chunk_srt_path) > 0:
                    logging.info(f"偵測到已存在的有效 SRT 檔案，跳過轉錄區塊: {os.path.basename(chunk_mp3_path)}")
                    continue
                all_tasks_were_skipped = False
                transcription_was_performed = True
                _, tokens = transcribe_audio(client, chunk_mp3_path, prompt_text, config.model_name, config.correction_threshold, config.overlap_tolerance)
                total_tokens_used += tokens if tokens else 0
            if all_tasks_were_skipped: logging.info("所有區塊轉錄都已完成並跳過。")
            logging.info(f"所有音訊區塊處理完成。總共使用了約 {total_tokens_used} 個 token。")
            all_chunk_srts = [os.path.splitext(p)[0] + ".srt" for p in chunk_mp3_files]
            valid_srts = [s for s in all_chunk_srts if os.path.exists(s) and os.path.getsize(s) > 0]
            if not valid_srts:
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
        # The GUI now handles the TASK_COMPLETE signal
        # if log_queue: log_queue.put(('TASK_COMPLETE', exit_code))
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
        # The GUI now handles the TASK_COMPLETE signal
        # if log_queue: log_queue.put(('TASK_COMPLETE', exit_code))
        return exit_code

def main_cli():
    force_utf8_encoding()
    start_time = time.time()
    parser = argparse.ArgumentParser(description="自動化音訊/視訊轉錄工作流程。")
    parser.add_argument("--file", dest="input_file", help="要處理的大型音訊或視訊檔案路徑。")
    parser.add_argument("--ffmpeg_path", help="FFmpeg 執行檔的完整路徑。")
    parser.add_argument("--temp_dir", default=os.path.join(APP_PATH, "temp"), help="儲存臨時音訊區塊和 SRT 檔案的目錄。")
    parser.add_argument("--chunk_duration", type=int, default=600, help="每個音訊區塊的持續時間（秒）。")
    parser.add_argument("--prompt_file", help="包含主要指令的提示檔案路徑。")
    parser.add_argument("--api_key", help="您的 API 金鑰。")
    parser.add_argument("--model_name", default="models/gemini-2.5-pro", help="要使用的 Gemini 模型名稱。")
    parser.add_argument("--correction_threshold", type=int, default=5, help="觸發自動重跑的嚴重修正次數閾值。")
    parser.add_argument("--overlap_tolerance", type=float, default=0.5, help="允許的字幕時間軸重疊容忍秒數。")
    parser.add_argument("--report", dest="enable_report", action='store_true', help="處理完成後，生成 SRT 轉錄情況報告。")
    parser.add_argument("--verbose", action='store_true', help="啟用更詳細的日誌記錄。")
    parser.add_argument("--resume", action='store_true', help="從上次的中斷處繼續任務。")
    parser.add_argument("--recreate", action='store_true', help="強制重新建立所有區塊。")
    parser.add_argument("--merge_only", action='store_true', help="僅執行合併 SRT 檔案的操作。")
    parser.add_argument("--summarize_only", action='store_true', help="僅重新生成 AI 報告。需要提供 --log_file。")
    parser.add_argument("--log_file", help="用於生成 AI 報告的日誌檔案路徑。")
    args = parser.parse_args()
    config = SimpleNamespace(**vars(args))
    if config.summarize_only:
        if not config.log_file or not os.path.exists(config.log_file):
            sys.exit("錯誤: 使用 --summarize_only 模式時，必須提供一個有效的 --log_file 路徑。")
        run_summarize_only_task(config)
    elif not config.input_file or not config.ffmpeg_path:
        parser.error("在非 --summarize_only 模式下，必須提供 '--file' 和 '--ffmpeg_path' 參數。")
    else:
        if hasattr(config, 'prompt_file') and config.prompt_file and os.path.exists(config.prompt_file):
            with open(config.prompt_file, 'r', encoding='utf-8') as f: config.prompt_text = f.read()
        else: config.prompt_text = ""
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
