# transcribe_pro_v5_branch_04_branch_11.py
# 
# 
# 

import os
import sys
import subprocess
import re
from datetime import datetime, timedelta
import argparse
import logging
import time
import io

# ---
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
    """
    """
    global log_string_io
    log_string_io = io.StringIO()

    for handler in logging.root.handlers[:]:
        logging.root.removeHandler(handler)

    log_format = '%(asctime)s - %(levelname)s - %(message)s'
    tee_stdout = Tee(sys.stdout, log_string_io)
    
    logging.basicConfig(
        level=logging.INFO,
        format=log_format,
        stream=tee_stdout,
        force=True
    )
    
    sys.stdout = tee_stdout
    sys.stderr = tee_stdout
    
    print(f"")

def save_log_file(log_filename):
    """
    """
    with open(log_filename, 'w', encoding='utf-8') as f:
        f.write(log_string_io.getvalue())
    print(f"")

# ---
try:
    from google import genai
except ImportError:
    print("="*80)
    print("")
    print("")
    print("")
    print("")
    print("="*80)
    exit(1)

# ==============================================================================
#  (v14 - )
# ==============================================================================

def parse_time_v9(time_str):
    """
    (v9) 'HH:_MM'
    """
    ts = time_str.strip()
    # 
    if ":_" in ts:
        ts = ts.replace(':_', ':')
        logging.warning(f" 'HH:_MM' -> '{ts}'")

    if re.match(r'^\d+:\d+:\d+:\d+[,.]\d+$', ts):
        logging.warning(f""); return None
    ts = re.sub(r'(\d+:\d+:\d+)[.:](\d+)$', r'\1,\2', ts)
    ts = re.sub(r'^(\d+:\d+)[.,:](\d+)$', r'00:\1,\2', ts)
    match = re.match(r'^(\d+):(\d+):(\d+),(\d+)$', ts)
    if not match:
        logging.warning(f": {ts}"); return None
    try:
        h, m, s, ms = (int(g) for g in match.groups())
        if s >= 60: logging.warning(f" ({s}) 59'\n'"); s = 59
        if m >= 60: logging.warning(f" ({m}) 59'\n'"); m = 59
        return timedelta(hours=h, minutes=m, seconds=s, milliseconds=ms)
    except ValueError:
        logging.error(f": {ts}"); return None

def format_timedelta_v7(td):
    if not isinstance(td, timedelta): return "00:00:00,000"
    total_seconds = td.total_seconds()
    if total_seconds < 0: total_seconds = 0
    hours, remainder = divmod(total_seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    milliseconds = td.microseconds // 1000
    return f"{int(hours):02}:{int(minutes):02}:{int(seconds):02},{milliseconds:03}"

def format_srt_from_text_v14(srt_content, audio_filename, overlap_tolerance_td):
    """
    (V14 - )
    1.  
    """
    if srt_content.startswith("```srt"): srt_content = srt_content[srt_content.find('\n') + 1:]
    if srt_content.endswith("\n```"): srt_content = srt_content[:-4]
    srt_content = srt_content.strip().replace('\r\n', '\n')

    # ---
    # 
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
            # 
            start_td, end_td = parse_time_v9(start_raw), parse_time_v9(end_raw)
            
            is_unparsable = start_td is None or end_td is None
            is_overlap_violation = not is_unparsable and start_td < (last_correct_end_td - overlap_tolerance_td)

            if is_unparsable or is_overlap_violation:
                severe_correction_count += 1
                if is_unparsable:
                    logging.warning(f" {i+1} ': {time_line}' ( {severe_correction_count})")
                else:
                    logging.warning(f" {i+1} (> {overlap_tolerance_td.total_seconds()}s): {format_timedelta_v7(last_correct_end_td)}, {format_timedelta_v7(start_td)} ( {severe_correction_count})")
                start_td = last_correct_end_td + timedelta(milliseconds=100)
                end_td = start_td + SAFE_FALLBACK_DURATION
            
            if end_td <= start_td:
                logging.warning(f" {i+1} ")
                end_td = start_td + SAFE_FALLBACK_DURATION
            
            if (end_td - start_td) > MAX_DURATION:
                logging.warning(f" {i+1} ")
                end_td = start_td + SAFE_FALLBACK_DURATION
            
            last_correct_end_td = end_td
            corrected_blocks.append(f"{entry_counter}\n{format_timedelta_v7(start_td)} --> {format_timedelta_v7(end_td)}\n{full_text}")
            entry_counter += 1
        except Exception as e:
            logging.error(f" {i+1} '{block[:50]}...' - : {e}")

    if not corrected_blocks:
        logging.warning(f" {audio_filename} ")
        return "", severe_correction_count

    return "\n\n".join(corrected_blocks) + "\n\n", severe_correction_count

# ==============================================================================
# 
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
            logging.info(f"[{file_basename} |  {attempt+1}/{max_retries}] ")
            uploaded_file = client.files.upload(file=audio_path)
            logging.info(f". '{model_name}' ")
            response = client.models.generate_content(model=model_name, contents=[prompt_text, uploaded_file])
            if hasattr(response, 'usage_metadata') and response.usage_metadata:
                tokens_used = response.usage_metadata.total_token_count
            if uploaded_file:
                logging.info(f", '{uploaded_file.name}'...")
                client.files.delete(name=uploaded_file.name)
            
            if response.text:
                raw_srt_path = os.path.splitext(srt_path)[0] + ".raw.txt"
                with open(raw_srt_path, 'w', encoding='utf-8') as f: f.write(response.text)
                
                # 
                corrected_srt, severe_correction_count = format_srt_from_text_v14(response.text, file_basename, overlap_tolerance_td)
                
                if severe_correction_count > correction_threshold:
                    logging.warning(f"[{file_basename}]  {severe_correction_count} {correction_threshold}.")
                    logging.warning("")
                    raise ValueError(f"Correction threshold exceeded: {severe_correction_count} > {correction_threshold}")
                
                with open(srt_path, 'w', encoding='utf-8') as f: f.write(corrected_srt)
                logging.info(f"! : {os.path.basename(srt_path)}")
                return srt_path, tokens_used
            else:
                logging.error(f" '{file_basename}' ")
                return None, tokens_used

        except Exception as e:
            logging.warning(f" '{file_basename}' ( : {type(e).__name__}) : {e}")
            if uploaded_file:
                try: client.files.delete(name=uploaded_file.name)
                except Exception as del_e: logging.warning(f": {del_e}")
            if attempt < max_retries -1:
                wait_time = retry_delays[attempt]
                logging.info(f" {wait_time} ...")
                time.sleep(wait_time)
            else:
                logging.error(f" ({max_retries}), '{file_basename}' ")
                return None, tokens_used
    return None, tokens_used

def split_audio(input_file, temp_dir, chunk_duration_seconds=600):
    if not os.path.exists(temp_dir): os.makedirs(temp_dir)
    base_name = os.path.splitext(os.path.basename(input_file))[0]
    output_pattern = os.path.join(temp_dir, f"{base_name}_chunk_%03d.mp3")
    existing_chunks = sorted([os.path.join(temp_dir, f) for f in os.listdir(temp_dir) if f.startswith(f"{base_name}_chunk_") and f.endswith(".mp3")])
    if existing_chunks:
        logging.info(f" {len(existing_chunks)} ")
        return existing_chunks
    
    command = ['ffmpeg', '-i', input_file, '-vn', '-f', 'segment', '-segment_time', str(chunk_duration_seconds), '-acodec', 'libmp3lame', '-b:a', '192k', output_pattern]
    try:
        logging.info(f" FFmpeg : {os.path.basename(input_file)}")
        startupinfo = None
        if os.name == 'nt':
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE

        process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, startupinfo=startupinfo, creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0)
        stdout, stderr = process.communicate()

        if process.returncode != 0:
            logging.error(f"FFmpeg : {stderr.decode('utf-8', errors='ignore')}")
            return []
            
    except FileNotFoundError: 
        logging.error("FFmpeg. PATH , ffmpeg.exe ."); return []
    except Exception as e:
        logging.error(f" FFmpeg : {e}"); return []
        
    return sorted([os.path.join(temp_dir, f) for f in os.listdir(temp_dir) if f.startswith(f"{base_name}_chunk_") and f.endswith(".mp3")])


def merge_srts(srt_files, final_srt_path, chunk_duration_seconds):
    logging.info(f" {len(srt_files)} SRT  {os.path.basename(final_srt_path)}")
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
                logging.warning(f": {srt_file}. .")
                if i < len(sorted_srts) - 1: global_offset += chunk_duration_td
            except Exception as e:
                logging.error(f" SRT '{os.path.basename(srt_file)}' : {e}")
                if i < len(sorted_srts) - 1: global_offset += chunk_duration_td

# ==============================================================================
# 
# ==============================================================================

def create_manual_summary(log_content, summary_filename):
    keywords = [
        "", "", "", "Error", "Failed", "Warning", "Exception",
        "", "Retry", "", "Threshold", "", "Correction",
        "", "Unparsable", "", "Overlap", "",
        "", "token", " SRT"
    ]
    summary_lines = [line for line in log_content.splitlines() if any(keyword in line for keyword in keywords)]
    summary_text = "\n".join(summary_lines)
    with open(summary_filename, 'w', encoding='utf-8') as f:
        f.write("="*40 + "\n\n" + "="*40 + "\n\n" + summary_text)
    logging.info(f": {summary_filename}")

def create_ai_summary(log_content, summary_filename, client, model_name):
    logging.info(". ")
    if not client:
        logging.error(": .")
        return

    prompt = f"""
    , Python .
    , .

    :
    1.  ****:, .
    2.  ****:(chunk) .
        *   .
        *   ?
        *   ?
    3.  **** Python (Exception), FFmpeg  API , .
    4.  **** token.
    5.  **SRT** SRT .

    , Markdown .

    ---  ---
    {log_content}
    ---  ---
    """
    
    try:
        logging.info(f" '{model_name}' ...")
        response = client.models.generate_content(model=model_name, contents=[prompt])
        ai_summary = response.text
        with open(summary_filename, 'a', encoding='utf-8') as f:
            f.write("\n\n" + "="*40 + "\nAI\n" + "="*40 + "\n\n" + ai_summary)
        logging.info(f": {summary_filename}")
    except Exception as e:
        logging.error(f" AI : {e}")
        error_message = f"\n\n--- AI ---\n: {e}\n"
        with open(summary_filename, 'a', encoding='utf-8') as f:
            f.write(error_message)

# ==============================================================================
# 
# ==============================================================================

def main():
    start_time = time.time()
    parser = argparse.ArgumentParser(description=" (v11 - ).")
    parser.add_argument("input_file", help=".")
    parser.add_argument("--temp_dir", default=os.path.join(os.getcwd(), "temp"), help=".")
    parser.add_argument("--chunk_duration", type=int, default=600, help="().")
    parser.add_argument("--prompt_file", default="rule.md", help=".")
    parser.add_argument("--api_key", help=" ().")
    parser.add_argument("--model_name", default="models/gemini-2.5-pro", help="要使用的 Gemini 模型名稱。")
    parser.add_argument("--correction_threshold", type=int, default=5, help=", .")
    parser.add_argument("--overlap_tolerance", type=float, default=0.5, help=".")
    parser.add_argument("--enable_ai_summary", action='store_true', help=", API .")
    
    args = parser.parse_args()

    file_basename = os.path.splitext(os.path.basename(args.input_file))[0]
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_filename = f"{file_basename}__{timestamp}.txt"
    summary_filename = f"{file_basename}__{timestamp}.txt"
    setup_logging(log_filename)

    logging.info("="*50)
    logging.info(f": {args.input_file}")
    logging.info(f": transcribe_pro_v5_branch_04_branch_11")
    logging.info("="*50)

    api_key = args.api_key or os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
    client = None
    try:
        client = genai.Client(api_key=api_key) if api_key else genai.Client()
        logging.info(f" API . : {args.model_name}")
    except Exception as e:
        logging.error(f" API : {e}");
        if not api_key: logging.error(" GEMINI_API_KEY  GOOGLE_API_KEY , --api_key .")
        save_log_file(log_filename)
        return
    
    prompt_text = ""
    if os.path.exists(args.prompt_file):
        with open(args.prompt_file, 'r', encoding='utf-8') as f: prompt_text = f.read()
    else:
        logging.warning(f": {args.prompt_file}. .")

    if not os.path.exists(args.temp_dir): os.makedirs(args.temp_dir)
    chunk_mp3_files = split_audio(args.input_file, args.temp_dir, args.chunk_duration)
    
    if chunk_mp3_files:
        all_chunk_srts, total_tokens_used = [], 0
        for chunk_mp3_path in chunk_mp3_files:
            chunk_srt_path = os.path.splitext(chunk_mp3_path)[0] + ".srt"
            all_chunk_srts.append(chunk_srt_path)
            if os.path.exists(chunk_srt_path) and os.path.getsize(chunk_srt_path) > 0:
                logging.info(f" SRT, : {os.path.basename(chunk_srt_path)}")
            else:
                _, tokens = transcribe_audio(client, chunk_mp3_path, prompt_text, args.model_name, args.correction_threshold, args.overlap_tolerance)
                total_tokens_used += tokens if tokens else 0

        logging.info(f".  {total_tokens_used} token.")
        final_srt_path = f"{file_basename}.srt"
        merge_srts(all_chunk_srts, final_srt_path, args.chunk_duration)
        logging.info(f". SRT : {final_srt_path}")

    end_time = time.time()
    logging.info(f": {end_time - start_time:.2f} .")
    
    log_content = log_string_io.getvalue()
    save_log_file(log_filename)
    create_manual_summary(log_content, summary_filename)
    
    if args.enable_ai_summary:
        create_ai_summary(log_content, summary_filename, client, args.model_name)

if __name__ == "__main__":
    original_stdout = sys.stdout
    original_stderr = sys.stderr
    try:
        main()
    except Exception as e:
        logging.error(f": {e}", exc_info=True)
    finally:
        if hasattr(sys.stdout, 'files'):
             sys.stdout = original_stdout
        if hasattr(sys.stderr, 'files'):
             sys.stderr = original_stderr
        print("\n.")
