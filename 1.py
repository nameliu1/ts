import datetime
import glob
import os
import shutil
import subprocess
import sys
import time
from contextlib import redirect_stderr, redirect_stdout

import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
URL_FILE = os.path.join(BASE_DIR, "url.txt")
DIR_FILE = os.path.join(BASE_DIR, "dirv2.txt")
JSON_FILE = os.path.join(BASE_DIR, "res.json")
EXCEL_FILE = os.path.join(BASE_DIR, "res_processed.xlsx")
TXT_FILE = os.path.join(BASE_DIR, "res_processed.txt")
HIDE_PYTHON_CONSOLE = False
COMMAND_TIMEOUT = 86400  # 24小时，覆盖最长扫描场景
EHOLE_WAIT_TIMEOUT = 600  # ehole 文件等待增加到 10 分钟
STATUS_CODE_COL_INDEX = 9
URL_COL_INDEX = 4
URL_COLUMN_CANDIDATES = ["url", "direct url", "directurl", "网址", "链接"]
STATUS_COLUMN_CANDIDATES = ["status", "status code", "status_code", "code", "状态码", "响应码", "http code"]
TO_DELETE_FILES = [
    os.path.join(BASE_DIR, "url.txt.stat"),
    os.path.join(BASE_DIR, "res_processed.txt"),
]
LOG_FILE_PATH = None


class TeeStream:
    def __init__(self, *streams):
        self.streams = streams

    def write(self, data):
        for stream in self.streams:
            stream.write(data)
        return len(data)

    def flush(self):
        for stream in self.streams:
            stream.flush()


class LoggerWriter:
    def __init__(self, logger):
        self.logger = logger
        self.buffer = ""

    def write(self, data):
        if not data:
            return 0
        self.buffer += data
        while "\n" in self.buffer:
            line, self.buffer = self.buffer.split("\n", 1)
            line = line.rstrip("\r")
            if line:
                self.logger(line)
        return len(data)

    def flush(self):
        if self.buffer:
            line = self.buffer.rstrip("\r")
            if line:
                self.logger(line)
            self.buffer = ""


def log(message):
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{timestamp}] {message}"
    print(line)


def hide_python_console():
    if HIDE_PYTHON_CONSOLE:
        try:
            import win32con
            import win32gui

            hwnd = win32gui.GetForegroundWindow()
            win32gui.ShowWindow(hwnd, win32con.SW_HIDE)
        except Exception:
            log("警告: 无法隐藏 Python 控制台窗口")


def run_command(args, stage_name, timeout=COMMAND_TIMEOUT):
    log(f"执行{stage_name}命令: {' '.join(args)}")
    proc = subprocess.Popen(
        args,
        shell=False,
        cwd=BASE_DIR,
        stdin=subprocess.DEVNULL,
        stdout=None,
        stderr=None,
    )
    start_time = time.time()

    try:
        while True:
            return_code = proc.poll()
            if return_code is not None:
                elapsed = time.time() - start_time
                log(f"{stage_name} 执行结束，耗时 {elapsed:.1f} 秒，退出码: {return_code}")
                return return_code == 0

            if time.time() - start_time > timeout:
                log(f"错误: {stage_name} 执行超时 ({timeout}秒)，正在强制终止进程树...")
                terminate_process_tree(proc, stage_name)
                return False

            time.sleep(1)

    except Exception as e:
        log(f"{stage_name} 执行异常: {e}")
        terminate_process_tree(proc, stage_name)
        return False


def terminate_process_tree(proc, stage_name):
    try:
        import psutil
        parent = psutil.Process(proc.pid)
        children = parent.children(recursive=True)
        for child in children:
            child.terminate()
        parent.terminate()
        gone, alive = psutil.wait_procs([parent, *children], timeout=5)
        for process in alive:
            process.kill()
        log(f"{stage_name} 进程树已终止，正常退出 {len(gone)} 个，强制结束 {len(alive)} 个")
    except ImportError:
        log("psutil 未安装，只能终止主进程")
        proc.kill()
    except Exception as e:
        log(f"终止进程树时出错: {e}")
        proc.kill()


def wait_for_file(file_path, timeout=300):
    log(f"等待文件生成: {file_path}")
    start_time = time.time()
    last_size = -1
    stable_seconds = 0
    while time.time() - start_time < timeout:
        if os.path.exists(file_path):
            current_size = os.path.getsize(file_path)
            if current_size > 0:
                if current_size == last_size:
                    stable_seconds += 1
                    if stable_seconds >= 3:
                        log(f"文件已稳定生成: {file_path} ({current_size} 字节)")
                        return True
                else:
                    stable_seconds = 0
                    log(f"文件大小变化中: {last_size} -> {current_size} 字节")
                last_size = current_size
        time.sleep(1)
    log(f"错误: 文件超时未生成: {file_path}")
    return False


def describe_file(file_path):
    if os.path.exists(file_path):
        return f"存在，大小 {os.path.getsize(file_path)} 字节"
    return "不存在"


def log_input_diagnostics(stage_name, executable_path):
    log(f"{stage_name} 可执行文件: {executable_path}")
    log(f"url.txt: {describe_file(URL_FILE)}")
    log(f"dirv2.txt: {describe_file(DIR_FILE)}")
    log(f"res.json: {describe_file(JSON_FILE)}")


def get_config_output_path():
    try:
        config_path = os.path.join(BASE_DIR, "config.yaml")
        with open(config_path, "r", encoding="utf-8") as f:
            for line in f:
                if line.startswith("CfgOutPath:"):
                    path = line.split(":", 1)[1].strip().strip('"')
                    if path:
                        normalized = os.path.normpath(path)
                        log(f"读取到CfgOutPath: {normalized}")
                        return normalized
    except Exception as e:
        log(f"警告: 读取 config.yaml 输出目录失败: {e}")
    return None


def unique_paths(paths):
    seen = set()
    unique = []
    for path in paths:
        if not path:
            continue
        normalized = os.path.normcase(os.path.normpath(path))
        if normalized in seen:
            continue
        seen.add(normalized)
        unique.append(path)
    return unique


def quote_path(path):
    return f'"{path}"'


def resolve_local_executable(*names):
    for name in names:
        candidate = os.path.join(BASE_DIR, name)
        if os.path.isfile(candidate):
            return candidate

    for name in names:
        resolved = shutil.which(name)
        if resolved:
            return resolved

    return None


def resolve_spray_executable():
    return resolve_local_executable("spray.exe", "spray")


def resolve_ehole_executable():
    return resolve_local_executable("ehole.exe", "ehole")


def read_valid_urls(file_path):
    urls = []
    seen = set()
    with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
        for line in f:
            url = line.strip()
            if not url.startswith(("http://", "https://")):
                continue
            if url in seen:
                continue
            seen.add(url)
            urls.append(url)
    return urls


def validate_ehole_input(file_path):
    if not os.path.exists(file_path):
        log(f"错误: ehole输入文件不存在: {file_path}")
        return []

    urls = read_valid_urls(file_path)
    if not urls:
        log(f"错误: ehole输入为空或无有效URL: {file_path}")
        return []

    log(f"ehole输入URL数量: {len(urls)}")
    return urls


def find_semantic_column(columns, candidates):
    normalized = {str(column).strip().lower(): column for column in columns}
    for candidate in candidates:
        match = normalized.get(candidate.lower())
        if match is not None:
            return match
    return None


def detect_spray_columns(df):
    url_col = df.columns[URL_COL_INDEX] if len(df.columns) > URL_COL_INDEX else None
    status_col = df.columns[STATUS_CODE_COL_INDEX] if len(df.columns) > STATUS_CODE_COL_INDEX else None
    semantic_url_col = find_semantic_column(df.columns, URL_COLUMN_CANDIDATES)
    semantic_status_col = find_semantic_column(df.columns, STATUS_COLUMN_CANDIDATES)

    return {
        "url_col": url_col,
        "status_col": status_col,
        "semantic_url_col": semantic_url_col,
        "semantic_status_col": semantic_status_col,
    }


def normalize_status_column(series):
    return pd.to_numeric(series, errors="coerce")


def normalize_url_values(values):
    urls = []
    seen = set()
    for value in values:
        if pd.isna(value):
            continue
        url = str(value).strip()
        if not url.startswith(("http://", "https://")):
            continue
        if url in seen:
            continue
        seen.add(url)
        urls.append(url)
    return urls


def extract_status_200_urls(df, status_col, url_col):
    if status_col is None or url_col is None:
        return []

    status_series = normalize_status_column(df[status_col])
    filtered = df[status_series == 200]
    return normalize_url_values(filtered[url_col].tolist())


def file_size_stable(file_path, settle_seconds=1):
    try:
        if not os.path.isfile(file_path):
            return False
        size_before = os.path.getsize(file_path)
        if size_before <= 0:
            return False
        time.sleep(settle_seconds)
        if not os.path.isfile(file_path):
            return False
        size_after = os.path.getsize(file_path)
        return size_after > 0 and size_before == size_after
    except OSError:
        return False


def discover_ehole_output(expected_path):
    config_out_path = get_config_output_path()
    expected_dir = os.path.dirname(expected_path)
    expected_name = os.path.basename(expected_path)
    expected_stem, expected_ext = os.path.splitext(expected_name)

    exact_candidates = unique_paths([
        expected_path,
        os.path.join(config_out_path, expected_name) if config_out_path else None,
    ])
    search_dirs = unique_paths([expected_dir, config_out_path, BASE_DIR])
    patterns = [expected_name, f"{expected_stem}*.xlsx"]

    for candidate in exact_candidates:
        if file_size_stable(candidate):
            return candidate, exact_candidates, search_dirs

    discovered = []
    for search_dir in search_dirs:
        if not search_dir or not os.path.isdir(search_dir):
            continue
        for pattern in patterns:
            discovered.extend(glob.glob(os.path.join(search_dir, pattern)))

    discovered = unique_paths(discovered)
    discovered = [path for path in discovered if os.path.isfile(path)]
    discovered.sort(key=lambda path: os.path.getmtime(path), reverse=True)

    for candidate in discovered:
        if file_size_stable(candidate):
            if expected_ext.lower() == ".xlsx" and not candidate.lower().endswith(".xlsx"):
                continue
            return candidate, exact_candidates, search_dirs

    return None, exact_candidates, search_dirs


def wait_for_ehole_file(expected_path, timeout=EHOLE_WAIT_TIMEOUT):
    start_time = time.time()
    last_logged_candidates = None

    while time.time() - start_time < timeout:
        candidate_path, exact_candidates, search_dirs = discover_ehole_output(expected_path)
        log_candidates = exact_candidates + search_dirs
        if log_candidates != last_logged_candidates:
            log(f"等待ehole结果文件生成: {' | '.join(log_candidates)}")
            last_logged_candidates = log_candidates

        if candidate_path:
            normalized_expected = os.path.normcase(os.path.normpath(expected_path))
            normalized_candidate = os.path.normcase(os.path.normpath(candidate_path))
            if normalized_candidate != normalized_expected:
                try:
                    shutil.move(candidate_path, expected_path)
                    log(f"已将ehole结果移动到目标目录: {expected_path}")
                    return expected_path
                except Exception as e:
                    log(f"警告: 无法移动ehole结果文件 {candidate_path} -> {expected_path}: {e}")
                    return candidate_path

            log(f"ehole结果文件已生成: {candidate_path}")
            return candidate_path

        time.sleep(2)

    log("错误: 在候选目录中未找到ehole结果文件")
    return None


def generate_unique_filename(base_dir, base_name, ext):
    counter = 1
    original_name = f"{base_name}{ext}"
    full_path = os.path.join(base_dir, original_name)

    while os.path.exists(full_path):
        new_name = f"{base_name}_{counter}{ext}"
        full_path = os.path.join(base_dir, new_name)
        counter += 1

    return full_path


def initialize_logging(output_dir):
    global LOG_FILE_PATH
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    LOG_FILE_PATH = generate_unique_filename(output_dir, f"run_{timestamp}", ".log")
    return LOG_FILE_PATH


def clean_process_files():
    log("开始清理上次运行的过程文件...")
    for file_path in TO_DELETE_FILES:
        if os.path.exists(file_path):
            try:
                os.remove(file_path)
                log(f"已删除: {file_path}")
            except Exception as e:
                log(f"删除文件 {file_path} 时出错: {e}")
        else:
            log(f"文件不存在，跳过删除: {file_path}")
    log("过程文件清理完成")


def process_spray_output(json_file, excel_file, txt_file):
    log(f"开始处理spray结果: {json_file}")
    log("process_data.py 处理大文件时可能持续较久，期间会显示实时输出。")
    process = subprocess.Popen(
        ["python", "process_data.py", json_file, excel_file],
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        bufsize=1,
    )
    if process.stdout is not None:
        for line in process.stdout:
            line = line.rstrip()
            if line:
                log(f"[process_data] {line}")
    return_code = process.wait()
    if return_code != 0:
        log(f"错误: 数据处理失败，退出码: {return_code}")
        return False
    if not os.path.exists(excel_file):
        log(f"错误: 处理后的Excel文件未生成: {excel_file}")
        return False
    if not os.path.exists(txt_file):
        log(f"错误: 未找到URL列表文件: {txt_file}")
        return False

    urls = read_valid_urls(txt_file)
    if not urls:
        log(f"错误: 处理后的URL列表为空或无有效URL: {txt_file}")
        return False

    log(f"成功提取 {len(urls)} 个URL")
    return True


def filter_status_200(excel_file, output_dir, count):
    try:
        log(f"开始从 {excel_file} 中筛选状态码为200的URL...")
        if not os.path.exists(excel_file):
            log(f"错误: Excel文件不存在: {excel_file}")
            return None

        df = pd.read_excel(excel_file)
        if df.empty:
            log("错误: Excel文件为空")
            return None

        columns_info = detect_spray_columns(df)
        url_col = columns_info["url_col"]
        status_col = columns_info["status_col"]
        semantic_url_col = columns_info["semantic_url_col"]
        semantic_status_col = columns_info["semantic_status_col"]

        if status_col is None or url_col is None:
            log(f"错误: 无法识别状态码列或URL列，列名: {list(df.columns)}")
            return None

        log(f"优先使用列 '{url_col}' 作为URL列，列 '{status_col}' 作为状态码列")
        urls_200 = extract_status_200_urls(df, status_col, url_col)

        if not urls_200 and (semantic_url_col != url_col or semantic_status_col != status_col):
            log("固定列未提取到有效URL，尝试按列名语义重新识别...")
            urls_200 = extract_status_200_urls(df, semantic_status_col, semantic_url_col)
            if urls_200:
                url_col = semantic_url_col
                status_col = semantic_status_col
                log(f"回退后使用列 '{url_col}' 作为URL列，列 '{status_col}' 作为状态码列")

        filtered_rows = len(urls_200)
        log(f"提取并去重后得到 {filtered_rows} 个状态码为200的URL")

        if filtered_rows == 0:
            log("警告: 未找到状态码为200的有效URL")
            return None

        date_str = datetime.datetime.now().strftime("%Y%m%d")
        base_filename = f"{date_str}_status200_urls_{count}"
        output_file = generate_unique_filename(output_dir, base_filename, ".txt")

        log(f"将状态码为200的URL写入文件: {output_file}")
        with open(output_file, "w", encoding="utf-8") as f:
            f.write("\n".join(urls_200))

        written_urls = read_valid_urls(output_file)
        if len(written_urls) != len(urls_200):
            log(f"警告: 写入的有效URL数量({len(written_urls)})与筛选的URL数量({len(urls_200)})不一致")

        if not written_urls:
            log("错误: 最终输出的ehole输入URL为空")
            return None

        log(f"状态码为200的URL已保存至: {output_file}")
        return output_file
    except Exception as e:
        log(f"筛选错误: {e}")
        return None


def run_pipeline():
    hide_python_console()
    log("开始自动化漏洞扫描和指纹识别流程")
    log(f"基础目录: {BASE_DIR}")

    date_folder = datetime.datetime.now().strftime("%m%d")
    full_date_dir = os.path.join(BASE_DIR, date_folder)
    os.makedirs(full_date_dir, exist_ok=True)
    log(f"创建日期文件夹: {full_date_dir}")

    clean_process_files()

    log("步骤1: 执行spray扫描...")
    spray_executable = resolve_spray_executable()
    if not spray_executable:
        log("错误: 未找到spray可执行文件，请确认仓库内存在spray.exe或系统PATH可访问spray")
        return 1
    log_input_diagnostics("spray", spray_executable)
    spray_cmd = [spray_executable, "-l", "url.txt", "-d", "dirv2.txt", "-f", "res.json"]
    if not run_command(spray_cmd, "spray"):
        log("错误: spray执行失败")
        return 1
    if not wait_for_file(JSON_FILE, timeout=10):
        log("错误: spray退出成功但未生成res.json，请优先检查上方的spray可执行文件路径和输入文件大小")
        log_input_diagnostics("spray", spray_executable)
        return 1

    log("步骤2: 处理spray结果，提取有效URL...")
    unique_excel_file = generate_unique_filename(BASE_DIR, "res_processed", ".xlsx")
    unique_txt_file = generate_unique_filename(BASE_DIR, "res_processed", ".txt")
    if not process_spray_output(JSON_FILE, unique_excel_file, unique_txt_file):
        log("错误: 处理spray输出失败")
        return 1

    log("步骤3: 筛选状态码200的URL...")
    filtered_txt_path = filter_status_200(unique_excel_file, full_date_dir, 1)
    if not filtered_txt_path:
        log("错误: 未生成状态码为200的URL文件")
        return 1

    log("步骤3.5: 移动Spray结果文件到日期文件夹...")
    spray_json_base = f"spray_original_{datetime.datetime.now().strftime('%Y%m%d')}"
    spray_json_dest = generate_unique_filename(full_date_dir, spray_json_base, ".json")
    spray_excel_base = f"spray_processed_{datetime.datetime.now().strftime('%Y%m%d')}"
    spray_excel_dest = generate_unique_filename(full_date_dir, spray_excel_base, ".xlsx")
    shutil.move(JSON_FILE, spray_json_dest)
    log(f"已移动Spray原始结果: {spray_json_dest}")
    shutil.move(unique_excel_file, spray_excel_dest)
    log(f"已移动Spray处理后Excel: {spray_excel_dest}")

    log("步骤4: 执行ehole指纹识别...")
    ehole_urls = validate_ehole_input(filtered_txt_path)
    if not ehole_urls:
        return 1

    ehole_executable = resolve_ehole_executable()
    if not ehole_executable:
        log("错误: 未找到ehole可执行文件，请确认仓库内存在ehole.exe或系统PATH可访问ehole")
        return 1

    ehole_base = f"ehole_result_{datetime.datetime.now().strftime('%Y%m%d')}"
    ehole_output = generate_unique_filename(full_date_dir, ehole_base, ".xlsx")
    config_out_path = get_config_output_path()
    log(f"ehole可执行文件: {ehole_executable}")
    log(f"ehole目标输出路径: {ehole_output}")
    if config_out_path:
        log(f"ehole配置输出目录: {config_out_path}")

    ehole_cmd = [
        ehole_executable,
        "finger",
        "-l",
        filtered_txt_path,
        "-o",
        ehole_output,
        "-t",
        "10",
    ]
    if not run_command(ehole_cmd, "ehole"):
        log(f"错误: ehole执行失败，输入文件: {filtered_txt_path}，URL数量: {len(ehole_urls)}")
        return 1

    actual_ehole_output = wait_for_ehole_file(ehole_output)
    if not actual_ehole_output:
        log(f"错误: ehole未生成结果文件，输入文件: {filtered_txt_path}，URL数量: {len(ehole_urls)}")
        return 1

    if not actual_ehole_output.lower().endswith((".xlsx", ".xls")):
        log(f"错误: ehole产出文件不是Excel工作簿: {actual_ehole_output}")
        return 1
    if os.path.getsize(actual_ehole_output) <= 0:
        log(f"错误: ehole产出文件为空: {actual_ehole_output}")
        return 1

    log("美化ehole结果表格...")
    process_logger = LoggerWriter(lambda line: log(f"[process_data] {line}"))
    with redirect_stdout(process_logger), redirect_stderr(process_logger):
        import process_data as process_data_module
        result_code = process_data_module.process_data(actual_ehole_output, actual_ehole_output)
    process_logger.flush()
    if result_code != 0:
        log(f"错误: ehole结果后处理失败，文件可能不是可读工作簿: {actual_ehole_output}")
        return 1
    log("ehole结果表格美化完成")

    log(f"自动化流程全部完成！所有结果保存在: {full_date_dir}")
    return 0


def main():
    os.system("chcp 65001 >nul 2>&1")
    try:
        import pandas as pd  # noqa: F401
    except ImportError:
        print("错误: 缺少pandas库，请执行 'pip install pandas'")
        return 1

    date_folder = datetime.datetime.now().strftime("%m%d")
    full_date_dir = os.path.join(BASE_DIR, date_folder)
    os.makedirs(full_date_dir, exist_ok=True)
    log_file_path = initialize_logging(full_date_dir)

    original_stdout = sys.stdout
    original_stderr = sys.stderr
    with open(log_file_path, "a", encoding="utf-8", buffering=1) as log_file:
        tee_stdout = TeeStream(original_stdout, log_file)
        tee_stderr = TeeStream(original_stderr, log_file)
        with redirect_stdout(tee_stdout), redirect_stderr(tee_stderr):
            log(f"日志文件: {log_file_path}")
            try:
                return run_pipeline()
            except Exception as e:
                log(f"程序异常: {str(e)}")
                return 1


if __name__ == "__main__":
    sys.exit(main())
