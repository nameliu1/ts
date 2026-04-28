import os
import time
import datetime
import sys
import psutil
import subprocess
import shutil
import pandas as pd

# 配置信息
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
URL_FILE = os.path.join(BASE_DIR, "url.txt")
DIR_FILE = os.path.join(BASE_DIR, "dirv2.txt")
JSON_FILE = os.path.join(BASE_DIR, "res.json")        # spray原始输出
EXCEL_FILE = os.path.join(BASE_DIR, "res_processed.xlsx")  # 处理后的Excel
TXT_FILE = os.path.join(BASE_DIR, "res_processed.txt")    # 提取的URL列表
# Default to showing the console so double-click or terminal runs do not
# look like an immediate crash when the script hides its own window.
HIDE_PYTHON_CONSOLE = False
MONITOR_INTERVAL = 5  # 进程监控间隔（秒）
STATUS_CODE_COL_INDEX = 9  # J列（Excel列索引从0开始，J列对应索引9）【Spray状态码结果在此列】
URL_COL_INDEX = 4  # E列（根据实际Excel列调整）
EHOLE_QUICK_TIMEOUT = 3  # ehole快速完成的超时时间（秒）

# 需要删除的过程文件列表
TO_DELETE_FILES = [
    os.path.join(BASE_DIR, "url.txt.stat"),
    os.path.join(BASE_DIR, "res_processed.txt")
]

def log(message):
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp}] {message}")

def hide_python_console():
    if HIDE_PYTHON_CONSOLE:
        try:
            import win32gui, win32con
            hwnd = win32gui.GetForegroundWindow()
            win32gui.ShowWindow(hwnd, win32con.SW_HIDE)
        except:
            log("警告: 无法隐藏 Python 控制台窗口")

def run_native_command(command, process_name):
    log(f"执行命令: {command}")
    # 仅针对ehole添加2秒延迟，避免进程快速结束导致监控失败
    if process_name.lower() == "ehole.exe":
        os.system(f'start cmd /c "{command} & ping -n 2 127.0.0.1 >nul"')
    else:
        os.system(f'start cmd /c "{command}"')

def monitor_process(process_name, timeout=3600):
    log(f"监控进程: {process_name}")
    start_time = time.time()
    
    # 特殊处理ehole进程的快速完成情况
    is_ehole = process_name.lower() == "ehole.exe"
    quick_timeout = EHOLE_QUICK_TIMEOUT
    
    # 等待进程启动
    while time.time() - start_time < timeout:
        if any(proc.name().lower() == process_name.lower() for proc in psutil.process_iter()):
            log(f"进程已启动: {process_name}")
            break
            
        # 检查是否是ehole并且已经超过快速超时时间
        if is_ehole and (time.time() - start_time > quick_timeout):
            log(f"警告: ehole在{quick_timeout}秒内未启动，可能已快速完成")
            return True
            
        time.sleep(1)
    else:
        log(f"错误: 等待 {process_name} 启动超时")
        return False
    
    start_time = time.time()
    # 等待进程结束
    while time.time() - start_time < timeout:
        if not any(proc.name().lower() == process_name.lower() for proc in psutil.process_iter()):
            log(f"进程已结束: {process_name}")
            return True
        time.sleep(1)
    log(f"错误: {process_name} 运行超时")
    return False

def wait_for_file(file_path, timeout=300):
    log(f"等待文件生成: {file_path}")
    start_time = time.time()
    while time.time() - start_time < timeout:
        if os.path.exists(file_path):
            log(f"文件已生成: {file_path}")
            return True
        time.sleep(1)
    log(f"错误: 文件未生成: {file_path}")
    return False


def get_config_output_path():
    try:
        config_path = os.path.join(BASE_DIR, "config.yaml")
        with open(config_path, "r", encoding="utf-8") as f:
            for line in f:
                if line.startswith("CfgOutPath:"):
                    path = line.split(":", 1)[1].strip().strip('"')
                    if path:
                        return path
    except Exception as e:
        log(f"警告: 读取 config.yaml 输出目录失败: {e}")
    return None


def wait_for_ehole_file(expected_path, timeout=300):
    config_out_path = get_config_output_path()
    candidate_paths = [expected_path]

    if config_out_path:
        config_candidate = os.path.join(config_out_path, os.path.basename(expected_path))
        if config_candidate not in candidate_paths:
            candidate_paths.append(config_candidate)

    log(f"等待ehole结果文件生成: {' | '.join(candidate_paths)}")
    start_time = time.time()

    while time.time() - start_time < timeout:
        for candidate_path in candidate_paths:
            if not os.path.exists(candidate_path):
                continue
            if os.path.getsize(candidate_path) <= 0:
                continue

            if candidate_path != expected_path:
                try:
                    shutil.move(candidate_path, expected_path)
                    log(f"已将ehole结果移动到目标目录: {expected_path}")
                    return expected_path
                except Exception as e:
                    log(f"警告: 无法移动ehole结果文件 {candidate_path} -> {expected_path}: {e}")
                    return candidate_path

            log(f"ehole结果文件已生成: {candidate_path}")
            return candidate_path

        time.sleep(1)

    log(f"错误: 未找到ehole结果文件: {' | '.join(candidate_paths)}")
    return None

# 生成不冲突的文件名
def generate_unique_filename(base_dir, base_name, ext):
    counter = 1
    original_name = f"{base_name}{ext}"
    full_path = os.path.join(base_dir, original_name)
    
    # 如果文件已存在，则添加序号后缀
    while os.path.exists(full_path):
        new_name = f"{base_name}_{counter}{ext}"
        full_path = os.path.join(base_dir, new_name)
        counter += 1
    
    return full_path

# 删除指定的过程文件
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
        bufsize=1
    )
    last_log_time = time.time()
    if process.stdout is not None:
        for line in process.stdout:
            line = line.rstrip()
            if line:
                log(f"[process_data] {line}")
            last_log_time = time.time()
    return_code = process.wait()
    if return_code != 0:
        log(f"错误: 数据处理失败，退出码: {return_code}")
        return False
    if not os.path.exists(excel_file):
        log(f"错误: 处理后的Excel文件未生成: {excel_file}")
        return False
    if not os.path.exists(txt_file):
        log(f"警告: 未找到URL列表文件: {txt_file}，可能没有有效URL")
        return False
    with open(txt_file, 'r', encoding='utf-8') as f:
        url_count = len(f.readlines())
    log(f"成功提取 {url_count} 个URL")
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
        
        # 标记状态码列位置（J列）
        log(f"注意: Spray扫描的状态码结果配置为J列，对应Python索引 {STATUS_CODE_COL_INDEX}")
        
        try:
            status_code_col = df.columns[STATUS_CODE_COL_INDEX]
            url_col = df.columns[URL_COL_INDEX]
        except IndexError:
            log(f"错误: Excel文件列数不足，无法获取索引为 {STATUS_CODE_COL_INDEX} (J列) 或 {URL_COL_INDEX} 的列")
            log(f"Excel实际列数: {len(df.columns)}，列名: {list(df.columns)}")
            return None
        
        log(f"使用列 '{url_col}' (E列) 作为URL列，列 '{status_code_col}' (J列) 作为状态码列")
        
        if df[status_code_col].dtype not in [int, float]:
            log(f"警告: 状态码列数据类型不是数值类型: {df[status_code_col].dtype}")
            log(f"尝试转换数据类型...")
            try:
                df[status_code_col] = pd.to_numeric(df[status_code_col], errors='coerce')
            except:
                log(f"错误: 无法将状态码列转换为数值类型")
                return None
        
        df_200 = df[df[status_code_col] == 200].copy()
        total_rows = len(df)
        filtered_rows = len(df_200)
        log(f"Excel总行数: {total_rows}，状态码为200的行数: {filtered_rows}")
        
        if filtered_rows == 0:
            log("警告: 未找到状态码为200的URL")
            return None
        
        urls_200 = df_200[df.columns[URL_COL_INDEX]].dropna().unique().tolist()
        log(f"提取并去重后得到 {len(urls_200)} 个状态码为200的URL")
        
        date_str = datetime.datetime.now().strftime("%Y%m%d")
        base_filename = f"{date_str}_status200_urls_{count}"
        
        # 使用新函数生成唯一文件名
        output_file = generate_unique_filename(output_dir, base_filename, ".txt")
        
        log(f"将状态码为200的URL写入文件: {output_file}")
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write('\n'.join(urls_200))
        
        with open(output_file, 'r', encoding='utf-8') as f:
            written_urls = f.read().splitlines()
        
        if len(written_urls) != len(urls_200):
            log(f"警告: 写入的URL数量({len(written_urls)})与筛选的URL数量({len(urls_200)})不一致")
        
        log(f"状态码为200的URL已保存至: {output_file}")
        return output_file
    except Exception as e:
        log(f"筛选错误: {e}")
        return None

def main():
    try:
        hide_python_console()
        log(f"开始自动化漏洞扫描和指纹识别流程")
        log(f"基础目录: {BASE_DIR}")
        
        # 创建日期文件夹
        date_folder = datetime.datetime.now().strftime("%m%d")
        full_date_dir = os.path.join(BASE_DIR, date_folder)
        os.makedirs(full_date_dir, exist_ok=True)
        log(f"创建日期文件夹: {full_date_dir}")
        
        # 清理指定的过程文件
        clean_process_files()
        
        # 步骤1: 执行spray扫描
        log("步骤1: 执行spray扫描...")
        spray_cmd = f'spray.exe -l "{URL_FILE}" -d "{DIR_FILE}" -f "{JSON_FILE}"'
        run_native_command(spray_cmd, "spray.exe")
        if not monitor_process("spray.exe", timeout=1800):
            log("错误: spray执行失败或超时")
            sys.exit(1)
        if not wait_for_file(JSON_FILE):
            log("错误: spray未生成结果文件")
            sys.exit(1)
        
        # 步骤2: 处理spray结果，提取有效URL
        log("步骤2: 处理spray结果，提取有效URL...")
        
        # 为输出文件生成唯一文件名
        unique_excel_file = generate_unique_filename(BASE_DIR, "res_processed", ".xlsx")
        unique_txt_file = generate_unique_filename(BASE_DIR, "res_processed", ".txt")
        
        if not process_spray_output(JSON_FILE, unique_excel_file, unique_txt_file):
            log("错误: 处理spray输出失败")
            sys.exit(1)
        
        # 步骤3: 筛选状态码200的URL
        log("步骤3: 筛选状态码200的URL...")
        filtered_txt_path = filter_status_200(unique_excel_file, full_date_dir, 1)
        if not filtered_txt_path:
            log("错误: 未生成状态码为200的URL文件")
            sys.exit(1)
        
        # 步骤3.5: 移动Spray结果文件到日期文件夹
        log("步骤3.5: 移动Spray结果文件到日期文件夹...")
        
        # 为移动的文件生成唯一文件名
        spray_json_base = f"spray_original_{datetime.datetime.now().strftime('%Y%m%d')}"
        spray_json_dest = generate_unique_filename(full_date_dir, spray_json_base, ".json")
        
        spray_excel_base = f"spray_processed_{datetime.datetime.now().strftime('%Y%m%d')}"
        spray_excel_dest = generate_unique_filename(full_date_dir, spray_excel_base, ".xlsx")
        
        shutil.move(JSON_FILE, spray_json_dest)
        log(f"已移动Spray原始结果: {spray_json_dest}")
        
        shutil.move(unique_excel_file, spray_excel_dest)
        log(f"已移动Spray处理后Excel: {spray_excel_dest}")
        
        # 步骤4: 执行ehole指纹识别
        log("步骤4: 执行ehole指纹识别...")
        
        # 为ehole结果生成唯一文件名
        ehole_base = f"ehole_result_{datetime.datetime.now().strftime('%Y%m%d')}"
        ehole_output = generate_unique_filename(full_date_dir, ehole_base, ".xlsx")
        
        ehole_cmd = f'ehole finger -l "{filtered_txt_path}" -o "{ehole_output}" -t 10'
        run_native_command(ehole_cmd, "ehole.exe")
        
        # 监控ehole进程并等待结束
        if not monitor_process("ehole.exe", timeout=1800):
            log("错误: ehole执行失败或超时")
            # 即使监控失败，也继续检查文件是否存在
            pass
        
        # 检查ehole结果文件是否生成
        actual_ehole_output = wait_for_ehole_file(ehole_output)
        if not actual_ehole_output:
            log("错误: ehole未生成结果文件")
            sys.exit(1)

        # 美化ehole结果表格
        log("美化ehole结果表格...")
        subprocess.run(["python", "process_data.py", actual_ehole_output, actual_ehole_output])
        log("ehole结果表格美化完成")
        
        log(f"自动化流程全部完成！所有结果保存在: {full_date_dir}")
    
    except Exception as e:
        log(f"程序异常: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    os.system("chcp 65001 >nul 2>&1")  # 确保中文显示正常
    
    # 检查依赖
    try:
        import psutil
        import pandas as pd
    except ImportError:
        log("错误: 缺少psutil或pandas库，请执行 'pip install psutil pandas'")
        sys.exit(1)
    
    main()
