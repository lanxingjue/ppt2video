import logging
import shutil
from pathlib import Path
import time
import configparser # 导入
import argparse  # 导入命令行参数解析器
import subprocess # 用于依赖检查
import shlex     # 用于依赖检查

# --- 配置解析 ---
config = configparser.ConfigParser()
config_path = Path(__file__).parent / 'config.ini'
if config_path.exists():
    try: config.read(config_path, encoding='utf-8')
    except Exception: logging.exception(f"加载配置文件 {config_path} 时出错。") # 使用 exception 记录完整错误
# 在日志配置前读取，以便设置日志级别
log_level_str = config.get('General', 'logging_level', fallback='INFO').upper()
log_level = getattr(logging, log_level_str, logging.INFO)

# --- 日志记录配置 ---
logging.basicConfig(
    level=log_level,
    format='%(asctime)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logging.info(f"日志级别设置为: {log_level_str}")


# --- 导入我们的模块 (放在日志配置后) ---
try:
    from ppt_processor import process_presentation
    logging.info("Successfully imported 'process_presentation' from ppt_processor.")
except Exception: logging.exception("导入 'process_presentation' 失败。"); exit()
try:
    from video_synthesizer import create_video_from_data
    logging.info("Successfully imported 'create_video_from_data' from video_synthesizer.")
except Exception: logging.exception("导入 'create_video_from_data' 失败。"); exit()


# --- 从 Config 读取全局配置 ---
BASE_OUTPUT_DIR = Path(config.get('General', 'base_output_dir', fallback='./full_process_output'))
CLEANUP_TEMP_DIR = config.getboolean('General', 'cleanup_temp_dir', fallback=True)
FFMPEG_PATH = config.get('Paths', 'ffmpeg_path', fallback='ffmpeg')


# --- 依赖检查 ---
def check_dependencies():
    """检查必要的外部依赖 (例如 FFmpeg)。"""
    logging.info("正在检查依赖项...")
    ffmpeg_ok = False
    try:
        # 使用 shutil.which 查找 ffmpeg 路径可能更跨平台
        resolved_ffmpeg_path = shutil.which(FFMPEG_PATH)
        if resolved_ffmpeg_path:
            logging.info(f"找到 FFmpeg 可执行文件: {resolved_ffmpeg_path}")
            cmd = [resolved_ffmpeg_path, "-version"]
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=5, encoding='utf-8', errors='ignore')
            if result.returncode == 0 and "ffmpeg version" in result.stdout.lower():
                logging.info("FFmpeg 版本检查通过。")
                ffmpeg_ok = True
            else:
                logging.error(f"FFmpeg 命令 '{FFMPEG_PATH}' 执行异常或输出不符合预期。")
                logging.error(f"Stdout:\n{result.stdout}")
                logging.error(f"Stderr:\n{result.stderr}")
        else:
            logging.error(f"错误：在系统 PATH 或指定路径中找不到 FFmpeg 可执行文件: '{FFMPEG_PATH}'")

    except FileNotFoundError:
        logging.error(f"错误：找不到 FFmpeg 命令 '{FFMPEG_PATH}'。请确保已安装 FFmpeg 并将其添加到系统 PATH 环境变量，或在 config.ini 中指定完整路径。")
    except subprocess.TimeoutExpired:
        logging.error("检查 FFmpeg 版本超时。")
    except Exception as e:
        logging.exception(f"检查 FFmpeg 时发生未知错误: {e}") # 使用 exception 记录堆栈

    if not ffmpeg_ok:
        logging.error("关键依赖项 FFmpeg 未满足。程序将退出。")
        return False

    # 可以添加其他依赖检查，例如 Office/LibreOffice (如果需要自动导出)

    logging.info("依赖项检查完成。")
    return True


# --- 主处理函数 ---
def run_full_process(input_pptx_path: Path): # 接受输入路径作为参数
    """
    Executes the entire PPT to Video conversion process for a given PPTX file.
    """
    start_time = time.time()
    logging.info("="*20 + f" Starting Process for: {input_pptx_path.name} " + "="*20)

    # 1. 确认基础输出目录
    try:
        BASE_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        logging.info(f"基础输出目录: {BASE_OUTPUT_DIR.resolve()}")
    except OSError as e:
        logging.error(f"无法创建或访问基础输出目录 '{BASE_OUTPUT_DIR}': {e}"); return

    # 2. 定义最终视频输出路径
    final_video_filename = input_pptx_path.stem + "_final_video.mp4"
    final_video_path = BASE_OUTPUT_DIR / final_video_filename
    if final_video_path.exists():
        logging.warning(f"输出视频文件已存在: {final_video_path}。将被覆盖。")
        try: final_video_path.unlink()
        except OSError as e: logging.error(f"无法删除已存在的输出文件 '{final_video_path}': {e}"); return

    # --- 步骤 1 & 2: 处理 PPT ---
    logging.info("--- 阶段 1/2: 处理演示文稿 (导出, 备注, 音频)... ---")
    processed_data = None; temp_run_dir = None
    try:
        result_tuple = process_presentation(input_pptx_path, BASE_OUTPUT_DIR)
        if result_tuple and isinstance(result_tuple, tuple) and len(result_tuple) == 2:
            processed_data, temp_run_dir = result_tuple
            if processed_data is None or temp_run_dir is None: # 检查内部是否返回了 None
                 raise ValueError("process_presentation 返回了无效数据或路径。") # 抛出异常以便统一处理
            logging.info(f"演示文稿处理完成。处理了 {len(processed_data)} 张幻灯片。")
            logging.info(f"临时文件位于: {temp_run_dir.resolve()}")
        else:
             raise ValueError("process_presentation 返回结果格式不符合预期。")
    except Exception as e:
        logging.exception(f"处理演示文稿时出错: {e}") # 使用 exception 记录完整错误
        if temp_run_dir and temp_run_dir.exists() and CLEANUP_TEMP_DIR:
            logging.warning(f"尝试清理部分生成的临时目录: {temp_run_dir}")
            try: shutil.rmtree(temp_run_dir)
            except Exception as clean_e: logging.error(f"清理失败: {clean_e}")
        return # 中止

    # --- 步骤 3: 合成视频 ---
    logging.info("--- 阶段 3/3: 合成视频 (拼接, 字幕)... ---")
    synthesis_success = False
    try:
        synthesis_success = create_video_from_data(
            processed_data, temp_run_dir, final_video_path
        )
    except Exception as e:
        logging.exception(f"合成视频时发生意外错误: {e}")
        synthesis_success = False

    # --- 步骤 4: 最终输出和清理 ---
    end_time = time.time()
    total_time = end_time - start_time
    if synthesis_success and final_video_path.exists():
        logging.info("="*20 + " 处理成功完成! " + "="*20)
        logging.info(f"最终视频: {final_video_path.resolve()}")
        logging.info(f"总耗时: {total_time:.2f} 秒")
        if CLEANUP_TEMP_DIR:
            logging.info(f"开始清理临时目录: {temp_run_dir}")
            try: shutil.rmtree(temp_run_dir); logging.info("临时目录已清理。")
            except Exception as e: logging.warning(f"清理临时目录时出错: {e}")
        else:
            logging.info(f"临时文件保留于: {temp_run_dir}")
    else:
        logging.error("="*20 + " 处理失败! " + "="*20)
        logging.error("视频合成失败或未生成输出文件。请检查以上日志。")
        logging.error(f"总耗时: {total_time:.2f} 秒")
        if temp_run_dir and temp_run_dir.exists():
            logging.info(f"临时文件保留于: {temp_run_dir.resolve()} 以供检查。")


# --- 主程序入口 ---
if __name__ == "__main__":
    # --- 设置命令行参数解析 ---
    parser = argparse.ArgumentParser(description="将 PowerPoint (.pptx) 文件转换为带旁白和字幕的视频。")
    parser.add_argument("input_pptx", help="输入的 PPTX 文件路径。")
    # 可以添加更多参数来覆盖 config.ini 中的设置，例如：
    parser.add_argument("-o", "--output-dir", help="指定输出目录 (覆盖 config.ini)。")
    parser.add_argument("--no-cleanup", action="store_true", help="即使成功也保留临时文件 (覆盖 config.ini)。")

    args = parser.parse_args()

    # --- 将参数转换为 Path 对象 ---
    input_file_path = Path(args.input_pptx)

    # --- 检查输入文件 ---
    if not input_file_path.is_file() or not input_file_path.name.lower().endswith(".pptx"):
        print(f"错误: 输入文件 '{input_file_path}' 不是一个有效的 .pptx 文件或文件不存在。")
        logging.error(f"无效的输入文件: {input_file_path}")
        exit(1) # 以错误码退出

    # --- 运行依赖检查 ---
    if check_dependencies():
        # --- 运行主处理流程 ---
        run_full_process(input_file_path)
    else:
        print("错误：缺少必要的依赖项，无法继续。请检查日志。")
        exit(1)

    print("\n脚本执行完毕。")