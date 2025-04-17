import os
import logging
import time
import shutil
from pathlib import Path
import subprocess # 用于调用外部命令 (FFmpeg)
import shlex     # 用于安全地构建命令行参数
import wave      # 需要在 ASR 部分之外，为下面的函数使用
import contextlib

try:
    from PIL import Image # <--- 确保这一行存在
except ImportError:
    logging.error("缺少 'Pillow' 库。请运行 'pip install Pillow'。")
    # exit()
# --- 导入 ASR 相关库 (保持不变) ---
try:
    import stable_whisper
except ImportError:
    logging.error("缺少 'stable-ts' 库。请运行 'pip install stable-ts'。")
    # exit()

# --- 配置 ---
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - [%(funcName)s] - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

TARGET_WIDTH = 1280      # 目标视频宽度 (像素)
# TARGET_HEIGHT = 720    # 目标视频高度 (像素) - FFmpeg 可以自动计算或需要指定
TARGET_FPS = 24          # 目标视频帧率
# TRANSITION_DURATION = 0.5 # 转场暂时移除
WHISPER_MODEL = "base"   # 使用的 Whisper 模型
DEFAULT_SLIDE_DURATION = 3.0 # 如果幻灯片无音频/备注，默认显示时长 (秒)
SUBTITLE_FONT = 'Arial'  # 字幕字体 (确保系统支持，FFmpeg可能需要特定配置)
SUBTITLE_FONTSIZE = 24   # 字幕字号
SUBTITLE_COLOR = 'white' # 字幕颜色
SUBTITLE_BG_COLOR = '0x000000@0.5' # 字幕背景色 (FFmpeg 格式: 0xRRGGBB@AA)
FFMPEG_PATH = "ffmpeg"   # 假设 ffmpeg 在 PATH 中，否则指定完整路径

# --- ASR 字幕生成函数 (基本保持不变) ---
def get_wav_duration(filepath: Path) -> float:
    """获取 WAV 文件的时长（秒）。"""
    if not filepath.is_file():
        logging.warning(f"尝试获取时长失败，文件不存在: {filepath}")
        return 0.0
    try:
        with contextlib.closing(wave.open(str(filepath), 'r')) as f:
            frames = f.getnframes()
            rate = f.getframerate()
            if rate == 0:
                logging.warning(f"文件采样率读取为零: {filepath}")
                return 0.0
            duration = frames / float(rate)
            return duration
    except wave.Error as e:
        logging.error(f"读取 WAV 文件头出错 {filepath}: {e}")
        return 0.0
    except Exception as e:
        logging.error(f"获取 WAV 时长时发生意外错误 {filepath}: {e}")
        return 0.0

def srt_formatter(result: stable_whisper.WhisperResult, **kwargs) -> str:
    """将 stable-ts 结果格式化为 SRT"""
    return result.to_srt_vtt(word_level=False) # 使用段落级别的 SRT

def generate_subtitles(
    audio_paths: list[str | None],
    output_srt_path: Path,
    temp_dir: Path,
    whisper_model_name: str = WHISPER_MODEL
) -> bool:
    """
    合并音频文件，使用 Whisper 生成 SRT 字幕文件。
    (使用 MoviePy 合并音频，因为这部分通常没问题且方便)
    """
    logging.info("开始生成字幕...")
    valid_audio_files = [p for p in audio_paths if p and Path(p).exists() and Path(p).stat().st_size > 100]

    if not valid_audio_files:
        logging.warning("没有有效的音频文件可用于生成字幕。")
        return False

    combined_audio_path = temp_dir / "combined_audio_for_asr.wav"

    # --- 使用 FFmpeg 合并音频 ---
    # 创建一个包含所有输入音频文件路径的文本文件 (safe way)
    concat_list_path = temp_dir / "audio_concat_list.txt"
    try:
        with open(concat_list_path, 'w', encoding='utf-8') as f:
            for audio_file in valid_audio_files:
                # FFmpeg concat demuxer 需要特定的格式，并处理特殊字符
                safe_path = str(Path(audio_file).resolve()).replace("'", "'\\''") # 基本的转义
                f.write(f"file '{safe_path}'\n")
        logging.info(f"为 FFmpeg 创建了音频合并列表: {concat_list_path.name}")

        cmd_concat = [
            FFMPEG_PATH,
            "-f", "concat",      # 使用 concat demuxer
            "-safe", "0",       # 允许非本地/相对路径 (如果需要)
            "-i", str(concat_list_path.resolve()), # 输入列表文件
            "-c", "copy",       # 直接复制音频流，不重新编码
            str(combined_audio_path.resolve()) # 输出合并后的文件
        ]
        logging.info(f"执行 FFmpeg 命令合并音频: {' '.join(shlex.quote(c) for c in cmd_concat)}")
        result = subprocess.run(cmd_concat, capture_output=True, text=True, check=True, encoding='utf-8')
        if result.stderr: logging.debug(f"FFmpeg (concat) stderr:\n{result.stderr}") # Debug 输出
        logging.info("使用 FFmpeg 合并音频完成。")
        # 可选删除列表文件: concat_list_path.unlink()

    except subprocess.CalledProcessError as e:
        logging.error(f"FFmpeg 合并音频失败。返回码: {e.returncode}")
        logging.error(f"FFmpeg 命令: {' '.join(shlex.quote(c) for c in cmd_concat)}")
        logging.error(f"FFmpeg 错误输出:\n{e.stderr}")
        return False
    except FileNotFoundError:
        logging.error(f"错误：找不到 '{FFMPEG_PATH}' 命令。请确保 FFmpeg 已安装并配置在系统 PATH 中。")
        return False
    except Exception as e:
        logging.error(f"合并音频时发生错误: {e}", exc_info=True)
        return False

    # --- 后续 ASR 步骤保持不变 ---
    try:
        logging.info(f"加载 Whisper 模型 '{whisper_model_name}'...")
        asr_start_time = time.time()
        model = stable_whisper.load_model(whisper_model_name)
        logging.info("开始语音识别 (ASR)...")
        result = model.transcribe(str(combined_audio_path), fp16=False, verbose=False)
        asr_end_time = time.time()
        logging.info(f"语音识别完成，耗时 {asr_end_time - asr_start_time:.2f} 秒。")

        logging.info(f"将结果格式化并保存到 {output_srt_path.name}...")
        srt_content = srt_formatter(result)
        with open(output_srt_path, "w", encoding="utf-8") as f:
            f.write(srt_content)
        logging.info("SRT 字幕文件生成成功。")
        return True
    except Exception as e:
        logging.error(f"运行 Whisper ASR 或保存字幕时出错: {e}", exc_info=True)
        return False
    finally:
        if 'model' in locals(): del model
        # 可选删除合并后的音频: combined_audio_path.unlink()


# --- FFmpeg 核心功能函数 ---

def create_video_segment(
    image_path: Path,
    duration: float,
    audio_path: Path | None,
    output_path: Path,
    width: int,
    fps: int
) -> bool:
    """使用 FFmpeg 将单张图片和可选音频转换为视频片段 (采用分步处理)。"""
    logging.info(f"  使用 FFmpeg 创建视频片段: {output_path.name} (时长: {duration:.2f}s)")

    temp_video_path = output_path.with_suffix(".temp_video.mp4") # 临时的无声视频文件
    step1_success = False
    step2_success = False

    # --- 步骤 1: 图片转为无声视频 ---
    cmd_step1 = [
        FFMPEG_PATH,
        "-loop", "1",
        "-framerate", str(fps),
        "-i", str(image_path.resolve()),
        "-vf", f"scale={width}:-2:force_original_aspect_ratio=decrease,pad={width}:{width*9//16}:(ow-iw)/2:(oh-ih)/2,format=yuv420p,fps={fps}",
        "-t", str(duration),
        "-c:v", "libx264",
        "-preset", "veryfast",
        "-crf", "23",
        "-pix_fmt", "yuv420p",
        "-an", # 明确无音频
        str(temp_video_path.resolve())
    ]
    try:
        logging.debug(f"    执行 FFmpeg 命令 (步骤 1 - 图片转无声视频): {' '.join(shlex.quote(c) for c in cmd_step1)}")
        result1 = subprocess.run(cmd_step1, capture_output=True, text=True, check=True, encoding='utf-8', errors='ignore')
        if result1.stderr: logging.debug(f"    FFmpeg (step1) stderr:\n{result1.stderr}")
        logging.info(f"    步骤 1 成功: 已生成无声视频 {temp_video_path.name}")
        step1_success = True
    except subprocess.CalledProcessError as e:
        logging.error(f"  FFmpeg 创建无声视频失败: {temp_video_path.name}。返回码: {e.returncode}")
        logging.error(f"  FFmpeg 命令: {' '.join(shlex.quote(c) for c in cmd_step1)}")
        logging.error(f"  FFmpeg 标准错误输出:\n{e.stderr}")
        logging.error(f"  FFmpeg 标准输出:\n{e.stdout}")
        # 如果第一步失败，尝试清理临时文件并返回 False
        if temp_video_path.exists():
             try: temp_video_path.unlink()
             except OSError: pass
        return False
    except FileNotFoundError:
        logging.error(f"错误：找不到 '{FFMPEG_PATH}' 命令。")
        return False
    except Exception as e:
        logging.error(f"  创建无声视频时发生未知错误 {temp_video_path.name}: {e}")
        return False

    # --- 步骤 2: 合并无声视频和音频 (如果音频存在) ---
    if step1_success:
        if audio_path and audio_path.is_file() and audio_path.stat().st_size > 0:
            logging.info(f"    步骤 2: 合并视频与音频 {audio_path.name} 到 {output_path.name}")
            cmd_step2 = [
                FFMPEG_PATH,
                "-i", str(temp_video_path.resolve()), # 输入无声视频
                "-i", str(audio_path.resolve()),    # 输入音频
                "-c:v", "copy",                     # 视频流直接复制 (因为编码已在步骤1完成)
                "-c:a", "aac",                      # 音频编码
                "-b:a", "128k",
                "-shortest",                        # 以最短输入结束
                str(output_path.resolve())          # 最终输出
            ]
            try:
                logging.debug(f"    执行 FFmpeg 命令 (步骤 2 - 合并音视频): {' '.join(shlex.quote(c) for c in cmd_step2)}")
                result2 = subprocess.run(cmd_step2, capture_output=True, text=True, check=True, encoding='utf-8', errors='ignore')
                if result2.stderr: logging.debug(f"    FFmpeg (step2) stderr:\n{result2.stderr}")
                logging.info(f"    步骤 2 成功: 已合并音视频到 {output_path.name}")
                step2_success = True
            except subprocess.CalledProcessError as e:
                logging.error(f"  FFmpeg 合并音视频失败: {output_path.name}。返回码: {e.returncode}")
                logging.error(f"  FFmpeg 命令: {' '.join(shlex.quote(c) for c in cmd_step2)}")
                logging.error(f"  FFmpeg 标准错误输出:\n{e.stderr}")
                logging.error(f"  FFmpeg 标准输出:\n{e.stdout}")
                # 合并失败，但无声视频已生成，可以保留它吗？或者算失败？这里算失败
                step2_success = False
            except FileNotFoundError:
                 logging.error(f"错误：找不到 '{FFMPEG_PATH}' 命令。")
                 step2_success = False
            except Exception as e:
                 logging.error(f"  合并音视频时发生未知错误 {output_path.name}: {e}")
                 step2_success = False
            finally:
                 # 无论合并成功与否，删除临时的无声视频文件
                 if temp_video_path.exists():
                     try: temp_video_path.unlink()
                     except OSError: pass
        else:
            # 如果没有音频，直接将无声视频重命名/移动为最终输出
            logging.info(f"    步骤 2: 无音频，直接使用无声视频 {temp_video_path.name} 作为输出 {output_path.name}")
            try:
                shutil.move(str(temp_video_path.resolve()), str(output_path.resolve()))
                step2_success = True
            except Exception as e:
                 logging.error(f"重命名/移动无声视频失败: {e}")
                 step2_success = False
                 # 如果移动失败，尝试删除临时文件
                 if temp_video_path.exists():
                      try: temp_video_path.unlink()
                      except OSError: pass

    return step1_success and step2_success # 只有两步（或第一步+移动）都成功才算成功

def concatenate_videos(video_file_list_path: Path, output_path: Path) -> bool:
    """使用 FFmpeg concat demuxer 拼接视频文件。"""
    logging.info(f"使用 FFmpeg concat demuxer 拼接视频...")
    cmd_list = [
        FFMPEG_PATH,
        "-f", "concat",
        "-safe", "0", # 允许绝对路径
        "-i", str(video_file_list_path.resolve()),
        "-c", "copy", # 直接复制代码流，速度快，不重新编码
        str(output_path.resolve())
    ]
    try:
        logging.debug(f"  执行 FFmpeg 命令: {' '.join(shlex.quote(c) for c in cmd_list)}")
        result = subprocess.run(cmd_list, capture_output=True, text=True, check=True, encoding='utf-8')
        if result.stderr: logging.debug(f"  FFmpeg (concat) stderr:\n{result.stderr}")
        logging.info(f"视频拼接成功: {output_path.name}")
        return True
    except subprocess.CalledProcessError as e:
        logging.error(f"FFmpeg 拼接视频失败。返回码: {e.returncode}")
        logging.error(f"FFmpeg 命令: {' '.join(shlex.quote(c) for c in cmd_list)}")
        logging.error(f"FFmpeg 错误输出:\n{e.stderr}")
        return False
    except FileNotFoundError:
         logging.error(f"错误：找不到 '{FFMPEG_PATH}' 命令。")
         return False
    except Exception as e:
         logging.error(f"拼接视频时发生未知错误: {e}")
         return False


import shlex # 确保导入 shlex

def add_subtitles(input_video: Path, srt_file: Path, output_video: Path) -> bool:
    """
    使用 FFmpeg 将 SRT 字幕硬编码到视频中 (改进路径处理)。

    Args:
        input_video: 输入视频文件的 Path 对象。
        srt_file: SRT 字幕文件的 Path 对象。
        output_video: 输出视频文件的 Path 对象。

    Returns:
        bool: 字幕添加成功返回 True，否则返回 False。
    """
    logging.info(f"使用 FFmpeg 添加字幕到视频...")

    # 直接使用 resolve() 获取绝对路径，交给 shlex.quote 处理特殊字符
    # 注意：shlex.quote 在 Windows 上可能不会按预期处理反斜杠和冒号，
    # FFmpeg 的 subtitles 滤镜对 Windows 路径有特殊要求。
    # 尝试一种更可靠的 Windows 路径转义方法给 FFmpeg filter。
    srt_path_str = str(srt_file.resolve())
    # FFmpeg 滤镜路径需要转义 '\' 为 '\\' 或 '/'，并且转义 ':' 为 '\:'
    srt_path_escaped_for_filter = srt_path_str.replace('\\', '/').replace(':', '\\:')

    input_video_str = str(input_video.resolve())
    output_video_str = str(output_video.resolve())

    # 构建 filtergraph 字符串，使用转义后的路径
    # 先尝试不带 force_style
    vf_filter = f"subtitles='{srt_path_escaped_for_filter}'"
    # 如果需要 force_style (确保字体文件可用或使用通用字体名):
    # style_options = f"FontName={SUBTITLE_FONT},FontSize={SUBTITLE_FONTSIZE}" # 简化版
    # vf_filter = f"subtitles='{srt_path_escaped_for_filter}':force_style='{style_options}'"

    cmd_list = [
        FFMPEG_PATH,
        "-i", input_video_str,
        "-vf", vf_filter,
        "-c:v", "libx264",
        "-preset", "medium", # 可以根据需要调整速度 vs 压缩率
        "-crf", "22",       # 调整视频质量 (18-28 常用)
        "-c:a", "copy",     # 直接复制音频流
        output_video_str
    ]
    try:
        # 使用 shlex.join 打印命令更安全可靠
        logging.debug(f"  执行 FFmpeg 命令: {shlex.join(cmd_list)}")
        result = subprocess.run(cmd_list, capture_output=True, text=True, check=True, encoding='utf-8', errors='ignore')
        if result.stderr: logging.debug(f"  FFmpeg (subtitles) stderr:\n{result.stderr}")
        logging.info(f"字幕添加成功: {output_video.name}")
        return True
    except subprocess.CalledProcessError as e:
        logging.error(f"FFmpeg 添加字幕失败。返回码: {e.returncode}")
        logging.error(f"FFmpeg 命令: {shlex.join(cmd_list)}")
        logging.error(f"FFmpeg 标准错误输出:\n{e.stderr}")
        logging.error(f"FFmpeg 标准输出:\n{e.stdout}")
        return False
    except FileNotFoundError:
         logging.error(f"错误：找不到 '{FFMPEG_PATH}' 命令。")
         return False
    except Exception as e:
         logging.error(f"添加字幕时发生未知错误: {e}")
         return False

# --- 视频合成主函数 (重写) ---
def create_video_from_data(
    processed_data: list[dict],
    temp_run_dir: Path,
    output_video_path: Path
) -> bool:
    """
    根据处理好的数据，使用 FFmpeg 合成最终视频 (无转场)。
    添加了对生成的 SRT 文件是否有效的检查。

    Args:
        processed_data: 字典列表，包含幻灯片信息。
        temp_run_dir: 临时工作目录。
        output_video_path: 最终视频输出路径。

    Returns:
        bool: 成功返回 True，失败返回 False。
    """
    logging.info("--- 开始基于 FFmpeg 的视频合成流程 ---")
    if not processed_data:
        logging.error("输入数据为空，无法合成视频。")
        return False

    temp_segments_dir = temp_run_dir / "video_segments"
    temp_segments_dir.mkdir(exist_ok=True)
    segment_files = [] # 存储成功生成的视频片段路径

    # --- 1. 生成各幻灯片的视频片段 ---
    logging.info("步骤 1: 使用 FFmpeg 生成各幻灯片的视频片段")
    for i, data in enumerate(processed_data):
        slide_num = data.get('slide_number', i + 1)
        image_path_str = data.get('image_path')
        audio_path_str = data.get('audio_path')
        duration = data.get('audio_duration', 0.0)

        if not image_path_str or not Path(image_path_str).is_file():
            logging.warning(f"幻灯片 {slide_num}: 图片路径无效或丢失。跳过此片段。")
            continue

        image_path = Path(image_path_str)
        audio_path = Path(audio_path_str) if audio_path_str else None

        clip_duration = duration if duration > 0.01 else DEFAULT_SLIDE_DURATION
        segment_output_path = temp_segments_dir / f"segment_{slide_num}.mp4"

        success = create_video_segment(
            image_path,
            clip_duration,
            audio_path,
            segment_output_path,
            TARGET_WIDTH,
            TARGET_FPS
        )
        if success:
            segment_files.append(segment_output_path)
        else:
            logging.error(f"未能创建幻灯片 {slide_num} 的视频片段。合成中止。")
            return False # 一个片段失败就中止

    if not segment_files:
        logging.error("未能成功生成任何视频片段。")
        return False

    # --- 2. 拼接视频片段 ---
    logging.info("步骤 2: 使用 FFmpeg 拼接视频片段")
    concat_list_file = temp_run_dir / "video_concat_list.txt"
    try:
        with open(concat_list_file, 'w', encoding='utf-8') as f:
            for segment_file in segment_files:
                # 使用正斜杠，ffmpeg concat demuxer 对其兼容性更好
                safe_path = str(segment_file.resolve()).replace('\\', '/')
                f.write(f"file '{safe_path}'\n")
    except Exception as e:
        logging.error(f"创建视频拼接列表文件时出错: {e}")
        return False

    base_video_path = temp_run_dir / "base_video_no_subs.mp4" # 拼接后无字幕的基础视频
    success = concatenate_videos(concat_list_file, base_video_path)
    if not success:
        logging.error("拼接视频片段失败。")
        try: concat_list_file.unlink()
        except OSError: pass
        return False
    try: concat_list_file.unlink()
    except OSError: pass

    # --- 3. 生成字幕 (如果需要) ---
    logging.info("步骤 3: 生成字幕文件 (ASR)")
    audio_segment_paths = [d.get('audio_path') for d in processed_data]
    subtitle_file_path = temp_run_dir / "subtitles.srt"
    subtitles_generated = generate_subtitles(
        audio_segment_paths,
        subtitle_file_path,
        temp_run_dir
    )

    # --- 检查生成的 SRT 文件是否有效 (非空且包含文本) ---
    srt_is_valid = False
    if subtitles_generated and subtitle_file_path.exists():
        try:
            if subtitle_file_path.stat().st_size > 10: # 基础大小检查
                with open(subtitle_file_path, 'r', encoding='utf-8') as f:
                    for line in f:
                        line_strip = line.strip()
                        # 检查是否包含有效文本行 (不是空行, 不是纯数字序列号, 不是时间戳)
                        if line_strip and not line_strip.isdigit() and '-->' not in line_strip:
                            srt_is_valid = True
                            break # 找到一个有效文本行就足够了
                if srt_is_valid:
                    logging.info("生成的 SRT 字幕文件包含有效文本。")
                else:
                     logging.warning("生成的 SRT 文件为空或不包含有效文本内容。")
            else:
                logging.warning("生成的 SRT 文件为空。")
        except Exception as e:
            logging.warning(f"检查 SRT 文件时出错: {e}")

    # --- 4. 添加字幕到视频 (仅当 SRT 文件有效时) ---
    if srt_is_valid: # <--- 使用检查结果
        logging.info("步骤 4: 使用 FFmpeg 添加字幕")
        final_video_with_subs_path = temp_run_dir / "final_video_with_subs.mp4"
        success = add_subtitles(base_video_path, subtitle_file_path, final_video_with_subs_path)
        if success:
            logging.info("字幕添加成功。将带有字幕的视频作为最终输出。")
            try:
                 shutil.move(str(final_video_with_subs_path), str(output_video_path))
                 logging.info(f"最终视频 (带字幕) 已保存到: {output_video_path}")
                 try: base_video_path.unlink() # 清理无字幕基础视频
                 except OSError: pass
                 return True # 整个流程成功
            except Exception as e:
                 logging.error(f"移动最终带字幕视频时出错: {e}. 最终文件可能仍在临时目录: {final_video_with_subs_path}")
                 return False
        else:
            logging.error("添加字幕失败。将输出不带字幕的视频。")
            # 添加字幕失败，则回退到输出不带字幕的视频
            try:
                 shutil.move(str(base_video_path), str(output_video_path))
                 logging.info(f"最终视频 (无字幕 - 因添加失败) 已保存到: {output_video_path}")
                 return True # 流程算成功，但功能未完全实现
            except Exception as e:
                 logging.error(f"移动最终无字幕视频时出错: {e}. 最终文件可能仍在临时目录: {base_video_path}")
                 return False
    else:
        # 如果 SRT 无效或生成失败，则直接输出不带字幕的视频
        logging.info("步骤 4: 跳过添加字幕 (文件为空、无效或生成失败)。")
        try:
             shutil.move(str(base_video_path), str(output_video_path))
             logging.info(f"最终视频 (无字幕) 已保存到: {output_video_path}")
             return True
        except Exception as e:
             logging.error(f"移动最终无字幕视频时出错: {e}. 最终文件可能仍在临时目录: {base_video_path}")
             return False

    # 清理逻辑已包含在上面的移动操作中，这里不再需要单独的清理步骤
    # （除非你想在失败时也清理临时片段，但保留它们有助于调试）

    # --- 5. 清理临时视频片段 ---
    # 注意：这步实际上被上面的移动操作覆盖了，如果移动成功，应该在移动后清理
    # 这里保留逻辑以防未来修改
    # logging.info("步骤 5: 清理临时视频片段文件...")
    # try:
    #     shutil.rmtree(temp_segments_dir)
    # except OSError as e:
    #     logging.warning(f"清理临时视频片段目录时出错: {e}")

    # return True # 如果代码能执行到这里，说明前面的步骤都成功了


# --- 主程序入口与测试 (使用 FFmpeg 版本) ---
# --- 主程序入口与测试 (使用 FFmpeg 版本 + 真实语音模拟) ---
if __name__ == "__main__":
    logging.info("--- 开始测试基于 FFmpeg 的视频合成模块 (使用 TTS 生成模拟语音) ---")

    # --- 引入 TTS 库 ---
    try:
        import pyttsx3
    except ImportError:
        logging.error("测试需要 'pyttsx3' 库。请运行 'pip install pyttsx3'。")
        exit() # 如果无法导入 TTS，则无法进行此测试

    # --- 模拟输入数据和环境 ---
    mock_run_dir = Path("./mock_run_for_ffmpeg_real_speech_test") # 使用新目录名
    mock_images_dir = mock_run_dir / "images"
    mock_audio_dir = mock_run_dir / "audio"

    if mock_run_dir.exists(): shutil.rmtree(mock_run_dir)
    mock_images_dir.mkdir(parents=True, exist_ok=True)
    mock_audio_dir.mkdir(parents=True, exist_ok=True)

    # --- 定义模拟的备注文本 (用于 TTS) ---
    mock_notes_texts = [
        "这是第一张幻灯片的语音旁白。", # 幻灯片 1
        None,                             # 幻灯片 2 (无旁白)
        "现在我们来看第三部分，这里有一些重要信息。" # 幻灯片 3
    ]

    mock_image_files = []
    mock_audio_files = []
    mock_durations = []
    tts_engine = None # 初始化 TTS 引擎变量

    try:
        # --- 初始化 TTS 引擎 ---
        tts_engine = pyttsx3.init()
        tts_engine.setProperty('rate', 180) # 设置语速
        logging.info("TTS 引擎初始化成功。")

        # --- 创建模拟的图片和 *语音* 文件 ---
        logging.info("正在创建模拟图片和生成语音文件...")
        for i, note_text in enumerate(mock_notes_texts):
            slide_num = i + 1
            # 创建模拟图片 (颜色区分)
            img_color = ['red', 'blue', 'green'][i % 3]
            img_path = mock_images_dir / f"slide_{slide_num}.png"
            try:
                Image.new('RGB', (TARGET_WIDTH, 720), color=img_color).save(img_path)
                mock_image_files.append(str(img_path))
                logging.debug(f"  创建模拟图片: {img_path.name}")
            except Exception as e:
                 logging.error(f"  创建模拟图片 {slide_num} 失败: {e}")
                 mock_image_files.append(None) # 添加占位符
                 mock_audio_files.append(None)
                 mock_durations.append(0.0)
                 continue # 处理下一张

            # 如果有备注文本，则生成语音
            if note_text:
                audio_path = mock_audio_dir / f"segment_{slide_num}.wav"
                abs_audio_path = str(audio_path.resolve())
                duration = 0.0
                audio_file_path_str = None
                try:
                    logging.info(f"  使用 TTS 生成幻灯片 {slide_num} 的音频...")
                    tts_engine.save_to_file(note_text, abs_audio_path)
                    tts_engine.runAndWait() # 等待文件保存

                    if audio_path.exists() and audio_path.stat().st_size > 100:
                        # 获取实际生成的音频时长
                        duration = get_wav_duration(audio_path)
                        if duration > 0.01:
                             logging.info(f"    音频生成成功: {audio_path.name}, 时长: {duration:.2f}s")
                             audio_file_path_str = abs_audio_path
                        else:
                             logging.warning(f"    音频文件 {audio_path.name} 已生成但无法获取有效时长。")
                             duration = 0.0 # 视为无效
                    else:
                        logging.warning(f"    TTS 未能保存文件或保存了空文件: {audio_path.name}")

                except Exception as e:
                     logging.error(f"    为幻灯片 {slide_num} 生成 TTS 音频时出错: {e}")
                finally:
                    mock_audio_files.append(audio_file_path_str)
                    mock_durations.append(duration)
            else:
                # 没有备注文本，无音频
                logging.info(f"  幻灯片 {slide_num} 无备注文本，跳过音频生成。")
                mock_audio_files.append(None)
                mock_durations.append(0.0)

        # --- 构建最终的 processed_data (过滤掉图片创建失败的项) ---
        mock_processed_data = []
        valid_count = 0
        for i in range(len(mock_notes_texts)):
             # 确保所有对应的数据都有效
             if i < len(mock_image_files) and mock_image_files[i] and \
                i < len(mock_audio_files) and \
                i < len(mock_durations):
                  mock_processed_data.append({
                      'slide_number': i + 1,
                      'image_path': mock_image_files[i],
                      'notes': mock_notes_texts[i] or "", # 确保 notes 是字符串
                      'audio_path': mock_audio_files[i],
                      'audio_duration': mock_durations[i]
                  })
                  valid_count += 1
             else:
                 logging.warning(f"跳过构建幻灯片 {i+1} 的数据，因为图片或关联数据缺失。")

        if valid_count > 0:
             logging.info(f"成功构建了 {valid_count} 条模拟数据。")
        else:
             logging.error("未能构建任何有效的模拟数据。")
             mock_processed_data = [] # 确保为空

    except Exception as e:
        logging.error(f"创建模拟文件或 TTS 初始化时出错: {e}")
        mock_processed_data = [] # 出错则清空数据
    finally:
        # 清理 TTS 引擎
        if tts_engine:
            del tts_engine

    # --- 指定最终视频输出路径 ---
    final_output_video = Path("./final_video_ffmpeg_real_speech.mp4") # 新文件名

    # --- 执行视频合成 ---
    if mock_processed_data:
        if final_output_video.exists(): final_output_video.unlink() # 确保输出文件不存在

        success = create_video_from_data(
            mock_processed_data,
            mock_run_dir,
            final_output_video
        )

        if success and final_output_video.exists():
            print("\n--- 视频合成测试成功 (使用模拟真实语音) ---")
            print(f"最终视频已保存到: {final_output_video.resolve()}")
            print(f"检查临时目录 '{mock_run_dir.name}' 中的 'subtitles.srt' 文件内容。")
            print("用播放器打开视频，检查画面、声音和字幕是否同步。")
            # 可选清理: shutil.rmtree(mock_run_dir)
        else:
            print("\n--- 视频合成测试失败 (使用模拟真实语音) ---")
            print("请检查上面的日志获取详细错误信息。")
            print(f"确保 FFmpeg 命令 '{FFMPEG_PATH}' 可执行。")
            print(f"检查临时目录 '{mock_run_dir.name}' 中的文件。")
    else:
        print("未能生成有效的模拟数据，无法进行视频合成测试。")

    logging.info("--- 基于 FFmpeg+模拟语音 的视频合成模块测试结束 ---")