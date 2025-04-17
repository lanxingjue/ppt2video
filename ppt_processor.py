import os
import platform
import time
import shutil
import wave # 用于读取 WAV 文件信息
import contextlib # 用于安全地打开和关闭文件
from pathlib import Path
import logging
import uuid # 用于创建唯一的临时文件夹名称

# 导入我们之前创建的导出模块
try:
    from ppt_exporter_win import export_slides_with_powerpoint
except ImportError:
    logging.error("无法导入 'ppt_exporter_win.py'。请确保该文件存在于同一目录或 Python 路径中。")
    # 如果无法导入，后续依赖此模块的功能将无法运行，可以选择退出或提供替代方案
    # exit() # 或者定义一个假的导出函数用于测试其他部分

# 导入 PPT 解析和 TTS 库
try:
    from pptx import Presentation
except ImportError:
    logging.error("缺少 'python-pptx' 库。请使用 'pip install python-pptx' 安装。")
    # exit()
try:
    import pyttsx3
except ImportError:
    logging.error("缺少 'pyttsx3' 库。请使用 'pip install pyttsx3' 安装。")
    # exit()

# --- 配置日志记录 ---
# (和之前脚本保持一致或根据需要调整)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - [%(funcName)s] - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

# --- TTS 相关配置 ---
TTS_RATE = 180  # 语速 (词/分钟)，可以根据需要调整
TTS_VOICE_ID = None # 默认使用第一个找到的语音。可以指定特定 ID (需要先查询可用 ID)

# --- 工具函数 ---
def get_wav_duration(filepath: Path) -> float:
    """获取 WAV 文件的时长（秒）。"""
    if not filepath.is_file():
        logging.warning(f"尝试获取时长失败，文件不存在: {filepath}")
        return 0.0
    try:
        with contextlib.closing(wave.open(str(filepath), 'r')) as f:
            frames = f.getnframes()
            rate = f.getframerate()
            if rate == 0: # 防止除以零错误
                logging.warning(f"文件采样率读取为零: {filepath}")
                return 0.0
            duration = frames / float(rate)
            return duration
    except wave.Error as e:
        logging.error(f"读取 WAV 文件头出错 {filepath}: {e}")
        return 0.0 # 出错时返回 0
    except Exception as e:
        logging.error(f"获取 WAV 时长时发生意外错误 {filepath}: {e}")
        return 0.0

# --- 核心处理函数 ---

def extract_speaker_notes(pptx_filepath: Path) -> list[str] | None:
    """
    从 PPTX 文件中提取每张幻灯片的演讲者备注。

    Args:
        pptx_filepath: 输入的 PPTX 文件的 Path 对象。

    Returns:
        一个字符串列表，列表中的每个元素是对应幻灯片的备注文本。
        如果某张幻灯片没有备注，则对应元素为空字符串 ""。
        如果发生错误，返回 None。
    """
    if not pptx_filepath.is_file():
        logging.error(f"输入文件不存在: {pptx_filepath}")
        return None

    notes_list = []
    try:
        logging.info(f"开始解析演示文稿以提取备注: {pptx_filepath.name}")
        prs = Presentation(pptx_filepath)
        num_slides = len(prs.slides)
        logging.info(f"演示文稿包含 {num_slides} 张幻灯片。")

        for i, slide in enumerate(prs.slides):
            slide_num = i + 1
            note_text = "" # 默认为空字符串
            if slide.has_notes_slide:
                notes_slide = slide.notes_slide
                text_frame = notes_slide.notes_text_frame
                if text_frame and text_frame.text:
                    note_text = text_frame.text.strip()
                    logging.debug(f"  找到幻灯片 {slide_num} 的备注: '{note_text[:50]}...'") # 记录备注开头部分
            else:
                logging.debug(f"  幻灯片 {slide_num} 没有备注。")
            notes_list.append(note_text)

        logging.info(f"成功提取了 {len(notes_list)} 条备注信息。")
        return notes_list

    except Exception as e:
        logging.error(f"解析 PPTX 文件以提取备注时出错: {e}", exc_info=True) # exc_info=True 会记录堆栈跟踪
        return None

def generate_audio_segments(
    notes: list[str],
    output_audio_dir: Path,
    rate: int = TTS_RATE,
    voice_id: str | None = TTS_VOICE_ID
) -> list[tuple[str | None, float]]:
    """
    将文本备注列表转换为 WAV 音频文件，并记录每个文件的时长。

    Args:
        notes: 包含每张幻灯片备注文本的字符串列表。
        output_audio_dir: 保存生成的 WAV 文件的目标目录。
        rate: TTS 语速。
        voice_id: 要使用的 TTS 语音 ID (如果为 None，则使用默认)。

    Returns:
        一个元组列表，每个元组包含 (生成的音频文件路径 | None, 音频时长)。
        如果某段文本为空或 TTS 生成失败，音频文件路径可能为 None，时长为 0 或默认值。
    """
    audio_results = []
    output_audio_dir.mkdir(parents=True, exist_ok=True) # 确保目录存在

    try:
        engine = pyttsx3.init()
        engine.setProperty('rate', rate)
        if voice_id:
            try:
                engine.setProperty('voice', voice_id)
                logging.info(f"已设置 TTS 语音 ID: {voice_id}")
            except Exception as e:
                 logging.warning(f"设置指定的语音 ID '{voice_id}' 失败: {e}. 将使用默认语音。")

        # --- 可选：打印所有可用语音以供选择 ---
        voices = engine.getProperty('voices')
        logging.info("系统可用 TTS 语音:")
        for i, voice in enumerate(voices):
            logging.info(f"  语音 {i}: ID={voice.id}, Name={voice.name}, Lang={voice.languages}")
        # --------------------------------------

        logging.info(f"开始生成音频片段 (共 {len(notes)} 段)...")
        total_duration = 0.0

        for i, text in enumerate(notes):
            segment_num = i + 1
            audio_filename = f"segment_{segment_num}.wav" # 使用 WAV 格式
            audio_filepath = output_audio_dir / audio_filename
            abs_audio_filepath = str(audio_filepath.resolve())
            duration_sec = 0.0 # 默认为 0
            result_path = None # 默认没有成功生成文件

            # 如果文本为空或仅包含空白，则跳过 TTS，生成一个非常短的静音占位或记录为0时长
            if not text or text.isspace():
                logging.info(f"  片段 {segment_num}: 文本为空，跳过 TTS。")
                # 可以选择生成一个极短的静音文件，或者直接记录 0 时长
                # 这里我们记录 0 时长和 None 路径
                audio_results.append((None, 0.0))
                continue # 处理下一个片段

            logging.info(f"  正在生成片段 {segment_num} 的音频 (文本: '{text[:50]}...')...")
            try:
                engine.save_to_file(text, abs_audio_filepath)
                engine.runAndWait() # 等待文件保存完成

                # 检查文件是否确实生成并获取时长
                if audio_filepath.exists() and audio_filepath.stat().st_size > 0: # 确保文件存在且非空
                    duration_sec = get_wav_duration(audio_filepath)
                    if duration_sec > 0:
                        result_path = abs_audio_filepath
                        logging.info(f"    片段 {segment_num} 音频已保存: {audio_filename} (时长: {duration_sec:.2f}s)")
                        total_duration += duration_sec
                    else:
                        logging.warning(f"    片段 {segment_num} 文件已生成但无法获取有效时长: {audio_filename}。将记为 0s。")
                        # 文件存在但时长为 0，可能文件损坏，路径仍记录，时长为 0
                        result_path = abs_audio_filepath
                        duration_sec = 0.0
                else:
                    logging.warning(f"    TTS 引擎未能成功保存或保存了空文件: {audio_filename}。")
                    # 文件未生成或为空，路径为 None，时长为 0
                    result_path = None
                    duration_sec = 0.0

            except RuntimeError as e:
                 logging.error(f"    处理片段 {segment_num} 时 TTS 引擎运行时错误: {e}")
                 # TTS 引擎内部错误，通常是 runAndWait 问题
                 result_path = None
                 duration_sec = 0.0
            except Exception as e:
                logging.error(f"    生成片段 {segment_num} 音频时发生未知错误: {e}", exc_info=True)
                result_path = None
                duration_sec = 0.0

            audio_results.append((result_path, duration_sec))

        logging.info(f"所有音频片段生成完成。总预估旁白时长: {total_duration:.2f} 秒。")
        del engine # 尝试释放 TTS 引擎资源
        return audio_results

    except Exception as e:
        logging.error(f"初始化或运行 TTS 引擎时出错: {e}", exc_info=True)
        # 如果 TTS 引擎本身失败，返回空列表或对应数量的 None 结果
        return [(None, 0.0)] * len(notes)


def process_presentation(pptx_filepath: Path, base_output_dir: Path):
    """
    完整的处理流程：导出幻灯片 -> 提取备注 -> 生成音频。

    Args:
        pptx_filepath: 输入的 PPTX 文件的 Path 对象。
        base_output_dir: 用于存放所有输出（包括临时文件和最终结果）的基础目录。

    Returns:
        一个包含处理结果的列表，每个元素是一个字典，代表一张幻灯片的信息：
        {
            'slide_number': int,
            'image_path': str | None, # 导出的图片路径
            'notes': str,             # 提取的备注文本
            'audio_path': str | None, # 生成的音频文件路径
            'audio_duration': float   # 音频时长（秒）
        }
        如果处理过程中发生严重错误，返回 None。
    """
    if not pptx_filepath.is_file():
        logging.error(f"输入 PPTX 文件不存在: {pptx_filepath}")
        return None

    # 1. 创建本次运行的唯一临时工作目录
    run_id = uuid.uuid4().hex[:8] # 使用 UUID 创建唯一标识符
    temp_run_dir = base_output_dir / f"temp_{pptx_filepath.stem}_{run_id}"
    temp_image_dir = temp_run_dir / "images"
    temp_audio_dir = temp_run_dir / "audio"

    try:
        temp_run_dir.mkdir(parents=True, exist_ok=True)
        logging.info(f"创建临时工作目录: {temp_run_dir}")
    except OSError as e:
         logging.error(f"无法创建临时工作目录 {temp_run_dir}: {e}")
         return None

    # 2. 导出幻灯片图片 (调用之前的模块)
    logging.info("--- 步骤 1: 导出幻灯片图片 ---")
    # 检查是否能调用导出函数
    if 'export_slides_with_powerpoint' not in globals():
         logging.error("幻灯片导出功能不可用（可能导入失败）。")
         # 根据需要决定是否继续，这里选择停止
         return None

    image_paths = export_slides_with_powerpoint(pptx_filepath, temp_image_dir)
    if not image_paths:
        logging.error("导出幻灯片图片失败。请查看之前的日志。处理中止。")
        # 可选：清理已创建的临时目录
        # shutil.rmtree(temp_run_dir)
        return None
    logging.info(f"成功导出 {len(image_paths)} 张幻灯片图片到: {temp_image_dir}")

    # 3. 提取演讲者备注
    logging.info("--- 步骤 2: 提取演讲者备注 ---")
    notes_list = extract_speaker_notes(pptx_filepath)
    if notes_list is None:
        logging.error("提取演讲者备注失败。处理中止。")
        # shutil.rmtree(temp_run_dir)
        return None
    logging.info(f"成功提取了 {len(notes_list)} 条备注。")

    # 4. 对齐图片和备注
    # 理论上，导出图片的数量应该等于从 Presentation 对象读取的幻灯片数量（也是备注列表的长度）
    num_images = len(image_paths)
    num_notes = len(notes_list)
    if num_images != num_notes:
        logging.warning(f"导出的图片数量 ({num_images}) 与提取的备注数量 ({num_notes}) 不一致！")
        # 决定处理策略：以数量较小者为准，避免索引越界
        min_count = min(num_images, num_notes)
        logging.warning(f"将仅处理前 {min_count} 个匹配的幻灯片/备注对。")
        image_paths = image_paths[:min_count]
        notes_list = notes_list[:min_count]
    else:
        logging.info("图片和备注数量匹配。")

    # 5. 生成音频片段
    logging.info("--- 步骤 3: 生成音频片段 (TTS) ---")
    audio_results = generate_audio_segments(notes_list, temp_audio_dir)
    if len(audio_results) != len(notes_list): # 检查返回结果数量是否一致
         logging.error(f"TTS 处理返回的结果数量 ({len(audio_results)}) 与预期 ({len(notes_list)}) 不符！处理中止。")
         # shutil.rmtree(temp_run_dir)
         return None
    logging.info("音频片段生成（或跳过）完成。")

    # 6. 组合最终结果数据结构
    logging.info("--- 步骤 4: 整理处理结果 ---")
    final_data = []
    for i in range(len(notes_list)): # 使用已对齐的数量
        slide_data = {
            'slide_number': i + 1,
            'image_path': image_paths[i] if i < len(image_paths) else None, # 确保索引有效
            'notes': notes_list[i],
            'audio_path': audio_results[i][0], # 元组的第一个元素是路径
            'audio_duration': audio_results[i][1] # 元组的第二个元素是时长
        }
        final_data.append(slide_data)
        logging.debug(f"  已整理幻灯片 {i+1} 数据: Image={slide_data['image_path']}, Audio={slide_data['audio_path']}, Duration={slide_data['audio_duration']:.2f}s")

    logging.info(f"成功整理了 {len(final_data)} 张幻灯片的数据。")
    # 注意：此时临时文件（图片、音频）仍然保留在 temp_run_dir 中，供后续视频合成使用
    # return final_data
    return final_data, temp_run_dir # Return both the data and the temp dir path

# --- 示例用法 ---
if __name__ == "__main__":
    logging.info("--- 开始测试 PPT 处理流程 (导出图片 -> 提取备注 -> 生成音频) ---")

    # --- 请修改以下路径 ---
    input_ppt_file = Path("智能短信分类平台方案.pptx") # 你的 PPTX 文件路径
    # 所有输出的基础目录，脚本会在其中创建唯一的临时子目录
    output_base_dir = Path("./processing_output")
    # --- 修改结束 ---

    if not input_ppt_file.exists():
         print(f"错误：请将第 298 行的 'your_presentation.pptx' 替换为您要测试的实际 PPTX 文件路径。")
    else:
        output_base_dir.mkdir(parents=True, exist_ok=True) # 确保基础输出目录存在

        # 调用主处理函数
        processing_result = process_presentation(input_ppt_file, output_base_dir)

        if processing_result:
            print("\n--- 处理成功 ---")
            print(f"已处理 {len(processing_result)} 张幻灯片的数据。")
            print("临时文件保存在以 'temp_' 开头的子目录中，位于:", output_base_dir.resolve())
            print("处理结果详情 (前 5 条):")
            for i, data in enumerate(processing_result[:5]):
                 print(f"  幻灯片 {data['slide_number']}:")
                 print(f"    图片: {data['image_path']}")
                 print(f"    备注: '{data['notes'][:50]}...'")
                 print(f"    音频: {data['audio_path']}")
                 print(f"    时长: {data['audio_duration']:.2f}s")
                 if i == 4 and len(processing_result) > 5:
                     print("    ...")

            # 在这里，下一步就是将 processing_result 传递给视频合成函数
            print("\n下一步: 使用这些数据合成最终视频。")

        else:
            print("\n--- 处理失败 ---")
            print("请检查上面的日志输出获取详细错误信息。")

    logging.info("--- 测试结束 ---")