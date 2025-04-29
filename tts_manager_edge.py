# tts_manager_edge.py
import logging
import asyncio # Edge TTS 是异步库
import edge_tts
import tempfile
from pathlib import Path
import os
import wave
import contextlib

# --- 日志记录配置 ---
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - [%(module)s:%(funcName)s] - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

# --- 精选的 Edge TTS 语音列表 ---
# 格式: 'Voice ID': {'name': '显示名称', 'lang': '语言代码', 'gender': '性别'}
# Voice ID 可以通过命令 `edge-tts --list-voices` 查看
# 这里选择一些常见且质量较好的中英文语音
KNOWN_EDGE_VOICES = {
    # --- 中文 (普通话) ---
    "zh-CN-XiaoxiaoNeural": {"name": "晓晓 (女声, 推荐)", "lang": "zh-CN", "gender": "Female"},
    "zh-CN-YunxiNeural": {"name": "云希 (男声, 推荐)", "lang": "zh-CN", "gender": "Male"},
    "zh-CN-YunjianNeural": {"name": "云健 (男声, 沉稳)", "lang": "zh-CN", "gender": "Male"},
    "zh-CN-XiaoyiNeural": {"name": "晓伊 (女声, 温柔)", "lang": "zh-CN", "gender": "Female"},
    "zh-CN-liaoning-XiaobeiNeural": {"name": "辽宁小北 (女声, 东北)", "lang": "zh-CN-liaoning", "gender": "Female"}, # 地方口音示例
    "zh-CN-shaanxi-XiaoniNeural": {"name": "陕西小妮 (女声, 陕西)", "lang": "zh-CN-shaanxi", "gender": "Female"}, # 地方口音示例
    # --- 英文 (美国) ---
    "en-US-JennyNeural": {"name": "Jenny (女声, 推荐)", "lang": "en-US", "gender": "Female"},
    "en-US-GuyNeural": {"name": "Guy (男声, 推荐)", "lang": "en-US", "gender": "Male"},
    "en-US-AriaNeural": {"name": "Aria (女声)", "lang": "en-US", "gender": "Female"},
    "en-US-DavisNeural": {"name": "Davis (男声)", "lang": "en-US", "gender": "Male"},
    "en-US-SaraNeural": {"name": "Sara (女声, 清晰)", "lang": "en-US", "gender": "Female"},
    "en-US-ChristopherNeural": {"name": "Christopher (男声, 成熟)", "lang": "en-US", "gender": "Male"},
    # --- 英文 (英国) ---
    "en-GB-LibbyNeural": {"name": "Libby (女声, UK)", "lang": "en-GB", "gender": "Female"},
    "en-GB-RyanNeural": {"name": "Ryan (男声, UK)", "lang": "en-GB", "gender": "Male"},
    "en-GB-SoniaNeural": {"name": "Sonia (女声, UK)", "lang": "en-GB", "gender": "Female"},
    # --- 英文 (澳大利亚) ---
    "en-AU-NatashaNeural": {"name": "Natasha (女声, AU)", "lang": "en-AU", "gender": "Female"},
    "en-AU-WilliamNeural": {"name": "William (男声, AU)", "lang": "en-AU", "gender": "Male"},
    # 可以根据需要添加更多，例如其他语言或风格
}

# --- 异步执行帮助函数 ---
def run_async_in_sync(async_func):
    """
    在同步代码中安全地运行异步函数。
    为每个调用创建一个新的事件循环，避免在已有循环（如 PyQt 的）中运行。
    """
    try:
        # 尝试获取现有事件循环 (通常在主线程会失败)
        loop = asyncio.get_event_loop()
        if loop.is_running():
            # 如果已有循环在运行（比如在异步框架内），需要不同处理
            # 但在典型的同步脚本或 PyQt 回调中，这通常不是问题
            # 为了简单起见，我们总是创建一个新循环并在其中运行
            raise RuntimeError("Existing loop is running, creating new one.")
    except RuntimeError: # 通常意味着没有当前事件循环
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        is_new_loop = True
    else:
        is_new_loop = False

    try:
        result = loop.run_until_complete(async_func)
        return result
    finally:
        if is_new_loop:
            loop.close()
            # 重置事件循环策略，以便下次调用能正确创建
            asyncio.set_event_loop_policy(None)


# --- TTS 功能函数 ---
def get_available_voices() -> list[dict]:
    """返回预定义的 Edge TTS 语音列表。"""
    logging.info("获取预定义的 Edge TTS 语音列表。")
    voice_list = []
    for voice_id, details in KNOWN_EDGE_VOICES.items():
        voice_info = details.copy()
        voice_info['id'] = voice_id
        voice_list.append(voice_info)
    # 按显示名称排序
    voice_list.sort(key=lambda x: x.get('name', ''))
    return voice_list

async def _synthesize_edge_audio(voice_id: str, text: str, output_path: Path, rate_str: str = "+0%"): # <<< 移除 pitch_str 参数
    """异步执行 Edge TTS 合成并保存到文件。"""
    logging.debug(f"开始异步合成: Voice='{voice_id}', Rate='{rate_str}', Text='{text[:30]}...'")
    # !!! CHANGE: 不再传递 pitch 参数 !!!
    communicate = edge_tts.Communicate(text, voice_id, rate=rate_str)
    # !!! --------------------------- !!!
    await communicate.save(str(output_path))
    logging.debug(f"异步合成完成，已保存到: {output_path.name}")


def generate_preview_audio(voice_id: str, text: str | None = None) -> str | None:
    """
    使用指定的 Edge TTS voice_id 生成一小段预览音频 (MP3)。

    Args:
        voice_id: 要使用的语音 ID (例如 'zh-CN-XiaoxiaoNeural')。
        text: (可选) 要转换为语音的示例文本。如果为 None，会根据语音语言选择默认文本。

    Returns:
        成功生成的临时音频文件 (mp3) 的绝对路径。如果失败则返回 None。
        注意：调用者负责在使用后删除此临时文件。
    """
    logging.info(f"请求 Edge TTS 预览: Voice ID='{voice_id}'")
    if voice_id not in KNOWN_EDGE_VOICES:
        logging.error(f"无效的语音 ID: '{voice_id}'")
        return None

    # 根据语言选择默认预览文本
    if text is None:
        lang_prefix = KNOWN_EDGE_VOICES[voice_id]['lang'].split('-')[0].lower()
        if lang_prefix == 'zh':
            text = "你好，这是一个使用微软 Edge 语音合成的试听示例。"
        else: # 默认为英文
            text = "Hello, this is an audio preview using Microsoft Edge speech synthesis."

    temp_file_path = None
    try:
        # Edge TTS 通常输出 MP3
        with tempfile.NamedTemporaryFile(suffix=".mp3", delete=False) as tmp_f:
            temp_file_path = Path(tmp_f.name)
        logging.info(f"创建临时预览文件: {temp_file_path}")
        
        # --- 运行异步合成 (调用修改后的 _synthesize_edge_audio) ---
        async_task = _synthesize_edge_audio(voice_id, text, temp_file_path) # 不再传 pitch
        run_async_in_sync(async_task)
        # --------------------

        if temp_file_path.exists() and temp_file_path.stat().st_size > 100:
            logging.info(f"Edge TTS 预览音频生成成功: {temp_file_path}")
            return str(temp_file_path.resolve())
        else:
            logging.error("Edge TTS 未能成功生成预览音频文件或文件为空。")
            if temp_file_path.exists(): os.remove(temp_file_path)
            return None

    # !!! CHANGE: 修改异常捕获 !!!
    # except edge_tts.NoAudioReceived: # <--- 移除这一行
    #     logging.error("Edge TTS 错误：未能从服务器接收到音频数据...")
    #     if temp_file_path and temp_file_path.exists(): os.remove(temp_file_path)
    #     return None
    # 改为捕获更通用的异常，或者查找特定网络错误
    except ConnectionError as e: # 捕获网络连接错误
         logging.error(f"网络连接错误: {e}")
         if temp_file_path and temp_file_path.exists(): os.remove(temp_file_path)
         return None
    except TimeoutError as e: # 捕获超时错误
         logging.error(f"请求超时: {e}")
         if temp_file_path and temp_file_path.exists(): os.remove(temp_file_path)
         return None
    except Exception as e: # 保留通用异常捕获
        # 检查错误消息是否指示没有音频数据（这只是一个猜测性的检查）
        if "no audio data received" in str(e).lower():
             logging.error(f"Edge TTS 错误：未能从服务器接收到音频数据 (捕获自通用异常: {e})")
        else:
             logging.error(f"生成 Edge TTS 预览音频时发生错误: {e}", exc_info=True)

        if temp_file_path and temp_file_path.exists():
            try: os.remove(temp_file_path)
            except OSError: pass
        return None
    # !!! ----------------------- !!!


def generate_segment_audio(voice_id: str, text: str, output_path: Path, rate: int = 100) -> bool: # <<< 移除 pitch 参数
    """
    为演讲稿的一个片段生成音频文件 (MP3)。

    Args:
        voice_id: 要使用的语音 ID。
        text: 要转换的文本片段。
        output_path: 要保存的音频文件路径 (Path 对象, e.g., segment_1.mp3)。
        rate: 语速百分比 (100 表示正常，范围通常 50-200)。
        pitch: 音调百分比 (100 表示正常)。

    Returns:
        True 如果成功生成音频文件, False 如果失败。
    """
    logging.debug(f"请求 Edge TTS 片段音频: Voice='{voice_id}', Rate={rate}%, Output='{output_path.name}', Text='{text[:30]}...'")
    if voice_id not in KNOWN_EDGE_VOICES:
        logging.error(f"无效的语音 ID: '{voice_id}'")
        return False
    if not text or text.isspace():
        logging.warning(f"文本片段为空，跳过 TTS: {output_path.name}")
        return False # 不生成文件算作失败

    # 将百分比转换为 Edge TTS 需要的格式 (+x% 或 -x%)
    rate_str = f"{rate-100:+d}%"
    # pitch_str 不再需要


    try:
        # 确保父目录存在
        output_path.parent.mkdir(parents=True, exist_ok=True)

        # --- 运行异步合成 (调用修改后的 _synthesize_edge_audio) ---
        async_task = _synthesize_edge_audio(voice_id, text, output_path, rate_str=rate_str) # 不再传 pitch
        run_async_in_sync(async_task)
        # --------------------

        if output_path.exists() and output_path.stat().st_size > 100:
            logging.info(f"  Edge TTS 片段音频生成成功: {output_path.name}")
            return True
        else:
            logging.error(f"  Edge TTS 未能成功生成片段音频文件或文件为空: {output_path.name}")
            if output_path.exists(): output_path.unlink(missing_ok=True)
            return False
    # !!! CHANGE: 修改异常捕获 !!!
    # except edge_tts.NoAudioReceived: # <--- 移除这一行
    #     logging.error(f"Edge TTS 错误：未能从服务器接收到片段 '{output_path.name}' 的音频数据。")
    #     if output_path.exists(): output_path.unlink(missing_ok=True)
    #     return False
    except ConnectionError as e: # 捕获网络连接错误
        logging.error(f"片段 '{output_path.name}' 网络连接错误: {e}")
        if output_path.exists(): output_path.unlink(missing_ok=True)
        return False
    except TimeoutError as e: # 捕获超时错误
        logging.error(f"片段 '{output_path.name}' 请求超时: {e}")
        if output_path.exists(): output_path.unlink(missing_ok=True)
        return False
    except Exception as e: # 保留通用异常捕获
        if "no audio data received" in str(e).lower():
             logging.error(f"Edge TTS 错误：未能从服务器接收到片段 '{output_path.name}' 的音频数据 (捕获自通用异常: {e})")
        else:
            logging.error(f"生成 Edge TTS 片段音频 '{output_path.name}' 时发生错误: {e}", exc_info=True)

        if output_path.exists(): output_path.unlink(missing_ok=True)
        return False
    # !!! ----------------------- !!!

# --- WAV 时长获取 (如果需要，但 Edge TTS 输出 MP3) ---
# 注意：准确获取 MP3 时长比 WAV 复杂，可以依赖外部库如 mutagen 或 tinytag
# 或者在视频合成时让 MoviePy/FFmpeg 处理 MP3 文件
def get_mp3_duration(filepath: Path) -> float:
    """尝试使用 mutagen 获取 MP3 时长 (如果安装了 mutagen)。"""
    if not filepath.is_file(): return 0.0
    try:
        from mutagen.mp3 import MP3
        audio = MP3(str(filepath))
        return audio.info.length
    except ImportError:
        logging.warning("无法导入 'mutagen' 库。MP3 时长计算将不准确。请 'pip install mutagen'。")
        # 可以返回一个估算值或 0
        return 0.0 # 返回 0 可能导致视频合成问题
    except Exception as e:
        logging.error(f"使用 mutagen 获取 MP3 时长失败 {filepath}: {e}")
        return 0.0

# --- 命令行测试 ---
if __name__ == "__main__":
    print("--- Edge TTS 管理器测试 (需要网络连接) ---")

    voices = get_available_voices()
    print(f"\n找到 {len(voices)} 个预定义的 Edge TTS 语音:")
    for i, v in enumerate(voices):
        print(f"  [{i+1}] {v['name']} (ID: {v['id']}, Lang: {v['lang']}, Gender: {v['gender']})")

    if not voices: exit()

    # --- 测试预览 (中文) ---
    test_voice_zh = "zh-CN-XiaoxiaoNeural"
    print(f"\n测试中文预览: {test_voice_zh}")
    preview_file_zh = generate_preview_audio(test_voice_zh)
    if preview_file_zh:
        print(f"中文预览 MP3: {preview_file_zh}")
        # 播放和删除逻辑与之前类似...
        input("按 Enter 删除中文预览文件...")
        try: os.remove(preview_file_zh)
        except OSError as e: print(f"删除失败: {e}")
    else:
        print("生成中文预览失败。")

    # --- 测试预览 (英文) ---
    test_voice_en = "en-US-JennyNeural"
    print(f"\n测试英文预览: {test_voice_en}")
    preview_file_en = generate_preview_audio(test_voice_en)
    if preview_file_en:
        print(f"英文预览 MP3: {preview_file_en}")
        # 播放和删除逻辑与之前类似...
        input("按 Enter 删除英文预览文件...")
        try: os.remove(preview_file_en)
        except OSError as e: print(f"删除失败: {e}")
    else:
        print("生成英文预览失败。")

    # --- 测试片段生成 ---
    print(f"\n测试中文片段生成: {test_voice_zh}")
    segment_text_zh = "这是主要的转换流程中会用到的一段中文旁白。"
    segment_output_zh = Path("./edge_test_segment_zh.mp3")
    success_zh = generate_segment_audio(test_voice_zh, segment_text_zh, segment_output_zh, rate=110) # 稍微快一点
    if success_zh:
        print(f"中文片段 MP3 已生成: {segment_output_zh.resolve()}")
        duration = get_mp3_duration(segment_output_zh) # 尝试获取时长
        print(f"  估算时长 (需要 mutagen): {duration:.2f} 秒")
        input("按 Enter 删除中文片段文件...")
        try: segment_output_zh.unlink()
        except OSError as e: print(f"删除失败: {e}")
    else:
        print("生成中文片段失败。")