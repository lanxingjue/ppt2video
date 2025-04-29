# -*- mode: python ; coding: utf-8 -*-

import sys
import os
from pathlib import Path
import platform
import opencc # 需要导入才能找到数据文件

block_cipher = None

# --- Helper function to find OpenCC data directory ---
def get_opencc_datadirs():
    """Finds the path to the OpenCC dictionary data."""
    try:
        # opencc.__path__ MIGHT give the library location, data is often relative
        oc_path = Path(opencc.__path__[0])
        # Common locations for data relative to the library path
        possible_data_paths = [
            oc_path / 'data',       # Standard location?
            oc_path.parent / 'opencc_data', # Some installations might put it here
            oc_path.parent / 'data' / 'opencc', # Another possibility
            Path(sys.prefix) / 'share' / 'opencc', # Linux system install?
            Path(sys.prefix) / 'Lib' / 'site-packages' / 'opencc' / 'data', # Inside site-packages
        ]
        # On Windows, also check relative to sys.executable if in venv
        if platform.system() == "Windows":
            possible_data_paths.append(Path(sys.executable).parent / 'Lib' / 'site-packages' / 'opencc' / 'data')

        for data_path in possible_data_paths:
            if data_path.is_dir() and (data_path / 't2s.json').exists(): # Check for a known file
                print(f"Found OpenCC data at: {data_path}")
                # We need to return a list of tuples for PyInstaller datas
                # Copy the entire data directory content into 'opencc' subdir in bundle
                data_files = []
                for root, _, files in os.walk(data_path):
                    for file in files:
                         if file.lower().endswith(('.json', '.txt', '.ocd2')): # Include relevant extensions
                            source_file = Path(root) / file
                            # Destination path relative to the data dir's parent in bundle
                            dest_rel_path = Path(root).relative_to(data_path.parent) / file
                            data_files.append((str(source_file), str(Path('opencc_data') / dest_rel_path))) # Put in 'opencc_data'
                if data_files:
                    return data_files
    except Exception as e:
        print(f"Error finding OpenCC data directory: {e}")
    print("Warning: Could not automatically find OpenCC data directory!")
    return [] # Return empty list if not found



# --- Analysis Section: Tells PyInstaller what to include ---
a = Analysis(
    ['gui_app.py'], # 你的主脚本
    pathex=[],      # 项目根目录加入 Python 路径
    binaries=[
        # --- 添加捆绑的 ffmpeg 和 ffprobe ---
        # 源路径是相对于 .spec 文件，目标路径是打包后的根目录下的 vendor 目录
        ('vendor/ffmpeg.exe', 'vendor'), # <<< 确保 vendor 目录和 exe 文件存在
        ('vendor/ffprobe.exe', 'vendor')  # <<< 确保 vendor 目录和 exe 文件存在
    ],
    datas=[
        # --- IMPORTANT: Include necessary data files ---
        ('config.ini', '.'),             # Include config.ini in the root directory of the package
        ('cyberpunk_style.qss', '.'),    # Include the QSS file in the root directory
        ('vendor', 'vendor'), # <-- 添加这一
        # If you have custom fonts in a 'fonts' subfolder:
        # ('fonts', 'fonts'),
        # If stable-ts/whisper requires specific model files not automatically found:
        # You might need to locate the model cache dir and include it,
        # or bundle a specific model file. This can be complex.
        # Example (replace with actual whisper model path if needed):
        # ('C:/Users/YourUser/.cache/whisper', '.whisper_cache')
        # --- 添加 OpenCC 数据文件 ---
        *get_opencc_datadirs()          # <<< 动态添加 OpenCC 数据
    ],
    # 在 PPT2VideoApp.spec 文件的 Analysis 部分
    hiddenimports=[
        'pyttsx3.drivers',
        'pyttsx3.drivers.sapi5',
        'pyttsx3.drivers.nsss',
        'pyttsx3.drivers.espeak',
        'comtypes', 
        # --- 添加可能需要的隐藏导入 ---
        'stable_whisper',
        'stable_whisper.audio',
        'stable_whisper.result',
        'stable_whisper.text_output',
        'stable_whisper.timing',
        'whisper',
        'whisper.audio', # Whisper 内部模块
        'whisper.model',
        'whisper.tokenizer',
        'whisper.utils',
        'tiktoken_ext.openai_public', # Whisper 依赖
        'tiktoken_ext.claude',      # Whisper 依赖
        'opencc', # 明确包含 opencc
        # 'opencc.clib', # 如果有 C 扩展
        'edge_tts',
        'edge_tts.communicate', # edge-tts 内部
        'edge_tts.submaker',    # edge-tts 内部
        'aiohttp', # edge-tts 依赖
        'asyncio',
        'pkg_resources.py2_warn', # 常见兼容性问题
        'win32com', 'win32com.client', 'pythoncom', # pywin32 相关 (Windows)
        'mutagen', 'mutagen.mp3', 'mutagen.id3', # mutagen 相关
        'soundfile',       # whisper 常用的音频处理库
        'sounddevice',     # 有时也需要
        'tiktoken',        # whisper 的 tokenizer
        'transformers',    # whisper 可能依赖 (尤其是新版本)
        'torch',           # PyTorch 核心 (非常重要)
        'torchaudio',      # PyTorch 音频处理
        'numpy',           # 几乎所有科学计算库都依赖
        # 'pkg_resources', # 有些旧库可能需要
        # --- 添加 PyQt6 相关 ---
        'PyQt6',
        'PyQt6.sip',
        'PyQt6.QtCore',
        'PyQt6.QtGui',
        'PyQt6.QtWidgets',
        # 如果用到其他 PyQt 模块，也可能需要添加
    ],
    hookspath=[],      # Path to custom hook files (advanced)
    hooksconfig={},
    runtime_hooks=[],  # Scripts to run at runtime before main script
    excludes=[],       # List modules to exclude (e.g., tkinter if unused)
    win_no_prefer_redirects=False, # Usually False
    win_private_assemblies=False, # Usually False
    cipher=block_cipher,
    noarchive=False    # Usually False
)

# --- PYC Files --- (Usually default is fine)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# --- Executable Section ---
exe = EXE(
    pyz,
    a.scripts,
    a.binaries, # 包含 ffmpeg/ffprobe
    a.zipfiles,
    a.datas,    # 包含 config, qss, opencc data
    [],
    name='PPT2VideoApp', # <<< 输出的 exe 文件名
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,  # 可以尝试用 UPX 压缩，减小体积，但可能引起杀软误报
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False, # <<< GUI 应用设为 False
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='titile.ico' # <<< 可选：指定图标文件路径
)

coll = COLLECT( # 用于创建单目录模式
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='PPT2VideoApp_dist' # 输出的文件夹名称
)
# --- Bundled App (macOS specific, usually not needed for Windows EXE) ---
# coll = COLLECT(...) # For creating a folder distribution instead of one-file
# app = BUNDLE(...) # For creating .app bundle on macOS