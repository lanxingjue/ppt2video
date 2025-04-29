# setup.py (极度简化版 - 用于调试 zlib 错误)
import sys
try:
    current_limit = sys.getrecursionlimit()
    new_limit = max(current_limit, 3000)
    print(f"[调试] 尝试将递归深度限制设置为: {new_limit}")
    sys.setrecursionlimit(new_limit)
except Exception as e:
    print(f"[调试] 警告: 设置递归深度限制失败: {e}")

from setuptools import setup
import os
from pathlib import Path

APP_NAME = "PPT视频转换工具_Test"
MAIN_SCRIPT = "gui_app.py"
VERSION = "0.0.1.dev0" # <--- 修改这里，使用符合 PEP 440 的格式
ICON_FILE = "assets/app_icon.icns"

# ... (DATA_FILES 和 OPTIONS 保持极简状态) ...
DATA_FILES = [
    ('', ['config.ini', 'cyberpunk_style.qss']),
]

OPTIONS = {
    'argv_emulation': True,
    'iconfile': ICON_FILE if Path(ICON_FILE).exists() else None,
    'plist': {
        # ... (plist 内容) ...
    },
    'packages': ['PyQt6'],
    'includes': [
        'sys', 'os', 'logging', 'pathlib', 'configparser', 'shutil', 'time', 'subprocess', 'json', 're', 'asyncio',
        'PyQt6.QtCore', 'PyQt6.QtGui', 'PyQt6.QtWidgets', 'PyQt6.QtMultimedia',
        'zlib', # <--- 尝试在这里也包含 zlib
    ],
    'excludes': [
        # ... (之前的 excludes 列表) ...
        'PIL', 'pptx', 'moviepy', 'stable_whisper', 'edge_tts', 'pdf2image',
        'mutagen', 'aiohttp', 'charset_normalizer', 'opencc', 'pyttsx3',
        'tts_manager_edge', 'ppt_processor', 'video_synthesizer',
        'main_controller',
        'ppt_exporter_libreoffice',
        'ppt_exporter_win', 'voice_selector_gui',
        'tkinter', 'unittest',
        'numpy', 'scipy',
        # 'zlib', # 既然 includes 里加了，excludes 里可以先注释掉试试
    ],
    'optimize': 1,
}


print("[调试] 使用极度简化的 setup.py 配置 (已修正版本号)...")

setup(
    name=APP_NAME,
    version=VERSION, # 使用修正后的 VERSION
    app=[MAIN_SCRIPT],
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)

print("\n--- py2app 极简测试打包完成 ---")