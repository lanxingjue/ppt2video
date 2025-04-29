# -*- mode: python ; coding: utf-8 -*-

# --- 导入需要的库 ---
from pathlib import Path
import sys # 需要导入 sys 以便在 setup.py 顶部设置递归限制

# --- 提高递归限制 (尝试解决 modulegraph 可能的深度问题) ---
# 这通常是打包复杂应用（尤其涉及科学计算库）时需要的
try:
    current_limit = sys.getrecursionlimit()
    new_limit = max(current_limit, 5000) # 设置一个较高的值，例如 5000
    print(f"[打包配置] 尝试将 Python 递归深度限制设置为: {new_limit}")
    sys.setrecursionlimit(new_limit)
except Exception as e:
    print(f"[打包配置] 警告: 设置递归深度限制失败: {e}")


# --- 定义需要包含的数据文件 ---
# 格式: ('源文件或目录路径', '打包后在应用包内的目标子目录')
# '.' 表示放在打包后的根目录 (通常是 .app/Contents/Resources/)
datas = [
    ('config.ini', '.'),            # 配置文件
    ('cyberpunk_style.qss', '.'),   # 样式表文件
    ('assets/app_icon.icns', '.'),  # 应用图标文件 (也放在根，方便 BUNDLE 查找)
    # 如果你有其他资源文件夹，例如 'fonts' 或 'templates'，像下面这样添加:
    # ('fonts', 'fonts'),
    # ('templates', 'templates'),
]

# --- 定义需要包含的二进制文件 (我们假设 ffmpeg/ffprobe 用户自行安装) ---
# 如果你需要捆绑某些 .dylib 或其他二进制文件，在这里添加
# 格式: ('源文件路径', '打包后在应用包内的目标子目录')
binaries = [
    # 示例：如果你需要捆绑一个特定的 .dylib
    # ('/usr/local/opt/some_lib/lib/libexample.dylib', '.')
]

# --- Analysis 配置: 分析依赖关系 ---
a = Analysis(
    ['gui_app.py'], # 你的主应用程序脚本
    pathex=['.'],   # 将当前目录加入 Python 模块搜索路径
    binaries=binaries,
    datas=datas,
    hiddenimports=[
        # --- 明确告知 PyInstaller 需要包含的核心第三方库 ---
        'PIL',                # Pillow
        'pptx',               # python-pptx (及其依赖 lxml 通常会被自动处理)
        'moviepy',            # MoviePy
        'moviepy.editor',     # MoviePy 常用的编辑器模块
        'stable_whisper',     # stable-ts 包导入名
        'torch',              # stable_whisper 依赖
        'torchaudio',         # stable_whisper 依赖
        'edge_tts',           # Edge TTS
        'pdf2image',          # PDF 转图片
        'mutagen',            # 音频元数据
        'aiohttp',            # edge_tts 依赖
        'charset_normalizer', # aiohttp/requests 依赖
        'certifi',            # ssl/https 依赖, aiohttp 可能需要
        'opencc',           # 如果你安装并使用了 opencc-python-reimplemented

        # --- PyQt6 框架 ---
        # 通常能自动检测，但显式包含更保险
        'PyQt6',
        'PyQt6.QtCore',
        'PyQt6.QtGui',
        'PyQt6.QtWidgets',
        'PyQt6.QtMultimedia',
        'PyQt6.sip',

        # --- 你自己的项目模块 ---
        'tts_manager_edge',
        'ppt_processor',
        'video_synthesizer',
        'main_controller',          # gui_app.py 导入了它
        'ppt_exporter_libreoffice', # ppt_processor.py 导入了它

        # --- 标准库中 PyInstaller 可能遗漏的模块 ---
        'logging', 'configparser', 'platform', 'shutil', 'tempfile', 'uuid',
        'subprocess', 'shlex', 'wave', 'contextlib', 'json', 'asyncio',
        'importlib.metadata', # pkg_resources / setuptools 可能需要
        'pkg_resources',      # 某些旧库或 hook 可能需要
        'pkg_resources.py2_warn', # 有时需要
    ],
    hookspath=[],           # 自定义 hook 路径 (通常不需要)
    hooksconfig={},         # hook 配置
    runtime_hooks=[],       # 运行时 hook (应用启动时执行的代码)
    excludes=[
        # --- 明确排除不需要的模块，减小体积 ---
        'tkinter',                # GUI 应用通常不需要 Tkinter
        'unittest',               # 测试框架
        '_pytest',                # 测试框架
        'pytest',                 # 测试框架
        'numpy.testing',          # numpy 的测试部分
        'scipy',                  # 大型科学计算库 (除非确实需要)
        'pandas',                 # 数据分析库 (除非确实需要)
        'matplotlib',             # 绘图库 (除非确实需要)
        'pyttsx3',                # 确认不再使用
        'FixTk', 'tcl', 'tk', '_tkinter', # 其他 Tkinter 相关
        'qt5', 'PySide2', 'PySide6', # 其他 Qt 绑定
        'ppt_exporter_win',       # Windows 专用导出模块
        'voice_selector_gui',     # 未使用的 GUI 文件
        # 注意: 不要排除 numpy，因为 torch/stable_whisper 依赖它
    ],
    noarchive=False,        # 不将 Python 库打包到 C 存档中 (通常保持 False)
    optimize=1,             # Python 字节码优化级别 (0, 1, or 2)
)

# --- 创建 PYZ (Python 库压缩包) ---
pyz = PYZ(a.pure)

# --- 创建 EXE (核心可执行文件) ---
exe = EXE(
    pyz,
    a.scripts,             # 主脚本
    a.binaries,            # 包含的二进制文件
    a.datas,               # 包含的数据文件
    [],                    # 传递给 bootloader 的参数 (通常为空)
    name='PPT视频转换工具',    # 在 .app 包内部的可执行文件名称
    debug=False,           # 是否启用调试模式 (False 用于发布)
    bootloader_ignore_signals=False,
    strip=False,           # 是否移除调试符号 (False 通常更安全，True 可减小体积)
    upx=False,             # 是否使用 UPX 压缩可执行文件 (不推荐，可能导致问题)
    console=False,         # 对于 GUI 应用设为 False (不在终端显示输出)
    disable_windowed_traceback=False,
    argv_emulation=True,   # 在 macOS 上模拟命令行参数给 GUI 应用
    target_arch=None,      # None 会自动检测当前构建架构 (例如 arm64)
    codesign_identity=None, # 代码签名标识 (需要 Apple Developer ID)
    entitlements_file=None, # 应用权限文件 (沙盒等需要)
)

# --- 创建 BUNDLE (macOS .app 包) ---
app = BUNDLE(
    exe,
    name='PPT视频转换工具.app', # 最终生成的 .app 包名称
    icon='assets/app_icon.icns', # 指定图标文件路径 (相对于 .spec 文件)
    bundle_identifier='com.yourdomain.ppt2video', # **请修改为你的唯一标识符** (例如 com.mycompany.ppttool)
    info_plist={ # --- 配置 Info.plist 文件 ---
        'NSPrincipalClass': 'NSApplication', # macOS 应用必须
        'NSAppleScriptEnabled': False,       # 通常禁用 AppleScript
        'LSMinimumSystemVersion': '10.15',   # 设置支持的最低 macOS 版本 (例如 Catalina)
        'CFBundleShortVersionString': '0.6.0', # **修改为你的应用版本号**
        'CFBundleVersion': '0.6.0',           # **修改为你的应用构建版本号**
        'CFBundleDisplayName': 'PPT视频转换工具', # 应用显示名称
        'CFBundleName': 'PPT视频转换工具',       # 应用内部名称
        'NSHumanReadableCopyright': 'Copyright © 2024 Your Name/Company. All rights reserved.' # **修改版权信息**
        # 你可以根据需要添加更多 Info.plist键值对，例如文件关联等
    }
)