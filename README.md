# PPT to Video Converter (赛博朋克版)

本项目是一个基于 Python 的桌面应用程序，旨在将 PowerPoint (.pptx) 文件自动化地转换为带有同步语音旁白、字幕和（未来可能包含）转场效果的视频文件。当前版本具有赛博朋克风格的 PyQt GUI。

## 主要功能

*   **PPT 页面转图片:** (当前版本依赖 Windows + Microsoft Office) 自动将 PPTX 文件的每一页导出为 PNG 图片。
*   **提取演讲稿:** 读取并解析 PPTX 文件中每页幻灯片的演讲者备注。
*   **文本转语音 (TTS):** 使用 `pyttsx3` (或其他可配置引擎) 将演讲稿转换为语音片段。
*   **语音转文本 (ASR) / 字幕生成:** 使用 `stable-ts` (基于 OpenAI Whisper) 将生成的语音转换为带时间戳的 SRT 字幕文件。
*   **视频合成:** 使用 **FFmpeg** (通过 Python `subprocess` 调用) 将图片序列、音频片段拼接成视频。
*   **字幕叠加:** 使用 **FFmpeg** 将生成的 SRT 字幕硬编码到最终视频中。
*   **图形用户界面 (GUI):** 提供基于 PyQt (PyQt6/PyQt5) 的用户界面，包含文件/目录选择、开始按钮、进度条和日志显示。
*   **赛博朋克风格:** 应用了 QSS 样式表，提供深色主题和霓虹效果。
*   **配置化:** 通过 `config.ini` 文件可以调整部分参数，如输出目录、TTS 速率、Whisper 模型、FFmpeg 路径等。
*   **依赖检查:** 启动时检查关键外部依赖 (如 FFmpeg) 是否可用。
*   **打包支持:** 提供 `PyInstaller` 的 `.spec` 文件配置，用于将应用打包成可执行文件 (包含捆绑的 FFmpeg)。

## 技术栈

*   **主要语言:** Python 3
*   **GUI 框架:** PyQt6 (或 PyQt5)
*   **PPT 解析:** `python-pptx`
*   **TTS:** `pyttsx3` (或其他通过配置指定的引擎)
*   **ASR (字幕):** `stable-ts` (依赖 `openai-whisper`, `ffmpeg`)
*   **图像处理:** `Pillow` (用于 GUI 和可能的图像操作)
*   **视频/音频处理核心:** **FFmpeg** (通过 `subprocess` 调用)
*   **打包:** `PyInstaller`
*   **外部依赖:**
    *   **FFmpeg:** 必须安装或随应用捆绑。
    *   **Microsoft PowerPoint (Windows):** 当前版本的幻灯片导出功能**必需**。

## 安装与运行

**1. 环境准备:**

*   **Python:** 安装 Python 3.8 或更高版本。建议使用虚拟环境。
*   **FFmpeg:**
    *   从 [ffmpeg.org](https://ffmpeg.org/download.html) 下载适用于你操作系统的 FFmpeg。
    *   解压并将 `bin` 目录添加到系统 PATH 环境变量，**或者**将 `ffmpeg.exe` (及相关 DLL) 放入项目根目录下的 `vendor` 文件夹中 (推荐用于打包)。
    *   在终端运行 `ffmpeg -version` 确认安装成功。
*   **Microsoft PowerPoint (仅 Windows):** 确保已安装 Microsoft Office 套件，包含 PowerPoint。

**2. 克隆或下载项目:**

```bash
git clone <your-repo-url> # 或者下载 ZIP 解压
cd ppt2video # 进入项目目录

** 3. 创建虚拟环境 (推荐):

python -m venv .venv
# Windows
.\.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

4. 安装 Python 依赖:
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
# 或者使用 uv (如果安装了)
# uv pip install -r requirements.txt

5. 配置 config.ini:
打开 config.ini 文件。
(可选) 根据需要修改 [通用设置], [路径设置], [音频设置], [视频设置] 下的参数。特别是 ffmpeg_path，如果 FFmpeg 不在系统 PATH 中且未使用 vendor 文件夹捆绑，需要指定完整路径。

6. 运行 GUI 应用:
python gui_app.py







你说得对，直接运行 ppt_exporter_libreoffice.py 并看到这个警告非常有意义，而且你发现问题是非常正确的！

这个警告的意义：

测试环境: 直接运行这个文件，实际上是在测试这个独立的导出模块是否能在你当前的 macOS 环境下正常工作（或者至少是找到它需要的工具）。

核心问题暴露: 这个警告 未能找到 soffice 直接告诉你，这个脚本找不到 LibreOffice 的核心命令行工具 soffice。soffice 是用来执行将 PPTX 转换为 PDF 这关键一步的程序。

预示失败: 如果现在找不到 soffice，那么当主程序 (ppt_processor.py) 调用 export_slides_with_libreoffice 函数时，它同样会失败，导致整个 PPT 转图片的过程无法进行。

为什么会找不到 soffice？

在 macOS 上，最常见的原因是：

未安装 LibreOffice: 你可能还没有在你的 Mac 上安装 LibreOffice。

LibreOffice 未添加到系统 PATH: 即使安装了 LibreOffice，它的命令行工具 soffice 默认可能不在系统的 PATH 环境变量中。这意味着当你在终端直接输入 soffice 或者脚本尝试调用它时，系统不知道去哪里找这个程序。

如何解决这个问题（macOS）:

确认/安装 LibreOffice:

前往 LibreOffice 官方网站 下载适用于 macOS 的版本并安装。

将 soffice 添加到 PATH (推荐方式之一):

找到 soffice 路径: LibreOffice 安装后，soffice 通常位于 /Applications/LibreOffice.app/Contents/MacOS/soffice。你可以在 Finder 中右键点击 LibreOffice.app -> 显示包内容 -> Contents -> MacOS 来确认。

添加到 PATH: 你需要编辑你的 shell 配置文件 (通常是 ~/.zshrc 对于较新的 macOS，或者 ~/.bash_profile 或 ~/.bashrc 对于旧版本)。打开终端，输入以下命令之一来编辑文件 (以 zshrc 为例):

open ~/.zshrc
# 或者使用 nano 编辑器:
# nano ~/.zshrc


在文件末尾添加一行（将下面的路径替换为你实际的 soffice 父目录路径）：

export PATH="/Applications/LibreOffice.app/Contents/MacOS:$PATH"
IGNORE_WHEN_COPYING_START
content_copy
download
Use code with caution.
Bash
IGNORE_WHEN_COPYING_END

保存文件并关闭编辑器。

让更改生效: 在终端运行 source ~/.zshrc (或者关闭并重新打开终端)。

验证: 在终端输入 soffice --version，如果看到版本信息而不是 "command not found"，说明添加成功。

通过 config.ini 指定路径 (替代方式):

如果不想修改系统 PATH，你可以编辑项目中的 config.ini 文件。

找到 [Paths] 部分，添加或修改 libreoffice_path，指向 soffice 的完整路径：

[Paths]
ffmpeg_path = ffmpeg
ffprobe_path = ffprobe
libreoffice_path = /Applications/LibreOffice.app/Contents/MacOS/soffice # <--- 添加或修改这里
# poppler_path = /opt/homebrew/opt/poppler/bin # 如果 poppler 不在 PATH 也需要配置
IGNORE_WHEN_COPYING_START
content_copy
download
Use code with caution.
Ini
IGNORE_WHEN_COPYING_END

保存 config.ini 文件。脚本在运行时会读取这个配置。





还要brew install poppler
安装liboffice。
注意：
ffmpeg  的安装  ： brew install  ffmpeg
liboffice 安装，以及环境变量的配置


# 打包
mac上采用  py2app打包
安装
pip install py2app
使用setup.py,
总是报错zlib没有文件夹


采用pyinstall来安装
pip install --upgrade pyinstaller
2. 实现 resource_path 辅助函数: