# -*- coding: utf-8 -*-
# ^^^ 确保 Python 能正确处理文件中的中文字符

import sys
import os
import logging
from pathlib import Path
import configparser
import shutil
import time
import subprocess

# --- PyQt6 导入 ---
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLineEdit, QFileDialog, QTextEdit, QProgressBar,
    QLabel, QMessageBox, QGroupBox, QComboBox, QSizePolicy # 增加了 QGroupBox, QComboBox, QSizePolicy
)
from PyQt6.QtCore import QThread, pyqtSignal, Qt, QObject, QUrl, QTimer, pyqtSlot # 增加了 QUrl, QTimer, pyqtSlot
from PyQt6.QtGui import QFont, QFontDatabase # 导入字体相关类
from PyQt6.QtMultimedia import QMediaPlayer, QAudioOutput # 增加了 QMediaPlayer, QAudioOutput

# --- 导入我们后端的函数 ---
# (确保这些文件在同一目录或 Python 路径中)
try:
    # 假设 run_full_process 不再需要，WorkerThread 直接调用处理和合成函数
    from main_controller import check_dependencies, BASE_OUTPUT_DIR as config_base_output_dir
    from ppt_processor import process_presentation # 导入处理函数
    from video_synthesizer import create_video_from_data # 导入合成函数
    import tts_manager_edge as tts_manager # Use alias
except ImportError as e:
    print(f"错误：导入后端或 TTS 模块失败: {e}...")
    sys.exit(1)
except Exception as e:
    print(f"错误：导入模块时发生意外错误: {e}")
    sys.exit(1)


# --- 配置日志记录 ---
log_level = logging.INFO
logging.basicConfig(
    level=log_level,
    format='%(asctime)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

# --- 后台工作线程 ---
class WorkerThread(QThread):
    """在独立线程中运行转换过程，避免 UI 冻结。"""
    log_signal = pyqtSignal(str)
    progress_signal = pyqtSignal(int, str)
    finished_signal = pyqtSignal(bool, str)

    # 修改构造函数以接收 voice_id
    def __init__(self, input_pptx_path, output_dir_path, cleanup_temp, selected_voice_id, process_func, synthesize_func):
        super().__init__()
        self.input_pptx_path = input_pptx_path
        self.output_dir_path = output_dir_path
        self.cleanup_temp = cleanup_temp
        self.selected_voice_id = selected_voice_id # 存储选定的语音 ID
        self._is_running = True
        self.process_presentation = process_func
        self.create_video_from_data = synthesize_func

    def run(self):
        """线程执行的主要工作。"""
        self.log_signal.emit("后台处理线程已启动...")
        self.progress_signal.emit(0, "正在初始化...")
        start_time = time.time()
        temp_run_dir = None
        final_video_path = None
        success = False

        try:
            # --- 1. 准备路径 ---
            # (与之前类似)
            base_output = self.output_dir_path
            base_output.mkdir(parents=True, exist_ok=True)
            final_video_filename = self.input_pptx_path.stem + "_gui_视频.mp4"
            final_video_path = base_output / final_video_filename
            if final_video_path.exists():
                self.log_signal.emit(f"警告：输出文件 {final_video_path.name} 已存在，将被覆盖。")
                try:
                    final_video_path.unlink()
                except OSError as e:
                    raise RuntimeError(f"无法删除已存在的输出文件: {e}")
            self.progress_signal.emit(5, "准备处理演示文稿")

            # --- 2. 处理演示文稿 (调用 ppt_processor) ---
            self.log_signal.emit("阶段 1/3: 正在处理演示文稿 (提取/转图片/生成语音)...")
            # !!! 关键修改：将 selected_voice_id 传递给 process_presentation !!!
            # !!! 你需要确保 ppt_processor.py 中的 process_presentation 函数能接收这个参数 !!!
            processed_data, temp_run_dir = self.process_presentation(
                self.input_pptx_path,
                base_output,
                voice_id=self.selected_voice_id # <---- 传递 Voice ID
            )
            # !!! =========================================================== !!!

            if processed_data is None or temp_run_dir is None:
                raise RuntimeError("演示文稿处理失败或返回无效数据。请检查日志。")
            self.log_signal.emit(f"演示文稿处理完成。临时目录: {temp_run_dir}")
            self.progress_signal.emit(40, f"已处理 {len(processed_data)} 张幻灯片，准备合成视频")

            # --- 3. 合成视频 (调用 video_synthesizer) ---
            self.log_signal.emit("阶段 2/3: 正在合成视频 (拼接/生成/添加字幕)...")
            # create_video_from_data 可能不需要 voice_id，除非它也涉及音频处理
            synthesis_success = self.create_video_from_data(
                processed_data,
                temp_run_dir,
                final_video_path
            )
            if not synthesis_success:
                raise RuntimeError("视频合成失败。请检查日志。")
            self.progress_signal.emit(95, "视频合成完成，正在清理")
            self.log_signal.emit("阶段 3/3: 处理完成。")
            success = True

        except Exception as e:
            error_msg = f"处理过程中发生错误: {e}"
            self.log_signal.emit(f"错误: {error_msg}")
            logging.exception("后台线程捕获到错误:")
            self.finished_signal.emit(False, error_msg)

        finally:
            # --- 清理和发送最终信号 ---
            # (与之前类似)
            if success and final_video_path and final_video_path.exists():
                end_time = time.time()
                duration = end_time - start_time
                success_msg = f"处理成功完成！\n输出文件: {final_video_path}\n总耗时: {duration:.2f} 秒"
                self.log_signal.emit(success_msg)
                self.progress_signal.emit(100, "完成")
                self.finished_signal.emit(True, str(final_video_path))
                if self.cleanup_temp and temp_run_dir and temp_run_dir.exists():
                    self.log_signal.emit(f"正在清理临时目录: {temp_run_dir}")
                    try:
                        shutil.rmtree(temp_run_dir)
                        self.log_signal.emit("临时目录已清理。")
                    except Exception as clean_e:
                        self.log_signal.emit(f"警告：清理临时目录失败: {clean_e}")
            elif not success:
                 if temp_run_dir and temp_run_dir.exists():
                    self.log_signal.emit(f"处理失败。临时文件保留在: {temp_run_dir}")
                 else:
                    self.log_signal.emit("处理失败。")


    def stop(self):
        self._is_running = False
        self.log_signal.emit("收到停止请求... (注意：当前无法强制中断 FFmpeg)")


# --- 日志处理器 ---
class QTextEditLogger(logging.Handler, QObject):
    log_requested = pyqtSignal(str)
    def __init__(self, text_edit_widget):
        logging.Handler.__init__(self)
        QObject.__init__(self)
        self.widget = text_edit_widget
        self.log_requested.connect(self.widget.append)
    def emit(self, record):
        msg = self.format(record)
        self.log_requested.emit(msg)


# --- 主应用程序窗口 ---
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.process_presentation_func = process_presentation
        self.create_video_from_data_func = create_video_from_data

        # --- TTS 相关初始化 ---
         # --- TTS 相关初始化 ---
        self.current_preview_file = None
        self.player = None
        self.audio_output = None

        self.setWindowTitle("PPT 转视频工具 | 赛博朋克版 v0.5 - Edge TTS (需联网)") # 更新标题
        self.setGeometry(100, 100, 800, 650)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout(self.central_widget)
        self.layout.setContentsMargins(20, 20, 20, 20)
        self.layout.setSpacing(15)

        # --- 初始化 TTS UI 部分 ---
        self.init_tts_section() # <--- 新增调用

        # --- 输入文件选择区域 ---
        self.input_layout = QHBoxLayout()
        self.input_label = QLabel("源 PPTX 文件:")
        self.input_path_edit = QLineEdit()
        self.input_path_edit.setPlaceholderText("请选择或拖入 PowerPoint 文件...")
        self.input_browse_button = QPushButton("浏览...")
        self.input_browse_button.setObjectName("browseButton")
        self.input_browse_button.clicked.connect(self.browse_input_file)
        self.input_layout.addWidget(self.input_label)
        self.input_layout.addWidget(self.input_path_edit, 1)
        self.input_layout.addWidget(self.input_browse_button)
        self.layout.addLayout(self.input_layout)

        # --- 输出目录选择区域 ---
        self.output_layout = QHBoxLayout()
        self.output_label = QLabel("视频输出目录:")
        self.output_path_edit = QLineEdit()
        default_output_dir = str(config_base_output_dir.resolve())
        self.output_path_edit.setText(default_output_dir)
        self.output_browse_button = QPushButton("选择...")
        self.output_browse_button.setObjectName("browseButton")
        self.output_browse_button.clicked.connect(self.browse_output_dir)
        self.output_layout.addWidget(self.output_label)
        self.output_layout.addWidget(self.output_path_edit, 1)
        self.output_layout.addWidget(self.output_browse_button)
        self.layout.addLayout(self.output_layout)

        # --- 开始转换按钮 ---
        self.start_button = QPushButton(">>  开始转换  <<")
        self.start_button.setObjectName("startButton")
        self.start_button.clicked.connect(self.start_conversion)
        self.layout.addWidget(self.start_button)

        # --- 进度条 ---
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFormat("")
        self.layout.addWidget(self.progress_bar)

        # --- 状态标签 ---
        self.status_label = QLabel("系统状态：待命")
        self.status_label.setObjectName("statusLabel")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.layout.addWidget(self.status_label)

        # --- 日志输出区域 ---
        self.log_output_edit = QTextEdit()
        self.log_output_edit.setObjectName("logOutput")
        self.log_output_edit.setReadOnly(True)
        self.layout.addWidget(self.log_output_edit, 1)

        # --- 设置日志处理器 ---
        self.log_handler = QTextEditLogger(self.log_output_edit)
        log_format = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', '%H:%M:%S')
        self.log_handler.setFormatter(log_format)
        logging.getLogger().addHandler(self.log_handler)
        self.log_handler.setLevel(logging.INFO)

        self.worker_thread = None
        self._dependencies_checked = False

        # --- 延迟加载 TTS 声音列表 ---
        QTimer.singleShot(100, self.load_voices)


    def init_tts_section(self):
        """初始化 TTS 语音选择相关的 UI 控件"""
        tts_groupbox = QGroupBox("语音设置")
        tts_groupbox.setObjectName("ttsGroup") # 给 GroupBox 也设置对象名
        tts_layout = QVBoxLayout(tts_groupbox)

        select_layout = QHBoxLayout()
        lbl_select = QLabel('选择旁白语音:', self)
        self.cmb_voices = QComboBox(self)
        self.cmb_voices.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.cmb_voices.setToolTip("选择用于生成视频旁白的语音")
        self.cmb_voices.setEnabled(False)
        select_layout.addWidget(lbl_select)
        select_layout.addWidget(self.cmb_voices)
        tts_layout.addLayout(select_layout)

        self.btn_preview = QPushButton('试听选中语音', self)
        self.btn_preview.setToolTip("播放选定语音的简短示例")
        self.btn_preview.setObjectName("previewButton") # 设置对象名
        self.btn_preview.clicked.connect(self.preview_selected_voice)
        self.btn_preview.setEnabled(False)
        tts_layout.addWidget(self.btn_preview)

        self.lbl_tts_status = QLabel('状态: 正在初始化...')
        self.lbl_tts_status.setStyleSheet("color: gray;")
        self.lbl_tts_status.setObjectName("ttsStatusLabel") # 对象名
        tts_layout.addWidget(self.lbl_tts_status)

        # 将 GroupBox 添加到主布局 (这里选择添加到顶部)
        self.layout.insertWidget(0, tts_groupbox) # 插入到索引 0 的位置

        # --- 初始化 QMediaPlayer ---
        self.player = QMediaPlayer(self)
        self.audio_output = QAudioOutput(self)
        self.player.setAudioOutput(self.audio_output)
        self.player.mediaStatusChanged.connect(self.handle_media_status)
        self.player.errorOccurred.connect(self.handle_player_error)

    def load_voices(self):
        """加载可用语音到下拉框 (现在从 tts_manager_edge 获取)"""
        self.lbl_tts_status.setText('状态: 正在加载可用语音列表...')
        self.lbl_tts_status.setStyleSheet("color: blue;")
        QApplication.processEvents()

        # !!! CHANGE: Calls the new manager's function !!!
        voices = tts_manager.get_available_voices()
        # !!! --------------------------------------- !!!
        self.cmb_voices.clear()

        if not voices:
            self.lbl_tts_status.setText('状态: 未加载到语音定义。请检查 tts_manager_edge.py。')
            self.lbl_tts_status.setStyleSheet("color: red;")
            QMessageBox.warning(self, "无可用语音", "未能加载预定义的 Edge TTS 语音列表。")
            self.cmb_voices.setEnabled(False)
            self.btn_preview.setEnabled(False)
            return

        logging.info(f"找到 {len(voices)} 个预定义 Edge TTS 语音，正在填充下拉框...")
        for voice in voices:
            # 显示名称和语言/性别等信息
            display_text = f"{voice.get('name', '未知')} ({voice.get('gender','?')}, {voice.get('lang','?')})"
            self.cmb_voices.addItem(display_text, userData=voice.get('id'))
        logging.info("下拉框填充完毕。")

        self.cmb_voices.setEnabled(True)
        self.btn_preview.setEnabled(True)
        self.lbl_tts_status.setText('状态: 语音加载完成 (需要网络连接进行合成)。')
        self.lbl_tts_status.setStyleSheet("color: green;")


    def get_selected_voice_id(self) -> str | None:
        """获取当前下拉框中选定的语音 ID"""
        # (No changes needed)
        current_index = self.cmb_voices.currentIndex()
        if current_index < 0: return None
        return self.cmb_voices.itemData(current_index)

    @pyqtSlot()
    def preview_selected_voice(self):
        """处理“试听”按钮点击事件 (现在调用 tts_manager_edge)"""
        selected_voice_id = self.get_selected_voice_id()
        if not selected_voice_id:
            QMessageBox.information(self, "提示", "请先选择一个语音。")
            return

        # --- 增加网络检查提示 ---
        reply = QMessageBox.question(self, "网络确认",
                                     "试听和生成语音需要网络连接到微软 Edge TTS 服务。\n请确保您的网络连接正常。\n\n是否继续？",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.Cancel,
                                     QMessageBox.StandardButton.Yes)
        if reply == QMessageBox.StandardButton.Cancel:
            return
        # --- ----------------- ---

        self.cleanup_preview_file()

        self.lbl_tts_status.setText(f'状态: 正在生成 "{self.cmb_voices.currentText()}" 的试听音频 (联网)...')
        self.lbl_tts_status.setStyleSheet("color: blue;")
        self.btn_preview.setEnabled(False)
        QApplication.processEvents()

        # !!! CHANGE: Calls the new manager's function !!!
        preview_file_path = tts_manager.generate_preview_audio(selected_voice_id)
        # !!! --------------------------------------- !!!
        self.btn_preview.setEnabled(True)

        if preview_file_path:
            self.current_preview_file = preview_file_path
            logging.info(f"Edge TTS 试听音频已生成: {preview_file_path}")
            self.lbl_tts_status.setText('状态: 正在准备播放试听音频...')
            self.lbl_tts_status.setStyleSheet("color: purple;")
            # QMediaPlayer 通常能播放 MP3
            media_url = QUrl.fromLocalFile(preview_file_path)
            self.player.setSource(media_url)
            self.player.play()
        else:
            self.current_preview_file = None
            self.lbl_tts_status.setText('状态: 生成试听音频失败 (请检查网络或日志)。')
            self.lbl_tts_status.setStyleSheet("color: red;")
            QMessageBox.critical(self, "错误", "生成试听音频失败。\n可能原因：\n- 网络连接问题。\n- Edge TTS 服务暂时不可用。\n- 文本包含不支持的字符。\n请检查日志了解详情。")

    @pyqtSlot(QMediaPlayer.MediaStatus)
    def handle_media_status(self, status):
        """处理 QMediaPlayer 的状态变化"""
        # (与之前版本相同，使用 PyQt6 枚举)
        if status == QMediaPlayer.MediaStatus.EndOfMedia:
            logging.info("试听音频播放结束。")
            self.lbl_tts_status.setText('状态: 试听播放完毕。')
            self.lbl_tts_status.setStyleSheet("color: green;")
        elif status == QMediaPlayer.MediaStatus.InvalidMedia:
            logging.error("媒体文件无效或无法播放。")
            self.lbl_tts_status.setText('状态: 无法播放试听音频文件。')
            self.lbl_tts_status.setStyleSheet("color: red;")
            self.cleanup_preview_file()
        # ... (其他状态处理保持不变)
        elif status == QMediaPlayer.MediaStatus.LoadedMedia:
            logging.info("媒体加载完成，准备播放。")
            self.lbl_tts_status.setText('状态: 正在播放试听音频...')
            self.lbl_tts_status.setStyleSheet("color: purple;")
        # ... etc ...


    @pyqtSlot(QMediaPlayer.Error, str)
    def handle_player_error(self, error, error_string):
        """处理 QMediaPlayer 的错误信号"""
        # (与之前版本相同)
        logging.error(f"QMediaPlayer 错误: {error} - {error_string}")
        self.lbl_tts_status.setText(f'状态: 播放器错误 - {error_string}')
        self.lbl_tts_status.setStyleSheet("color: red;")
        QMessageBox.critical(self, "播放错误", f"播放音频时遇到错误:\n{error_string}")
        self.cleanup_preview_file()

    def cleanup_preview_file(self):
        """安全地删除当前的临时预览文件"""
        # (与之前版本类似，确保使用 PyQt6 兼容的检查)
        file_to_delete = self.current_preview_file
        if file_to_delete and os.path.exists(file_to_delete):
            logging.info(f"准备清理预览文件: {file_to_delete}")
            player_using_file = False
            if self.player:
                current_source = self.player.source()
                if current_source == QUrl.fromLocalFile(file_to_delete):
                    player_using_file = True

            if player_using_file:
                logging.debug("停止播放器并清除源以释放文件...")
                self.player.stop()
                self.player.setSource(QUrl()) # 清除源
                # 使用 QTimer 延迟删除
                QTimer.singleShot(150, lambda f=file_to_delete: self._delete_file(f))
            else:
                 self._delete_file(file_to_delete)
            self.current_preview_file = None

    def _delete_file(self, filepath):
        """实际执行文件删除操作"""
        # (与之前版本相同)
        try:
            if os.path.exists(filepath):
                os.remove(filepath)
                logging.info(f"已删除临时预览文件: {filepath}")
            else:
                 logging.warning(f"尝试删除文件时文件已不存在: {filepath}")
        except OSError as e:
            logging.warning(f"无法删除文件 '{filepath}': {e} (可能仍被占用)")
        except Exception as e:
            logging.error(f"删除文件 '{filepath}' 时发生意外错误: {e}")


    def showEvent(self, event):
        """重写窗口显示事件，在窗口可见后执行依赖检查。"""
        super().showEvent(event)
        if not self._dependencies_checked:
            self.check_initial_dependencies()
            self._dependencies_checked = True

    def browse_input_file(self):
        # (与之前版本相同)
        filepath, _ = QFileDialog.getOpenFileName(self,"选择 PPTX 文件",self.input_path_edit.text() or str(Path.home()),"PowerPoint 文件 (*.pptx)")
        if filepath: self.input_path_edit.setText(filepath)

    def browse_output_dir(self):
        # (与之前版本相同)
        dirpath = QFileDialog.getExistingDirectory(self,"选择输出目录",self.output_path_edit.text())
        if dirpath: self.output_path_edit.setText(dirpath)

    def check_initial_dependencies(self):
        # (与之前版本相同，只是确认日志输出)
        self.log_output_edit.append("正在检查依赖项...")
        QApplication.processEvents()
        if not check_dependencies():
                # 简化错误信息，详细错误应在日志中
                QMessageBox.critical(self, "依赖错误", "关键依赖项检查失败！\n请查看控制台或日志文件获取详细信息。\n应用程序可能无法正常工作。")
                self.start_button.setEnabled(False)
                self.log_output_edit.append("错误：依赖检查失败！")
        else:
            logging.info("依赖检查通过，应用程序准备就绪。")
            self.log_output_edit.append("依赖检查通过，系统就绪。")


    def start_conversion(self):
        """在后台线程中启动转换过程 (使用 Edge TTS Voice ID)"""
        input_file = self.input_path_edit.text()
        output_dir = self.output_path_edit.text()
        selected_voice_id = self.get_selected_voice_id()

        # --- 输入验证 (增加网络提示和 Voice ID 检查) ---
        if not input_file or not Path(input_file).is_file() or not input_file.lower().endswith(".pptx"):
            QMessageBox.warning(self, "输入无效", "请选择一个有效的 .pptx 文件。")
            return
        output_path_obj = Path(output_dir)
        if not output_dir:
             QMessageBox.warning(self, "输出无效", "请选择一个有效的输出目录。")
             return
        try:
            output_path_obj.mkdir(parents=True, exist_ok=True)
        except Exception as e:
             QMessageBox.warning(self, "输出错误", f"无法创建或访问输出目录。\n错误: {e}")
             return
        if not selected_voice_id:
             QMessageBox.warning(self, "语音未选择", "请在“语音设置”中选择一个旁白语音。")
             return

        # --- 网络确认 ---
        reply = QMessageBox.question(self, "网络确认",
                                     "转换过程中的语音合成需要持续连接网络到微软 Edge TTS 服务。\n请确保您的网络连接稳定。\n\n是否开始转换？",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.Cancel,
                                     QMessageBox.StandardButton.Yes)
        if reply == QMessageBox.StandardButton.Cancel:
            return

        # --- 禁用界面元素 ---
        self.start_button.setEnabled(False)
        self.input_path_edit.setEnabled(False)
        self.input_browse_button.setEnabled(False)
        self.output_path_edit.setEnabled(False)
        self.output_browse_button.setEnabled(False)
        self.cmb_voices.setEnabled(False) # 禁用语音选择
        self.btn_preview.setEnabled(False) # 禁用试听按钮
        self.status_label.setText("系统状态：正在处理")
        self.progress_bar.setValue(0)
        self.log_output_edit.clear()

        config = configparser.ConfigParser()
        config_path = Path(__file__).parent / 'config.ini'
        cleanup = True
        try:
            if config_path.exists(): config.read(config_path, encoding='utf-8')
            cleanup = config.getboolean('通用设置', 'cleanup_temp_dir', fallback=True)
            logging.info(f"根据配置，处理成功后将清理临时文件: {cleanup}")
        except Exception as e:
             logging.warning(f"读取清理配置失败: {e}。将使用默认值: {cleanup}")

        # --- 创建并启动工作线程，传递选定的 Voice ID ---
        self.worker_thread = WorkerThread(
            Path(input_file),
            output_path_obj,
            cleanup,
            selected_voice_id, # <---- 传递 Voice ID
            self.process_presentation_func,
            self.create_video_from_data_func
        )
        self.worker_thread.log_signal.connect(self.update_log)
        self.worker_thread.progress_signal.connect(self.update_progress)
        self.worker_thread.finished_signal.connect(self.on_conversion_finished)
        self.worker_thread.start()

    @staticmethod
    def update_log(message):
        # (与之前版本相同)
        logging.info(f"[后台线程] {message}")

    def update_progress(self, value, status_text):
        # (与之前版本相同)
        self.progress_bar.setValue(value)
        self.status_label.setText(f"系统状态：{status_text.upper()}")
        self.progress_bar.setFormat(f"{status_text} - %p%")

    def on_conversion_finished(self, success, message):
        """处理来自工作线程的 finished_signal 信号。"""
        # --- 重新启用界面元素 ---
        self.start_button.setEnabled(True)
        self.input_path_edit.setEnabled(True)
        self.input_browse_button.setEnabled(True)
        self.output_path_edit.setEnabled(True)
        self.output_browse_button.setEnabled(True)
        self.cmb_voices.setEnabled(True) # 启用语音选择
        self.btn_preview.setEnabled(True) # 启用试听按钮
        self.progress_bar.setFormat("%p%")

        # (后续的成功/失败处理与之前版本相同)
        if success:
            self.status_label.setText("系统状态：处理成功")
            self.progress_bar.setValue(100)
            reply = QMessageBox.information(self,"转换完成",f"处理成功！\n输出文件位于: {message}\n\n是否立即打开输出文件夹？",QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,QMessageBox.StandardButton.Yes)
            if reply == QMessageBox.StandardButton.Yes:
                try:
                    output_folder = str(Path(message).parent.resolve())
                    if sys.platform == "win32": os.startfile(output_folder)
                    elif sys.platform == "darwin": subprocess.run(["open", output_folder])
                    else: subprocess.run(["xdg-open", output_folder])
                except Exception as e:
                    logging.error(f"无法自动打开输出文件夹: {e}")
                    QMessageBox.warning(self, "打开失败", f"无法自动打开输出文件夹。\n请手动前往: {output_folder}")
        else:
            self.status_label.setText("系统状态：处理失败")
            self.progress_bar.setValue(0)
            QMessageBox.critical(self, "转换失败", f"处理过程中发生错误。\n详情请查看日志区域。\n\n错误信息: {message}")
        self.worker_thread = None

    def closeEvent(self, event):
        """处理窗口关闭事件。"""
        # (与之前版本相同，但增加了 TTS 清理)
        if self.worker_thread and self.worker_thread.isRunning():
             reply = QMessageBox.question(self, '退出确认',"转换仍在进行中，确定要强制退出吗？\n（可能导致文件损坏或未清理）",QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,QMessageBox.StandardButton.No)
             if reply == QMessageBox.StandardButton.Yes:
                 logging.warning("用户在处理过程中强制退出。")
                 # 尝试停止线程（效果有限）
                 # self.worker_thread.stop()
                 # self.worker_thread.quit() # 尝试退出事件循环
                 # self.worker_thread.wait(1000) # 等待一小段时间
                 # if self.worker_thread.isRunning():
                 #    self.worker_thread.terminate() # 最后手段
                 self.cleanup_preview_file() # 尝试清理预览文件
                 event.accept()
             else:
                 event.ignore()
        else:
            self.cleanup_preview_file() # 正常退出前清理
            event.accept()


# --- 主应用程序入口点 ---
if __name__ == '__main__':
    app = QApplication(sys.argv)
    qss_file = Path(__file__).parent / "cyberpunk_style.qss"
    if qss_file.exists():
        try:
            with open(qss_file, "r", encoding="utf-8") as f:
                style_sheet = f.read()
                app.setStyleSheet(style_sheet)
            logging.info(f"成功加载样式表: {qss_file.name}")
        except Exception as e:
            logging.error(f"加载样式表 '{qss_file.name}' 失败: {e}")
    else:
        logging.warning(f"样式表文件 '{qss_file.name}' 未找到。将使用默认 Qt 样式。")

    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec())