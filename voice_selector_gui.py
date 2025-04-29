import sys
import os
from pathlib import Path
import logging

# 导入 PyQt6 相关模块
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QComboBox, QPushButton, QMessageBox, QSizePolicy
)
from PyQt6.QtMultimedia import QMediaPlayer, QMediaContent
from PyQt6.QtCore import QUrl, Qt, QTimer

# 导入我们自己的 TTS 管理器
import tts_manager

# --- 配置日志 ---
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - [%(module)s:%(funcName)s] - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

class VoiceSelectorWindow(QWidget):
    """
    一个用于选择和试听 TTS 语音的 PyQt 窗口。
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        self.current_preview_file = None # 存储当前预览文件的路径，用于稍后删除
        self.player = QMediaPlayer(self) # 创建媒体播放器实例
        self.player.mediaStatusChanged.connect(self.handle_media_status) # 连接状态变化信号

        self.initUI()
        self.load_voices()

    def initUI(self):
        """初始化用户界面元素"""
        self.setWindowTitle('语音选择与试听')
        self.setGeometry(300, 300, 500, 200) # 设置窗口位置和大小 (x, y, width, height)

        # --- 创建布局 ---
        main_layout = QVBoxLayout(self)
        form_layout = QHBoxLayout() # 用于标签和下拉框

        # --- 创建控件 ---
        lbl_select = QLabel('选择语音:', self)

        self.cmb_voices = QComboBox(self)
        self.cmb_voices.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed) # 让下拉框水平扩展
        self.cmb_voices.setToolTip("从系统中可用的语音选择一个") # 添加提示信息
        # 当选择项变化时，可以考虑启用/禁用试听按钮（如果需要）
        # self.cmb_voices.currentIndexChanged.connect(self.on_voice_selection_change)

        btn_preview = QPushButton('试听选中语音', self)
        btn_preview.setToolTip("播放选定语音的简短示例")
        btn_preview.clicked.connect(self.preview_selected_voice) # 连接按钮点击事件

        self.lbl_status = QLabel('状态: 请选择语音进行试听', self)
        self.lbl_status.setStyleSheet("color: gray;") # 设置初始状态颜色

        # --- 布局控件 ---
        form_layout.addWidget(lbl_select)
        form_layout.addWidget(self.cmb_voices)

        main_layout.addLayout(form_layout)
        main_layout.addWidget(btn_preview)
        main_layout.addWidget(self.lbl_status)
        # main_layout.addStretch(1) # 添加伸缩项，让控件靠上 (如果需要)

        self.setLayout(main_layout)

    def load_voices(self):
        """加载可用语音到下拉框"""
        self.lbl_status.setText('状态: 正在加载可用语音...')
        self.lbl_status.setStyleSheet("color: blue;")
        QApplication.processEvents() # 处理事件，让状态文本更新显示

        voices = tts_manager.get_available_voices()
        self.cmb_voices.clear() # 清空现有项

        if not voices:
            self.lbl_status.setText('状态: 未找到可用语音。请检查 TTS 引擎。')
            self.lbl_status.setStyleSheet("color: red;")
            QMessageBox.warning(self, "无可用语音", "未能检测到任何 TTS 语音。请确保您的系统已安装并配置了 TTS 引擎（如 Windows SAPI 语音）。")
            self.cmb_voices.setEnabled(False) # 禁用下拉框
            # self.findChild(QPushButton, '试听选中语音').setEnabled(False) # 禁用按钮，如果用名字查找的话
            # 更健壮的方式是保存按钮引用
            preview_button = self.findChild(QPushButton)
            if preview_button:
                preview_button.setEnabled(False)
            return

        logging.info(f"找到 {len(voices)} 个语音，正在填充下拉框...")
        for voice in voices:
            # 显示名称和语言，将 ID 存储为 userData
            display_text = f"{voice.get('name', '未知名称')} ({voice.get('lang', '未知语言')})"
            self.cmb_voices.addItem(display_text, userData=voice.get('id'))
        logging.info("下拉框填充完毕。")

        self.cmb_voices.setEnabled(True)
        preview_button = self.findChild(QPushButton)
        if preview_button:
            preview_button.setEnabled(True)

        self.lbl_status.setText('状态: 语音加载完成，请选择并试听。')
        self.lbl_status.setStyleSheet("color: green;")

    # 可选：如果需要在选择变化时做些事
    # def on_voice_selection_change(self, index):
    #     selected_id = self.cmb_voices.itemData(index)
    #     logging.debug(f"下拉框选择变化: Index={index}, ID='{selected_id}'")
        # 这里可以根据需要更新状态或启用/禁用按钮

    def get_selected_voice_id(self) -> str | None:
        """获取当前下拉框中选定的语音 ID"""
        current_index = self.cmb_voices.currentIndex()
        if current_index < 0: # 没有选中项 (例如列表为空)
            return None
        return self.cmb_voices.itemData(current_index)

    def preview_selected_voice(self):
        """处理“试听”按钮点击事件"""
        selected_voice_id = self.get_selected_voice_id()

        if not selected_voice_id:
            self.lbl_status.setText('状态: 请先从列表中选择一个语音。')
            self.lbl_status.setStyleSheet("color: orange;")
            QMessageBox.information(self, "提示", "请先选择一个语音。")
            return

        self.lbl_status.setText(f'状态: 正在生成 "{self.cmb_voices.currentText()}" 的试听音频...')
        self.lbl_status.setStyleSheet("color: blue;")
        QApplication.processEvents() # 强制 UI 更新

        # --- 在生成新预览前，尝试清理上一个临时文件 ---
        self.cleanup_previous_preview_file()

        # 调用 TTS 管理器生成预览音频
        preview_file_path = tts_manager.generate_preview_audio(selected_voice_id)

        if preview_file_path:
            self.current_preview_file = preview_file_path # 存储新文件的路径
            logging.info(f"试听音频已生成: {preview_file_path}")
            self.lbl_status.setText('状态: 正在播放试听音频...')
            self.lbl_status.setStyleSheet("color: purple;")

            # 使用 QMediaPlayer 播放
            media_url = QUrl.fromLocalFile(preview_file_path)
            media_content = QMediaContent(media_url)
            self.player.setMedia(media_content)
            self.player.play()

        else:
            self.lbl_status.setText('状态: 生成试听音频失败。')
            self.lbl_status.setStyleSheet("color: red;")
            QMessageBox.critical(self, "错误", "生成试听音频失败，请查看日志了解详情。")

    def handle_media_status(self, status):
        """处理 QMediaPlayer 的状态变化"""
        if status == QMediaPlayer.EndOfMedia:
            logging.info("试听音频播放结束。")
            self.lbl_status.setText('状态: 试听播放完毕。请选择其他语音或继续。')
            self.lbl_status.setStyleSheet("color: green;")
            # 可以在这里自动清理文件，但为了避免竞争条件，
            # 我们选择在生成下一个预览或窗口关闭时清理
            # self.cleanup_previous_preview_file()

        elif status == QMediaPlayer.InvalidMedia:
            logging.error("媒体文件无效或无法播放。")
            self.lbl_status.setText('状态: 无法播放试听音频文件。')
            self.lbl_status.setStyleSheet("color: red;")
            self.cleanup_previous_preview_file() # 尝试清理无效文件

        elif status == QMediaPlayer.LoadingMedia:
             logging.debug("正在加载媒体...")
             self.lbl_status.setText('状态: 正在加载试听音频...')
             self.lbl_status.setStyleSheet("color: gray;")

        elif status == QMediaPlayer.LoadedMedia:
             logging.debug("媒体加载完成.")
             # 状态会变为 StalledMedia 或 PlayingMedia

        elif status == QMediaPlayer.StalledMedia:
             logging.warning("媒体播放暂停/缓冲中...")
             self.lbl_status.setText('状态: 试听音频缓冲中...')
             self.lbl_status.setStyleSheet("color: orange;")


    def cleanup_previous_preview_file(self):
        """安全地删除上一个临时预览文件"""
        if self.current_preview_file and os.path.exists(self.current_preview_file):
            try:
                # 在删除前确保播放器已停止使用该文件
                if self.player.state() == QMediaPlayer.PlayingState:
                    self.player.stop() # 尝试停止播放

                # 添加短暂延迟（可能有助于释放文件句柄，但非完全保证）
                # QTimer.singleShot(100, lambda: self._delete_file(self.current_preview_file))
                # 更直接的方式：
                self._delete_file(self.current_preview_file)
                self.current_preview_file = None # 清除引用

            except Exception as e:
                logging.warning(f"自动删除上一个预览文件 '{self.current_preview_file}' 时出错: {e}")

    def _delete_file(self, filepath):
        """实际执行文件删除"""
        try:
            os.remove(filepath)
            logging.info(f"已删除临时预览文件: {filepath}")
        except OSError as e:
            logging.warning(f"无法删除文件 '{filepath}': {e}")
        except Exception as e:
             logging.error(f"删除文件 '{filepath}' 时发生意外错误: {e}")

    def closeEvent(self, event):
        """重写窗口关闭事件，确保清理最后的预览文件"""
        logging.info("窗口关闭事件触发，清理最后的预览文件...")
        self.cleanup_previous_preview_file()
        event.accept() # 接受关闭事件

# --- 主程序入口 ---
if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = VoiceSelectorWindow()
    window.show()
    sys.exit(app.exec_())