# utils.py (或放在 gui_app.py 顶部)
import sys
import os
from pathlib import Path

def resource_path(relative_path):
    """ 获取资源的绝对路径，
        兼容开发环境和 PyInstaller 打包环境。 """
    try:
        # PyInstaller 会创建一个临时文件夹并将路径存储在 sys._MEIPASS
        base_path = Path(sys._MEIPASS)
    except Exception:
        # 如果没有 _MEIPASS，说明在开发环境中
        # 假设此函数所在的 utils.py 与资源文件夹或主脚本在同一层级或可相对访问
        base_path = Path(__file__).parent.parent # 如果 utils.py 在子目录，需要调整
        # 如果直接放在 gui_app.py 里，可以用下面这行:
        # base_path = Path(__file__).parent

    return base_path / relative_path