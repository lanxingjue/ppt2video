import os
import platform
import time
import shutil
from pathlib import Path
import logging

# --- 配置日志记录 ---
# 设置日志记录，方便查看过程和错误
logging.basicConfig(
    level=logging.INFO,  # 记录 INFO 及以上级别的日志
    format='%(asctime)s - %(levelname)s - [%(funcName)s] - %(message)s', # 日志格式
    datefmt='%Y-%m-%d %H:%M:%S' # 时间格式
)

def export_slides_with_powerpoint(pptx_filepath: Path, output_dir: Path) -> list[str] | None:
    """
    使用 Microsoft PowerPoint 的 COM 接口将 PPTX 文件的所有幻灯片导出为 PNG 图片。
    此函数仅在 Windows 平台上有效。

    Args:
        pptx_filepath: 输入的 PPTX 文件的 Path 对象。
        output_dir: 保存导出 PNG 图片的目标目录的 Path 对象。

    Returns:
        一个包含所有成功导出的图片文件绝对路径的列表 (list[str])。
        如果发生错误或不在 Windows 平台，则返回 None。
    """
    logging.info(f"开始导出任务: 文件 '{pptx_filepath.name}' 到目录 '{output_dir}'")

    # 1. 检查运行平台
    if platform.system() != "Windows":
        logging.error("此导出方法仅支持 Windows 平台。")
        return None

    # 动态导入 pywin32，避免非 Windows 平台出错
    try:
        import win32com.client
        import pythoncom # 需要初始化 COM 环境，尤其是在多线程中
    except ImportError:
        logging.error("缺少 'pywin32' 库。请使用 'pip install pywin32' 安装。")
        return None

    # 2. 检查输入文件是否存在
    if not pptx_filepath.is_file():
        logging.error(f"输入文件不存在: {pptx_filepath}")
        return None

    # 3. 确保输出目录存在，如果不存在则创建
    try:
        output_dir.mkdir(parents=True, exist_ok=True)
        logging.info(f"确保输出目录存在: {output_dir}")
    except OSError as e:
        logging.error(f"创建或访问输出目录失败: {output_dir} - {e}")
        return None

    # --- COM 自动化操作 ---
    powerpoint = None
    presentation = None
    exported_files = []

    try:
        # 初始化 COM 环境 (对于某些情况，尤其是在线程中可能是必要的)
        pythoncom.CoInitialize()

        logging.info("正在启动或连接 PowerPoint 应用程序...")
        # 尝试连接到已运行的 PowerPoint 实例，如果失败则启动新的实例
        try:
            powerpoint = win32com.client.GetActiveObject("PowerPoint.Application")
            logging.info("已连接到正在运行的 PowerPoint 实例。")
        except pythoncom.com_error:
            powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            logging.info("已启动新的 PowerPoint 实例。")

        # 设置 PowerPoint 是否可见 (0=不可见, 1=可见)，后台运行设为 0
        powerpoint.Visible = 1 # 调试时可以设为 1 查看过程
        # powerpoint.Visible = 0 # 实际运行时设为 0

        # 设置自动化时是否显示警告框 (0=不显示, 1=显示)
        powerpoint.DisplayAlerts = 0 # ppAlertsNone

        logging.info(f"正在打开演示文稿: {pptx_filepath.resolve()}")
        # 使用绝对路径打开文件，WithWindow=False 表示不在单独窗口打开（即使主程序可见）
        presentation = powerpoint.Presentations.Open(str(pptx_filepath.resolve()), WithWindow=False)

        logging.info(f"开始导出幻灯片 (共 {len(presentation.Slides)} 张)...")
        start_time = time.time()

        # 遍历所有幻灯片并导出
        for i, slide in enumerate(presentation.Slides):
            slide_number = i + 1 # 幻灯片编号通常从 1 开始
            output_filename = f"slide_{slide_number}.png"
            output_filepath = output_dir / output_filename
            abs_output_filepath = str(output_filepath.resolve()) # 获取绝对路径

            logging.info(f"  正在导出幻灯片 {slide_number} 到 {abs_output_filepath}...")
            try:
                # 执行导出操作
                # 参数: 文件路径, 导出格式 (PNG, JPG, GIF, TIF, BMP), [宽度], [高度]
                slide.Export(abs_output_filepath, "PNG")
                if output_filepath.exists():
                    exported_files.append(abs_output_filepath)
                    logging.info(f"    幻灯片 {slide_number} 导出成功。")
                else:
                    # 理论上 Export 不报错就应该成功，但加个检查更保险
                    logging.warning(f"    幻灯片 {slide_number} 导出后文件未找到: {abs_output_filepath}")

            except pythoncom.com_error as e:
                logging.error(f"    导出幻灯片 {slide_number} 时发生 COM 错误: {e}")
                # 可以选择在这里中断或继续导出其他幻灯片
                # return None # 如果希望一出错就停止，取消此行注释
            except Exception as e:
                 logging.error(f"    导出幻灯片 {slide_number} 时发生未知错误: {e}")
                 # return None

        end_time = time.time()
        logging.info(f"幻灯片导出完成，耗时 {end_time - start_time:.2f} 秒。共成功导出 {len(exported_files)} 张图片。")

        return exported_files

    except pythoncom.com_error as e:
        logging.error(f"处理 PowerPoint 文件时发生 COM 错误: {e}")
        return None
    except Exception as e:
        logging.error(f"处理 PowerPoint 文件时发生未知错误: {e}")
        return None
    finally:
        # --- 清理 COM 对象 ---
        # 确保无论成功还是失败都尝试关闭演示文稿和退出 PowerPoint
        if presentation:
            try:
                logging.info("正在关闭演示文稿...")
                presentation.Close()
            except Exception as e:
                logging.warning(f"关闭演示文稿时出错: {e}")
        if powerpoint:
            try:
                logging.info("正在退出 PowerPoint 应用程序...")
                powerpoint.Quit()
            except Exception as e:
                logging.warning(f"退出 PowerPoint 时出错: {e}")

        # 释放 COM 对象引用 (可选，有助于垃圾回收)
        # del presentation
        # del powerpoint

        # 反初始化 COM 环境
        pythoncom.CoUninitialize()
        logging.info("COM 环境已清理。")


# --- 示例用法 ---
if __name__ == "__main__":
    logging.info("--- 开始测试 PPT 导出功能 (仅 Windows) ---")

    # --- 请修改以下路径为您系统中的实际路径 ---
    # 输入 PPTX 文件路径
    input_ppt_file = Path("智能短信分类平台方案.pptx") # 例如: Path("C:/Users/YourUser/Documents/presentation.pptx")
    # 输出图片的目标目录
    output_image_dir = Path("./exported_slides_output")
    # --- 修改结束 ---


    if not input_ppt_file.exists():
         print(f"错误：请将第 148 行的 'your_presentation.pptx' 替换为您要测试的实际 PPTX 文件路径。")
    else:
        # 如果输出目录已存在，先清空（可选）
        if output_image_dir.exists():
            logging.warning(f"输出目录 '{output_image_dir}' 已存在，将清空其中内容。")
            try:
                # 确保是目录再删除
                if output_image_dir.is_dir():
                    shutil.rmtree(output_image_dir)
                else:
                    output_image_dir.unlink() # 如果是文件，删除文件
                output_image_dir.mkdir(parents=True, exist_ok=True) # 重新创建目录
            except Exception as e:
                 logging.error(f"清空输出目录 '{output_image_dir}' 时出错: {e}")
                 # 如果无法清空，可能后续导出也会失败，可以选择退出

        # 调用导出函数
        exported_image_paths = export_slides_with_powerpoint(input_ppt_file, output_image_dir)

        if exported_image_paths:
            print("\n--- 导出成功 ---")
            print(f"图片已保存到目录: {output_image_dir.resolve()}")
            print("导出的文件列表:")
            for img_path in exported_image_paths:
                print(f"- {img_path}")
        else:
            print("\n--- 导出失败 ---")
            print("请检查上面的日志输出获取详细错误信息。")
            print("可能的原因包括：")
            print("- 未安装 Microsoft PowerPoint。")
            print("- 未安装 'pywin32' 库。")
            print("- 输入的 PPTX 文件路径错误或文件已损坏。")
            print("- 输出目录权限不足。")
            print("- PowerPoint 程序在自动化过程中崩溃或无响应。")

    logging.info("--- 测试结束 ---")