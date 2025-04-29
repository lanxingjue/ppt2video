# ppt_exporter_libreoffice.py
import os
import platform
import subprocess
import logging
import shutil
from pathlib import Path
import tempfile
import configparser
import shlex
# 导入 pdf2image
try:
    from pdf2image import convert_from_path, pdfinfo_from_path
    PDF2IMAGE_AVAILABLE = True
except ImportError:
    logging.error("缺少 'pdf2image' 库。请运行 'pip install pdf2image'。")
    PDF2IMAGE_AVAILABLE = False

# --- 配置解析 (用于获取 LibreOffice 和 Poppler 路径) ---
config = configparser.ConfigParser()
config_path = Path(__file__).parent / 'config.ini'
if config_path.exists():
    try: config.read(config_path, encoding='utf-8')
    except Exception as e: logging.error(f"加载配置失败 (LibreOffice Exporter): {e}")

# --- 获取 LibreOffice 和 Poppler 路径 ---
# (可以使用和 main_controller.py 中类似的 get_ffmpeg_tool_path 函数思路)
def get_tool_path(config_section, config_key, default_name):
    tool_path_config = config.get(config_section, config_key, fallback=default_name)
    resolved_path = shutil.which(tool_path_config)
    if resolved_path:
        return str(Path(resolved_path).resolve())
    logging.warning(f"未能找到 {default_name}。请确保已安装并添加到 PATH 或在 config.ini 中配置。")
    # 检查是否在标准 macOS 应用目录 (如果未找到)
    if platform.system() == "Darwin":
        common_path = f"/Applications/LibreOffice.app/Contents/MacOS/{default_name}"
        if Path(common_path).exists():
            logging.info(f"在默认 macOS 路径找到 {default_name}: {common_path}")
            return common_path
    return None # 返回 None 表示找不到

LIBREOFFICE_PATH = get_tool_path('Paths', 'libreoffice_path', 'soffice')
# Poppler 路径 (pdf2image 需要) - pdf2image 可以接受 poppler_path 参数
POPPLER_PATH_CONFIG = config.get('Paths', 'poppler_path', fallback=None)
if POPPLER_PATH_CONFIG and not Path(POPPLER_PATH_CONFIG).is_dir():
    logging.warning(f"配置的 Poppler 路径 '{POPPLER_PATH_CONFIG}' 不是有效目录，将依赖系统 PATH。")
    POPPLER_PATH_CONFIG = None # 如果配置路径无效，则依赖 PATH

def export_slides_with_libreoffice(pptx_filepath: Path, output_dir: Path) -> list[str] | None:
    """
    使用 LibreOffice 将 PPTX 转换为 PDF，然后使用 pdf2image 将 PDF 转换为 PNG 图片。

    Args:
        pptx_filepath: 输入的 PPTX 文件的 Path 对象。
        output_dir: 保存导出 PNG 图片的目标目录的 Path 对象。

    Returns:
        一个包含所有成功导出的图片文件绝对路径的列表 (list[str])。
        如果发生错误，则返回 None。
    """
    logging.info(f"尝试使用 LibreOffice 导出: '{pptx_filepath.name}' 到 '{output_dir}'")

    if not PDF2IMAGE_AVAILABLE:
        logging.error("pdf2image 库不可用，无法进行 PDF 到图片的转换。")
        return None
    if LIBREOFFICE_PATH is None:
        logging.error("LibreOffice (soffice) 未找到。请检查安装和配置。")
        return None

    # 1. 检查输入文件
    if not pptx_filepath.is_file():
        logging.error(f"输入文件不存在: {pptx_filepath}")
        return None

    # 2. 确保输出目录存在
    try:
        output_dir.mkdir(parents=True, exist_ok=True)
        logging.info(f"确保输出目录存在: {output_dir}")
    except OSError as e:
        logging.error(f"创建或访问输出目录失败: {output_dir} - {e}")
        return None

    # 3. 创建临时目录存放中间 PDF 文件
    with tempfile.TemporaryDirectory(prefix="lo_pdf_") as temp_pdf_dir_str:
        temp_pdf_dir = Path(temp_pdf_dir_str)
        pdf_output_path = temp_pdf_dir / f"{pptx_filepath.stem}.pdf"
        logging.info(f"创建临时 PDF 目录: {temp_pdf_dir}")

        # 4. 调用 LibreOffice 将 PPTX 转换为 PDF
        cmd_convert_to_pdf = [
            LIBREOFFICE_PATH,
            "--headless",          # 无头模式运行
            "--convert-to", "pdf", # 指定转换目标格式为 PDF
            "--outdir", str(temp_pdf_dir.resolve()), # 指定 PDF 输出目录
            str(pptx_filepath.resolve()) # 输入的 PPTX 文件
        ]
        logging.info(f"执行 LibreOffice 命令将 PPTX 转为 PDF: {' '.join(shlex.quote(c) for c in cmd_convert_to_pdf)}")
        try:
            result_pdf = subprocess.run(cmd_convert_to_pdf, capture_output=True, text=True, timeout=120, check=True, encoding='utf-8', errors='ignore') # 增加超时
            logging.info("LibreOffice 转换 PDF 命令执行完成。")
            if result_pdf.stdout: logging.debug(f"LibreOffice STDOUT:\n{result_pdf.stdout}")
            if result_pdf.stderr: logging.debug(f"LibreOffice STDERR:\n{result_pdf.stderr}")

            if not pdf_output_path.exists():
                logging.error(f"LibreOffice 命令执行后未找到预期的 PDF 文件: {pdf_output_path}")
                # 尝试列出临时目录内容帮助调试
                try:
                    files_in_temp = list(temp_pdf_dir.glob('*'))
                    logging.error(f"临时 PDF 目录内容: {files_in_temp}")
                except Exception as list_e:
                    logging.error(f"无法列出临时 PDF 目录内容: {list_e}")
                return None

        except subprocess.CalledProcessError as e:
            logging.error(f"LibreOffice 转换 PDF 失败。返回码: {e.returncode}")
            logging.error(f"命令: {' '.join(shlex.quote(c) for c in cmd_convert_to_pdf)}")
            logging.error(f"STDERR:\n{e.stderr}")
            logging.error(f"STDOUT:\n{e.stdout}")
            return None
        except subprocess.TimeoutExpired:
            logging.error("LibreOffice 转换 PDF 超时 (120 秒)。")
            return None
        except FileNotFoundError:
            logging.error(f"错误：找不到 LibreOffice 命令 '{LIBREOFFICE_PATH}'。")
            return None
        except Exception as e:
            logging.error(f"执行 LibreOffice 转换时发生未知错误: {e}", exc_info=True)
            return None

        # 5. 调用 pdf2image 将 PDF 转换为图片
        logging.info("开始使用 pdf2image 将 PDF 转换为 PNG 图片...")
        exported_files = []
        try:
            # 使用 poppler_path 参数 (如果配置了)
            images = convert_from_path(
                pdf_output_path,
                output_folder=output_dir,
                fmt='png',                # 输出格式为 PNG
                output_file="slide_",     # 文件名前缀
                paths_only=True,          # 只返回路径列表
                use_pdftocairo=True,     # 尝试使用 pdftocairo (通常质量更好)
                poppler_path=POPPLER_PATH_CONFIG # 传递 Poppler 路径
            )
            # pdf2image 返回的路径可能需要排序和处理
            # 它生成的路径类似 output_dir/slide_-000001.png
            # 我们需要重命名为 slide_1.png, slide_2.png ...
            image_paths_generated = sorted(output_dir.glob("slide_*.png"))
            num_pages = len(image_paths_generated)
            logging.info(f"pdf2image 成功转换了 {num_pages} 页。")

            for i, old_path in enumerate(image_paths_generated):
                slide_number = i + 1
                new_filename = f"slide_{slide_number}.png"
                new_path = output_dir / new_filename
                try:
                    # 如果目标文件已存在（不太可能，但保险起见），先删除
                    if new_path.exists() and old_path != new_path:
                        new_path.unlink()
                    old_path.rename(new_path)
                    exported_files.append(str(new_path.resolve()))
                    logging.debug(f"  重命名图片: {old_path.name} -> {new_path.name}")
                except OSError as rename_e:
                    logging.error(f"  重命名图片 {old_path.name} 失败: {rename_e}")
                    # 如果重命名失败，尝试将原始路径加入列表，但可能后续处理出问题
                    exported_files.append(str(old_path.resolve()))


            if len(exported_files) != num_pages:
                 logging.warning("重命名后的图片数量与转换的页面数量不符。")

            logging.info(f"成功导出并整理了 {len(exported_files)} 张图片。")
            return exported_files

        except Exception as e:
            # 区分 pdfinfo 错误和 convert 错误
            if "Unable to get page count" in str(e) or "pdfinfo" in str(e).lower():
                 logging.error(f"pdf2image 错误: 无法获取 PDF 信息。请确保 Poppler 工具已安装并可在 PATH 或配置路径中找到。Poppler Path Config: {POPPLER_PATH_CONFIG}")
            elif "pdftocairo" in str(e).lower():
                 logging.error(f"pdf2image 错误: pdftocairo 执行失败。请确保 Poppler 工具安装完整。")
            else:
                 logging.error(f"pdf2image 转换 PDF 到图片时出错: {e}", exc_info=True)
            return None

    # 临时 PDF 目录会在 with 语句结束时自动清理