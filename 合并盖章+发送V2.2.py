import os                           # 导入 Python 的操作系统接口模块，用于执行文件路径操作、系统命令等功能（如文件存在性判断、路径拼接等）
import shutil                       # 导入高级文件操作模块，提供文件复制、移动、删除等更便捷的文件操作功能
import tkinter as tk                # 导入Tkinter GUI 库并为其指定别名tk，用于创建图形用户界面（如窗口、按钮、文件对话框等）
from tkinter import filedialog      # 从 Tkinter 中导入文件对话框模块，专门用于弹出 “打开文件” 或 “保存文件” 的对话框
import fitz  # PyMuPDF              # 导入PyMuPDF 库（别名fitz），用于 PDF 文件的高级处理（如读取、修改、添加内容到 PDF 等）
from PIL import Image               # 从 **Python Imaging Library（PIL）** 中导入Image模块，用于图像处理（如打开、裁剪、缩放图像等）
import time                         # 导入时间模块，用于执行时间相关操作（如延时、获取当前时间戳等）
import threading                    # 导入线程模块，用于实现多线程编程，让程序可以并发执行多个任务
import datetime                     # 导入日期时间模块，用于处理日期、时间的创建、格式化、计算等
import sys                          # 导入系统模块，用于访问 Python 解释器的系统级参数（如命令行参数、退出程序等）
import PyPDF2                       # 用于调整PDF页面大小为A4
from pdfrw import PdfReader, PdfWriter, IndirectPdfDict

def create_stamp_layer(pdf_width_px, pdf_height_px, stamp_path, center_x_px, center_y_px):
    """创建印章图层（A4 300DPI专用）"""
    try:
        img = Image.open(stamp_path)
        img_w, img_h = img.size
        dpi = 300
        px2pt = 72 / dpi  # 像素转点

        # 计算印章在PDF中的坐标和尺寸
        center_x = center_x_px * px2pt
        center_y = center_y_px * px2pt
        img_w_pt = img_w * px2pt
        img_h_pt = img_h * px2pt
        x0 = center_x - img_w_pt / 2
        y0 = center_y - img_h_pt / 2

        # 创建临时印章PDF
        stamp_doc = fitz.open()
        stamp_doc.new_page(width=595.2, height=841.8)  # A4标准尺寸（点）
        stamp_doc[0].insert_image(
            fitz.Rect(x0, y0, x0 + img_w_pt, y0 + img_h_pt),
            filename=stamp_path
        )
        temp_pdf = "temp_stamp.pdf"
        stamp_doc.save(temp_pdf)
        stamp_doc.close()
        return temp_pdf
    except Exception as e:
        print(f"创建印章失败: {str(e)}")
        return None


def apply_stamp(pdf_path, stamp_path):
    """非增量模式合并印章（解决加密文档问题）"""
    try:
        # 创建临时文件
        temp_pdf = f"{pdf_path}.tmp"
        shutil.copy2(pdf_path, temp_pdf)

        # 合并印章
        with fitz.open(temp_pdf) as doc, fitz.open(stamp_path) as stamp:
            doc[0].show_pdf_page(doc[0].rect, stamp, 0)
            doc.save(pdf_path)  # 直接覆盖原文件（非增量）

        # 清理临时文件
        if os.path.exists(temp_pdf):
            os.remove(temp_pdf)
        if os.path.exists(stamp_path):
            os.remove(stamp_path)
        return True
    except Exception as e:
        print(f"合并失败: {str(e)}")
        # 恢复原文件
        if os.path.exists(temp_pdf):
            shutil.copy2(temp_pdf, pdf_path)
            os.remove(temp_pdf)
        return False


def resize_pdf_to_a4(pdf_path):
    """将PDF完整缩放到A4大小（不截取内容，保持全部可见）"""
    try:
        # A4标准尺寸（点，1点=1/72英寸）
        a4_width = 595
        a4_height = 842

        # 读取原PDF
        reader = PyPDF2.PdfReader(pdf_path)
        writer = PyPDF2.PdfWriter()

        for page in reader.pages:
            # 获取原页面实际宽高
            orig_width = float(page.mediabox.upper_right[0] - page.mediabox.lower_left[0])
            orig_height = float(page.mediabox.upper_right[1] - page.mediabox.lower_left[1])

            # 计算缩放比例（取最小比例，确保内容完整放入A4）
            scale = min(a4_width / orig_width, a4_height / orig_height)

            # 计算缩放后的内容尺寸
            scaled_width = orig_width * scale
            scaled_height = orig_height * scale

            # 计算居中偏移量（让内容在A4页面中间显示）
            offset_x = (a4_width - scaled_width) / 2
            offset_y = (a4_height - scaled_height) / 2  # PyPDF2原点在左下角

            # 应用缩放+偏移（通过转换矩阵实现完整缩小+居中）
            page.add_transformation(PyPDF2.Transformation().scale(scale).translate(offset_x, offset_y))

            # 设置页面尺寸为A4
            page.mediabox = PyPDF2.generic.RectangleObject((0, 0, a4_width, a4_height))
            writer.add_page(page)

        # 写入临时文件并替换原文件
        temp_path = f"{pdf_path}.resize.tmp"
        with open(temp_path, "wb") as f:
            writer.write(f)
        os.replace(temp_path, pdf_path)
        print(f"  已完整缩放到A4: {os.path.basename(pdf_path)}")
        return True
    except Exception as e:
        print(f"  缩放A4失败: {str(e)}")
        return False

def process_single_file(file_path):
    """处理单个文件：备份→盖章"""
    # 初始化印章路径
    success_stamp = r"\\192.168.110.248\guoguan\电子印章\过关通过.png"
    fail_stamp = r"\\192.168.110.248\guoguan\电子印章\过关未通过.png"

    dir_path = os.path.dirname(file_path)
    filename = os.path.basename(file_path)
    backup_dir = os.path.join(dir_path, "盖章前留存")
    os.makedirs(backup_dir, exist_ok=True)

    # 备份文件
    try:
        shutil.copy2(file_path, os.path.join(backup_dir, filename))
        print(f"  已备份: {filename}")
    except Exception as e:
        print(f"  备份失败: {filename} - {str(e)}")
        return False

    # 选择印章
    stamp_path = success_stamp if "过关通过" in file_path else fail_stamp
    if not os.path.exists(stamp_path):
        print(f"  错误: 印章文件不存在 - {stamp_path}")
        return False

    # 创建印章图层
    temp_stamp = create_stamp_layer(2481, 3508, stamp_path, 2010, 1035)
    if not temp_stamp:
        return False

    # 合并印章（非增量模式）
    if apply_stamp(file_path, temp_stamp):
        print(f"  成功: {filename}")
        return True
    else:
        print(f"  失败: {filename}")
        return False




def compress_pdf(pdf_path):
    """使用pdfrw压缩PDF（优化结构、清理冗余，保持清晰度）"""
    try:
        temp_path = f"{pdf_path}.compressed.tmp"

        # 读取PDF
        reader = PdfReader(pdf_path)
        writer = PdfWriter()

        # 复制所有页面并优化
        for page in reader.pages:
            # 压缩页面内容流（减小指令体积，不影响图像）
            if page.Contents:
                page.Contents = IndirectPdfDict(
                    stream=page.Contents.stream,
                    filters=['/FlateDecode']  # 启用flate压缩算法
                )
            writer.addpage(page)

        # 写入优化后的PDF（移除冗余对象）
        writer.write(temp_path)

        # 替换原文件
        os.replace(temp_path, pdf_path)
        print(f"  已压缩: {os.path.basename(pdf_path)}")
        return True
    except Exception as e:
        print(f"  PDF压缩失败: {str(e)}")
        if os.path.exists(temp_path):
            os.remove(temp_path)
        return False

def send_single_file(file_path, sent_dir):
    """发送单个文件：复制到目标目录→移动到已发送"""
    filename = os.path.basename(file_path)
    source_dir = os.path.dirname(file_path)

    # 复用学号范围规则
    student_id_ranges = {
        (1000, 2299): r"\\ZHISHOU\fasong\25失败",
        (2300, 2999): r"\\ZHISHOU\fasong\29失败",
        (3000, 3999): r"\\ZHISHOU\fasong\28失败",
        (5000, 6999): r"\\ZHISHOU\fasong\26失败",
        (7500, 7999): r"\\ZHISHOU\fasong\30失败",
        (8000, 8999): r"\\ZHISHOU\fasong\27失败",
        (9700, 9999): r"\\ZHISHOU\fasong\31失败",
    }

    # 匹配学号规则
    import re
    match = re.search(r'^[^_]*?_(\d{4})', filename)
    if not match:
        print(f"  跳过无效文件: {filename}（第一个下划线后未找到4位数字）")
        return False

    try:
        student_id = int(match.group(1))
    except ValueError:
        print(f"  跳过无效文件: {filename}（学号不是有效数字）")
        return False

    target_dir = None
    for (start, end), dir_path in student_id_ranges.items():
        if start <= student_id <= end:
            target_dir = dir_path
            break

    if not target_dir:
        print(f"  跳过未匹配文件: {filename}（学号{student_id}不在任何范围内）")
        return False

    # 复制到目标目录
    try:
        os.makedirs(target_dir, exist_ok=True)
        dest_path = os.path.join(target_dir, filename)
        shutil.copy2(file_path, dest_path)
    except Exception as e:
        print(f"  复制失败 {filename}: {str(e)}")
        return False

    # 移动到已发送文件夹
    try:
        shutil.move(file_path, os.path.join(sent_dir, filename))
        print(f"  已发送: {filename}")
        return True
    except Exception as e:
        print(f"  移动失败 {filename}: {str(e)}")
        return False


def process_directory(source_dir):
    # 获取源目录的父目录，用于创建"已发送"文件夹
    parent_dir = os.path.dirname(source_dir)
    sent_dir = os.path.join(parent_dir, "已发送")

    print(f"开始处理目录: {source_dir}")
    # 执行文件复制
    success, failed, total = copy_files_by_pattern(source_dir, sent_dir)
    print(f"复制完成: 总{total}个PDF文件，成功{success}个，失败{failed}个")

    # 移动已处理的PDF到"已发送"文件夹
    moved = move_processed_files(source_dir, sent_dir)
    print(f"已移动{moved}个文件到'已发送'文件夹")


def monitor_directory(source_dir, interval=600):  # 10分钟=600秒
    """后台监测目录变化，逐个处理文件确保完整流程"""
    while True:
        try:
            if not os.path.exists(source_dir):
                print(f"错误：监测目录不存在 '{source_dir}'，10秒后重试")
                time.sleep(10)
                continue

            # 获取当前目录下所有PDF文件（每次循环重新获取，确保处理新文件）
            pdf_files = [
                os.path.join(source_dir, f)
                for f in os.listdir(source_dir)
                if os.path.isfile(os.path.join(source_dir, f)) and f.lower().endswith('.pdf')
            ]

            if pdf_files:
                print(f"\n发现{len(pdf_files)}个PDF文件，开始逐个处理...")
                # 逐个处理文件
                for file_path in pdf_files:
                    # 1. 处理单个文件（备份+盖章）
                    if not process_single_file(file_path):
                        print(f"  跳过发送未成功处理的文件: {os.path.basename(file_path)}")
                        continue

                    # 新增：调整PDF为A4大小（盖章后、发送前）
                    if not resize_pdf_to_a4(file_path):
                        print(f"  调整A4失败，跳过发送: {os.path.basename(file_path)}")
                        continue

                    # 新增PDF压缩步骤
                    if not compress_pdf(file_path):
                        print(f"  压缩失败，继续发送: {os.path.basename(file_path)}")
                        # 这里可以选择继续发送或中断，根据需求调整

                    # 2. 发送单个文件（复制+移动）
                    sent_dir = os.path.join(os.path.dirname(source_dir), "已发送")
                    os.makedirs(sent_dir, exist_ok=True)
                    send_single_file(file_path, sent_dir)

            else:
                # 显示倒计时
                for i in range(interval, 0, -1):
                    mins, secs = divmod(i, 60)
                    timer = f"{mins:02d}:{secs:02d}"
                    print(f"\r等待下一次检测: {timer}", end="", flush=True)
                    time.sleep(1)
                print("\r", end="", flush=True)  # 清除倒计时
        except Exception as e:
            print(f"\n监测过程中出错: {str(e)}，10秒后重试")
            time.sleep(10)


def main():
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口，只显示文件对话框

    # 弹出对话框让用户选择源目录
    source_dir = filedialog.askdirectory(title="选择需要监测的目录")
    if not source_dir:
        print("未选择源目录，程序退出")
        return

    print(f"已选择监测目录: {source_dir}")
    print("程序开始后台监测，每30分钟检查一次新PDF文件...")

    # 创建后台监测线程
    monitor_thread = threading.Thread(
        target=monitor_directory,
        args=(source_dir,)
    )
    monitor_thread.daemon = True  # 设为守护线程，程序退出时自动结束
    monitor_thread.start()

    # 保持主程序运行
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("\n程序被用户中断")

if __name__ == "__main__":
    main()