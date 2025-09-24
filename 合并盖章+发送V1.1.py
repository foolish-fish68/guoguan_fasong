import os
import shutil
import tkinter as tk
from tkinter import filedialog
import fitz  # PyMuPDF
from PIL import Image
import time
import threading
import datetime
import sys


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


def process_files(pdf_files):
    """完整流程：选择文件 → 备份 → 盖章 → 结果输出"""
    # 初始化印章路径
    success_stamp = r"\\192.168.110.248\guoguan\电子印章\过关通过.png"
    fail_stamp = r"\\192.168.110.248\guoguan\电子印章\过关未通过.png"

    # 按目录分组处理文件
    dir_groups = {}
    for path in pdf_files:
        dir_path = os.path.dirname(path)
        if dir_path not in dir_groups:
            dir_groups[dir_path] = []
        dir_groups[dir_path].append(path)

    # 统计结果
    total = len(pdf_files)
    success = 0
    failed = 0

    # 按目录处理文件
    for dir_path, files in dir_groups.items():
        # 创建目录下的备份文件夹
        backup_dir = os.path.join(dir_path, "盖章前留存")
        os.makedirs(backup_dir, exist_ok=True)

        print(f"\n处理目录: {dir_path}")
        print(f"备份文件夹: {backup_dir}")

        # 处理当前目录下的文件
        for path in files:
            # 备份文件
            try:
                shutil.copy2(path, os.path.join(backup_dir, os.path.basename(path)))
                print(f"  已备份: {os.path.basename(path)}")
            except Exception as e:
                print(f"  备份失败: {os.path.basename(path)} - {str(e)}")
                failed += 1
                continue

            # 选择印章
            stamp_path = success_stamp if "过关通过" in path else fail_stamp
            if not os.path.exists(stamp_path):
                print(f"  错误: 印章文件不存在 - {stamp_path}")
                failed += 1
                continue

            # 创建印章图层
            temp_stamp = create_stamp_layer(2480, 3507, stamp_path, 1008, 1100)
            if not temp_stamp:
                failed += 1
                continue

            # 合并印章（非增量模式）
            if apply_stamp(path, temp_stamp):
                success += 1
                print(f"  成功: {os.path.basename(path)}")
            else:
                failed += 1
                print(f"  失败: {os.path.basename(path)}")

    # 输出结果
    print("\n" + "=" * 50)
    print(f"处理结果: 总文件 {total} | 成功 {success} | 失败 {failed}")
    print("=" * 50)

    return success, failed


# 根据模式复制文件到目标目录
def copy_files_by_pattern(source_dir, sent_dir):
    # 定义模式和对应的目标路径，可根据需求调整
    patterns = ["_1", "_2", "_3", "_5","_6", "_8"]
    target_dirs = [
        r"\\ZHISHOU\fasong\25失败",
        r"\\ZHISHOU\fasong\25失败",  # _2 模式对应路径
        r"\\ZHISHOU\fasong\28失败",
        r"\\ZHISHOU\fasong\26失败",
        r"\\ZHISHOU\fasong\26失败",
        r"\\ZHISHOU\fasong\27失败"
    ]
    pattern_map = dict(zip(patterns, target_dirs))
    success, failed, total = 0, 0, 0

    # 确保目标目录存在
    for target_dir in target_dirs:
        os.makedirs(target_dir, exist_ok=True)

    # 确保"已发送"目录存在
    os.makedirs(sent_dir, exist_ok=True)

    for filename in os.listdir(source_dir):
        file_path = os.path.join(source_dir, filename)
        # 仅处理PDF文件
        if not (os.path.isfile(file_path) and filename.lower().endswith('.pdf')):
            continue

        for pattern, target in pattern_map.items():
            if pattern in filename:
                total += 1
                try:
                    # 复制文件到目标目录
                    shutil.copy2(file_path, os.path.join(target, filename))
                    success += 1
                except Exception as e:
                    print(f"复制失败 {filename}: {str(e)}")
                    failed += 1
                break  # 匹配后跳出模式循环

    return success, failed, total


# 移动已处理的PDF到"已发送"文件夹
def move_processed_files(source_dir, sent_dir):
    moved = 0
    for filename in os.listdir(source_dir):
        file_path = os.path.join(source_dir, filename)
        if os.path.isfile(file_path) and filename.lower().endswith('.pdf'):
            try:
                shutil.move(file_path, os.path.join(sent_dir, filename))
                moved += 1
            except Exception as e:
                print(f"移动失败 {filename}: {str(e)}")
    return moved


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
    """后台监测目录变化，interval为检测间隔(秒)"""
    while True:
        try:
            if not os.path.exists(source_dir):
                print(f"错误：监测目录不存在 '{source_dir}'，10秒后重试")
                time.sleep(10)
                continue

            # 检查目录中是否有PDF文件
            has_pdf = any(
                filename.lower().endswith('.pdf')
                for filename in os.listdir(source_dir)
            )

            if has_pdf:
                print(f"\n发现PDF文件，开始处理...")
                pdf_files = [os.path.join(source_dir, f) for f in os.listdir(source_dir) if f.lower().endswith('.pdf')]
                process_files(pdf_files)
                process_directory(source_dir)
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