#!/usr/bin/env python
# -*- coding: utf-8 -*-
r"""
PDZ -> PDF 转换工具 (通用版)

用法:
    python pdz2pdf.py <输入.pdz> [输出.pdf]

示例:
    python pdz2pdf.py book.pdz
    python pdz2pdf.py book.pdz output.pdf

原理:
    利用超星阅读器(ssReader)的页面缓存机制，通过 PostMessage 模拟翻页，
    从 %%LOCALAPPDATA%%\Temp\buffer 中逐个提取页面图像，最终合成为 PDF。

要求:
    - 超星阅读器(ssReader) 5.x 已安装
    - 脚本会自动请求管理员权限（PostMessage 需要绕过 UIPI）
    - Python 3.x + Pillow + PyMuPDF + pywin32
"""

import ctypes
import time
import os
import sys
import subprocess
import argparse
import glob
import shutil
import re

# ============================================================================
# 管理员权限处理
# ============================================================================
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    """以管理员权限重新启动当前脚本"""
    if not is_admin():
        print("=" * 50)
        print("此脚本需要管理员权限才能正常工作。")
        print('正在请求权限提升，请在弹出的 UAC 对话框中点击"是"...')
        print("(如果没看到弹窗，请检查任务栏是否有闪烁图标)")
        print("=" * 50)

        # 构建完整命令行（全部使用绝对路径）
        script = os.path.abspath(__file__)
        args = [script] + sys.argv[1:]
        cmd = f'"{sys.executable}" ' + ' '.join(f'"{a}"' for a in args)

        # 构建参数字符串
        params = " ".join(f'"{a}"' for a in args)

        # 方式1: ShellExecute (直接)
        ret = ctypes.windll.shell32.ShellExecuteW(
            None, "runas", sys.executable,
            params,
            os.path.dirname(script), 1
        )
        if ret > 32:  # ShellExecute 成功（返回值 > 32 表示成功）
            sys.exit(0)

        # 方式2: PowerShell (备用)
        print("ShellExecute 失败，尝试通过 PowerShell 提升权限...")
        ps_script = (
            f'Start-Process -FilePath "{sys.executable}" '
            f'-ArgumentList "{params}" -Verb RunAs'
        )
        subprocess.run(["powershell", "-Command", ps_script])
        sys.exit(0)

# ============================================================================
# 依赖检查
# ============================================================================
def check_dependencies():
    missing = []
    try:
        import win32gui
        import win32con
    except ImportError:
        missing.append("pywin32 (pip install pywin32)")
    try:
        from PIL import Image
    except ImportError:
        missing.append("Pillow (pip install Pillow)")
    try:
        import fitz
    except ImportError:
        missing.append("PyMuPDF (pip install PyMuPDF)")

    if missing:
        print("缺少依赖库:\n  " + "\n  ".join(missing))
        print("\n请安装后再运行。")
        sys.exit(1)

# ============================================================================
# ssReader 路径检测
# ============================================================================
def find_ssreader():
    """自动查找超星阅读器安装路径"""
    candidates = []

    # 1. 检查脚本同目录或上级目录
    script_dir = os.path.dirname(os.path.abspath(__file__))
    for base in [script_dir, os.path.dirname(script_dir)]:
        for pattern in ["超星阅读器*", "ssreader*", "SSReader*"]:
            for d in glob.glob(os.path.join(base, pattern)):
                exe = os.path.join(d, "ssReader.exe")
                if os.path.exists(exe):
                    candidates.append(exe)

    # 2. 检查常见安装位置
    common_paths = [
        r"C:\Program Files (x86)\SSReader\ssReader.exe",
        r"C:\Program Files\SSReader\ssReader.exe",
        r"D:\ssreader\超星阅读器5\ssReader.exe",
    ]
    for p in common_paths:
        if os.path.exists(p):
            candidates.append(p)

    # 3. 从注册表查找
    try:
        import winreg
        for root in [winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER]:
            for subkey in [r"Software\SSReader", r"Software\Chaoxing"]:
                try:
                    key = winreg.OpenKey(root, subkey)
                    val, _ = winreg.QueryValueEx(key, "InstallPath")
                    exe = os.path.join(val, "ssReader.exe")
                    if os.path.exists(exe):
                        candidates.append(exe)
                    winreg.CloseKey(key)
                except:
                    pass
    except:
        pass

    # 4. 从 .pdz 文件关联查找
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, ".pdz") as key:
            prog_id = winreg.QueryValue(key, "")
        if prog_id:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT,
                               f"{prog_id}\\shell\\open\\command") as key:
                cmd = winreg.QueryValue(key, "")
                import shlex
                exe = shlex.split(cmd)[0]
                if os.path.exists(exe):
                    candidates.append(exe)
    except:
        pass

    return candidates[0] if candidates else None

# ============================================================================
# 核心转换逻辑
# ============================================================================
def convert_pdz(pdz_path, output_pdf, ssreader_path=None, max_pages=None):
    """
    将 PDZ 文件转换为 PDF

    Args:
        pdz_path:     输入的 .pdz 文件路径
        output_pdf:   输出的 .pdf 文件路径
        ssreader_path: ssReader.exe 的路径 (可选，自动检测)
        max_pages:    手动指定最多转换页数 (可选，默认自动检测)
    """
    import win32gui
    import win32con
    from PIL import Image

    # ---- 验证输入 ----
    if not os.path.exists(pdz_path):
        print(f"错误: 文件不存在 - {pdz_path}")
        sys.exit(1)

    pdz_path = os.path.abspath(pdz_path)
    output_pdf = os.path.abspath(output_pdf)
    work_dir = os.path.dirname(pdz_path)

    # ---- 找到 ssReader ----
    if ssreader_path is None:
        ssreader_path = find_ssreader()
    if ssreader_path is None or not os.path.exists(ssreader_path):
        print("错误: 未找到超星阅读器 (ssReader.exe)")
        print("请用 --ssreader 参数指定路径")
        sys.exit(1)

    print(f"ssReader: {ssreader_path}")
    print(f"输入文件: {pdz_path}")
    print(f"输出文件: {output_pdf}")

    # ---- 路径配置 ----
    buffer_dir = os.path.join(os.environ.get("LOCALAPPDATA", ""), "Temp", "buffer")
    pages_dir = os.path.join(os.path.dirname(output_pdf),
                             f".pdz_pages_{os.path.basename(pdz_path)}")

    print(f"缓存目录: {buffer_dir}")
    print()

    # =====================================================================
    # 阶段 1: 清理环境并启动 ssReader
    # =====================================================================
    print("[1/4] 启动超星阅读器...")

    # 关闭已有的 ssReader
    subprocess.run(["taskkill", "/F", "/IM", "ssReader.exe"],
                   capture_output=True)
    time.sleep(2)

    # 清理旧的缓存
    os.makedirs(buffer_dir, exist_ok=True)
    for f in os.listdir(buffer_dir):
        try:
            p = os.path.join(buffer_dir, f)
            if os.path.isfile(p):
                os.remove(p)
        except:
            pass

    # 清理旧的页面目录
    if os.path.exists(pages_dir):
        shutil.rmtree(pages_dir)
    os.makedirs(pages_dir, exist_ok=True)

    # 启动 ssReader 并打开 PDZ 文件
    subprocess.Popen([ssreader_path, pdz_path], cwd=work_dir)
    time.sleep(8)  # 等待程序加载

    # 查找 ssReader 窗口
    hwnd = None
    for _ in range(15):
        hwnd = win32gui.FindWindow("ssReader", None)
        if hwnd:
            break
        time.sleep(1)

    if not hwnd:
        print("错误: 未能找到 ssReader 窗口")
        print("请确认超星阅读器已正确安装并能打开 PDZ 文件")
        sys.exit(1)

    print(f"  窗口句柄: {hwnd}")

    # =====================================================================
    # 阶段 2: 确定总页数
    # =====================================================================
    def post_key(vk):
        """向 ssReader 窗口发送按键消息"""
        win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, vk, 0)
        time.sleep(0.05)
        win32gui.PostMessage(hwnd, win32con.WM_KEYUP, vk, 0)

    if max_pages is not None:
        total_pages = max_pages
        print(f"[2/4] 使用手动指定页数: {total_pages}")
        # 跳到末页触发完整渲染，然后回首页
        post_key(0x23)  # VK_END
        time.sleep(5)
        post_key(0x24)  # VK_HOME
        time.sleep(3)
    else:
        print("[2/4] 自动检测文档页数...")
        time.sleep(5)

        # 从 buffer 文件名推断总页数
        total_pages = 0
        for f in os.listdir(buffer_dir):
            if f.endswith(".bmp") and "_page_" in f:
                try:
                    parts = f.replace(".bmp", "").split("_")
                    total_pages = max(total_pages, int(parts[-1]))
                except:
                    pass

        if total_pages == 0:
            print("警告: 无法确定页数，使用默认值 300")
            total_pages = 300
        else:
            print(f"  检测到约 {total_pages} 页")

        # 回到第一页
        post_key(0x24)  # VK_HOME
        time.sleep(3)

    # =====================================================================
    # 阶段 3: 逐页收集图像
    # =====================================================================
    print(f"[3/4] 逐页提取 (共 {total_pages} 页)...")

    collected = set()
    page_count = 0
    no_new_count = 0
    consecutive_stall = 0

    while page_count < total_pages:
        # 扫描新文件
        new_snapshot = set()
        for f in os.listdir(buffer_dir):
            if f.endswith(".bmp"):
                new_snapshot.add(os.path.join(buffer_dir, f))

        new_files = new_snapshot - collected

        if new_files:
            consecutive_stall = 0
            no_new_count = 0
            for fp in sorted(new_files, key=os.path.getmtime):
                page_count += 1
                collected.add(fp)

                # 复制到输出目录
                dst = os.path.join(pages_dir, f"page_{page_count:04d}.bmp")
                try:
                    shutil.copy2(fp, dst)
                    size_kb = os.path.getsize(dst) // 1024
                    print(f"  [{page_count:4d}] page_{page_count:04d}.bmp"
                          f" ({size_kb} KB)")
                except Exception as e:
                    print(f"  [{page_count:4d}] 复制失败: {e}")
        else:
            no_new_count += 1
            if no_new_count > 5:
                consecutive_stall += 1
                if consecutive_stall >= 3:
                    print(f"\n  已连续 {consecutive_stall} 轮无新页面，到达末页")
                    break

        # 翻页
        post_key(0x22)  # VK_NEXT (Page Down)
        time.sleep(1.5)

    print(f"\n  共提取 {page_count} 页")

    # =====================================================================
    # 阶段 4: 合成 PDF
    # =====================================================================
    print("[4/4] 合成 PDF...")

    page_files = sorted(
        [f for f in os.listdir(pages_dir) if f.endswith(".bmp")],
        key=lambda x: int(re.search(r"(\d+)", x).group())
    )

    if not page_files:
        print("错误: 未提取到任何页面")
        sys.exit(1)

    # 使用 Pillow 合成 PDF
    images = []
    for pf in page_files:
        fp = os.path.join(pages_dir, pf)
        try:
            img = Image.open(fp)
            if img.mode != "RGB":
                img = img.convert("RGB")
            images.append(img)
        except Exception as e:
            print(f"  跳过损坏文件 {pf}: {e}")

    if images:
        images[0].save(
            output_pdf,
            "PDF",
            save_all=True,
            append_images=images[1:]
        )
        pdf_size = os.path.getsize(output_pdf) / 1024 / 1024
        print(f"\n  输出文件: {output_pdf}")
        print(f"  文件大小: {pdf_size:.1f} MB")
        print(f"  页    数: {len(images)}")
    else:
        print("错误: 没有有效页面可供合成")
        sys.exit(1)

    # =====================================================================
    # 清理
    # =====================================================================
    # 关闭 ssReader
    subprocess.run(["taskkill", "/F", "/IM", "ssReader.exe"],
                   capture_output=True)

    # 删除临时页面文件
    try:
        shutil.rmtree(pages_dir)
    except:
        pass

    print(f"\n完成! PDF 已保存到: {output_pdf}")
    return output_pdf


# ============================================================================
# 命令行入口
# ============================================================================
def main():
    check_dependencies()

    parser = argparse.ArgumentParser(
        description="PDZ → PDF 转换工具 (超星电子书格式转换)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  python pdz2pdf.py book.pdz
  python pdz2pdf.py book.pdz D:/output/book.pdf
  python pdz2pdf.py book.pdz --ssreader "D:/tools/ssReader.exe"

注意:
  本脚本需要管理员权限才能正常工作。
  首次运行时会弹出 UAC 提示，请选择"是"。
        """
    )
    parser.add_argument(
        "input",
        help="输入的 .pdz 文件路径"
    )
    parser.add_argument(
        "output",
        nargs="?",
        default=None,
        help="输出的 .pdf 文件路径 (默认: 与输入同名，扩展名改为 .pdf)"
    )
    parser.add_argument(
        "--ssreader",
        default=None,
        help="ssReader.exe 的路径 (默认: 自动检测)"
    )
    parser.add_argument(
        "--pages",
        type=int,
        default=None,
        help="手动指定转换页数 (默认: 自动检测，如 --pages 100 只转换前100页)"
    )

    args = parser.parse_args()

    # 生成默认输出路径
    if args.output is None:
        base = os.path.splitext(args.input)[0]
        args.output = base + ".pdf"

    # 确保管理员权限
    run_as_admin()

    # 执行转换
    convert_pdz(args.input, args.output, args.ssreader, args.pages)


if __name__ == "__main__":
    main()
