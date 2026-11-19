# -*- coding: utf-8 -*-
"""
生成系统演示视频（无声 MP4）
================================================================
原理：用无头 Chrome 对**真实运行中的系统**按分镜逐帧截图（URL 参数 ?demo=N
驱动前端演示模式），再用 PIL 叠加字幕条、OpenCV 合成带交叉溶解的无声 MP4。
画面全部来自真实浏览器渲染的真实接口数据，不是模拟界面、也不是 AI 生成画面。

前置条件：
  1) 服务已启动：python src/api/server.py   （http://localhost:8000 可访问）
  2) 本机有 Chrome 或 Edge（Chromium 内核）
  3) Python 依赖：opencv-python、Pillow（运行系统本身不需要这两个）

运行：
  python scripts/make_demo_video.py
  python scripts/make_demo_video.py --keep-frames     # 保留中间截图供检查

输出：D:\\rag\\系统演示视频.mp4（默认 1280x962、20fps、无声、14 个分镜、约 3 分钟）
     D:\\rag\\系统演示视频_轻量版.mp4（960x720、15fps、约 9~12 MB，便于上传提交）
"""
import os
import sys
import time
import glob
import shutil
import argparse
import subprocess
import urllib.request

import cv2
import numpy as np
from PIL import Image, ImageDraw, ImageFont

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
FRAME_DIR = os.path.join(ROOT, "_demo_frames")
OUT_MP4 = os.path.join(ROOT, "系统演示视频.mp4")
BASE_URL = "http://localhost:8000"

CHROME_CANDIDATES = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
]
FONT_CANDIDATES = [
    r"C:\Windows\Fonts\msyh.ttc",
    r"C:\Windows\Fonts\simhei.ttf",
    r"C:\Windows\Fonts\simsun.ttc",
]

# 分镜：(演示步骤编号, 字幕, 停留秒数)
STEPS = [
    (1,  "系统概览 · 742 个节点 / 13002 条关系",                       12.0),
    (2,  "图谱可视化 · 加载示例知识图谱（力导向布局）",                 12.0),
    (3,  "图谱可视化 · 搜索实体「牛肉」并查看节点详情",                 12.0),
    (4,  "图谱增强 RAG 问答 · 「补气血的汤有哪些？」",                   12.0),
    (5,  "图谱增强 RAG 问答 · 「牛奶过敏能用什么替代？」（画像：奶过敏）", 12.0),
    (6,  "图谱增强 RAG 问答 · 「孕妇能吃螃蟹吗？」",                     12.0),
    (11, "图谱增强 RAG 问答 · 「番茄和鸡蛋能一起吃吗？」",               12.0),
    (12, "图谱增强 RAG 问答 · 「适合老人的补钙菜」",                     12.0),
    (7,  "食材替代推理 · 五花肉 → 瘦肉 / 猪肉 等候选",                   12.0),
    (13, "食材替代推理 · 牛奶 + 奶过敏画像，自动过滤含奶候选",           12.0),
    (8,  "饮食禁忌校验 · 螃蟹、年糕、姜（画像：孕妇）",                   12.0),
    (14, "饮食禁忌校验 · 鸡蛋 + 豆浆（相克关系）",                       12.0),
    (9,  "系统后台管理 · 菜谱数据列表",                                  12.0),
    (10, "系统后台管理 · 新增菜谱表单",                                  12.0),
]

FPS = 20
CROSS_FRAMES = 10          # 镜头间交叉溶解帧数（0.5 秒）
BAR_H = 62                 # 字幕条高度
COLOR_BG = (44, 82, 130)   # 与前端 header 一致的深蓝
COLOR_WHITE = (255, 255, 255)


def log(msg):
    print(msg, flush=True)


def find_chrome():
    for p in CHROME_CANDIDATES:
        if os.path.exists(p):
            return p
    raise RuntimeError("未找到 Chrome / Edge，无法生成演示视频帧")


def load_font(size):
    for p in FONT_CANDIDATES:
        if os.path.exists(p):
            try:
                return ImageFont.truetype(p, size)
            except Exception:
                pass
    return ImageFont.load_default()


def check_service():
    try:
        with urllib.request.urlopen(BASE_URL + "/", timeout=8) as r:
            return r.status == 200
    except Exception:
        return False


def capture(chrome, step, path, budget=9000):
    """用无头浏览器截取某一演示步骤的真实界面；等待截图文件落盘。"""
    if os.path.exists(path):
        os.remove(path)
    profile = os.path.join(ROOT, "_chrome_step%d" % step)
    args = [
        chrome, "--headless=new", "--disable-gpu", "--no-sandbox",
        "--no-first-run", "--disable-extensions", "--hide-scrollbars",
        "--user-data-dir=" + profile,
        "--window-size=1280,900", "--virtual-time-budget=%d" % budget,
        "--screenshot=" + path,
        "%s/?demo=%d" % (BASE_URL, step),
    ]
    try:
        subprocess.run(args, capture_output=True, timeout=120)
    except subprocess.TimeoutExpired:
        log("  步骤 %d：浏览器超时，继续等待文件..." % step)
    # Chrome 为 GUI 子系统程序，命令返回后截图可能延迟落盘
    for _ in range(80):
        if os.path.exists(path) and os.path.getsize(path) > 10000:
            return True
        time.sleep(0.5)
    return False


def add_caption(img, idx, total, text):
    """在画面下方扩展一条字幕带（不遮挡界面内容）。"""
    w, h = img.size
    canvas = Image.new("RGB", (w, h + BAR_H), COLOR_WHITE)
    canvas.paste(img, (0, 0))
    d = ImageDraw.Draw(canvas)
    d.rectangle([0, h, w, h + BAR_H], fill=COLOR_BG)
    d.text((26, h + 17), "%02d / %02d" % (idx, total),
           font=load_font(20), fill=(190, 214, 242))
    d.text((168, h + 14), text, font=load_font(26), fill=COLOR_WHITE)
    return canvas


def cover_frame(w, h, title, subtitle, footer):
    """片头 / 片尾整屏文字帧。"""
    im = Image.new("RGB", (w, h), COLOR_BG)
    d = ImageDraw.Draw(im)
    f_big, f_mid, f_small = load_font(46), load_font(26), load_font(20)
    def center(text, y, font, fill):
        box = d.textbbox((0, 0), text, font=font)
        d.text(((w - (box[2] - box[0])) / 2, y), text, font=font, fill=fill)
    center(title, h * 0.33, f_big, COLOR_WHITE)
    center(subtitle, h * 0.33 + 82, f_mid, (205, 224, 246))
    center(footer, h * 0.33 + 150, f_small, (170, 198, 232))
    return im


def to_bgr(img):
    return np.array(img.convert("RGB"))[:, :, ::-1]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--keep-frames", action="store_true", help="保留中间截图")
    ap.add_argument("--reuse-frames", action="store_true", help="复用已截取的帧，不重新启动浏览器")
    ap.add_argument("--scale", type=float, default=1.0, help="画面缩放比例，如 0.75 生成轻量版")
    ap.add_argument("--fps", type=int, default=FPS, help="输出帧率（静态画面 15 即可）")
    ap.add_argument("--out", default=None, help="输出 mp4 路径")
    args = ap.parse_args()

    out_path = args.out or OUT_MP4
    fps = args.fps
    cross = max(1, int(0.5 * fps))

    if not args.reuse_frames:
        if not check_service():
            log("错误：服务未运行。请先在 D:\\rag 执行  python src/api/server.py")
            return 1
        chrome = find_chrome()
        log("浏览器：%s" % chrome)
        log("服务：%s 可访问" % BASE_URL)
    else:
        chrome = None
        log("复用已截取的界面帧，跳过浏览器截图")

    os.makedirs(FRAME_DIR, exist_ok=True)
    frames = []
    total = len(STEPS)
    for i, (step, caption, _dur) in enumerate(STEPS, 1):
        path = os.path.join(FRAME_DIR, "frame%02d.png" % i)
        cached = (args.reuse_frames and os.path.exists(path)
                  and os.path.getsize(path) > 10000)
        ok = True if cached else capture(chrome, step, path)
        if not ok:
            log("错误：步骤 %d 截图失败，已中止。" % step)
            return 1
        img = Image.open(path).convert("RGB")
        frames.append(add_caption(img, i, total, caption))
        log("  [%02d/%02d] 步骤 %-2d %s %s" %
            (i, total, step, "复用" if cached else "已截取", img.size))

    if args.scale != 1.0:
        w0, h0 = frames[0].size
        tw = int(w0 * args.scale) // 2 * 2
        th = int(h0 * args.scale) // 2 * 2
        frames = [f.resize((tw, th), Image.LANCZOS) for f in frames]
        log("已缩放至 %d x %d" % (tw, th))

    w, h = frames[0].size
    log("画面尺寸：%d x %d" % (w, h))

    clips = [cover_frame(w, h,
                         "基于RAG的烹饪知识图谱构建与推理系统",
                         "系统功能演示（无声）",
                         "烹饪知识图谱 · 图谱增强RAG问答 · 食材替代与禁忌推理 · 可视化交互 · 后台管理")]
    clips += frames
    clips.append(cover_frame(w, h,
                             "演示结束 · 感谢观看",
                             "742 节点 / 13002 边 / 600 条菜谱",
                             "对比实验：图谱增强 88.5% vs 纯文本 38.5%（+50.0pt），幻觉率 0"))

    # 时长分配：片头/片尾各 4 秒，中间按分镜设定
    durations = [4.0] + [d for _s, _c, d in STEPS] + [4.0]

    fourcc = cv2.VideoWriter_fourcc(*"mp4v")
    writer = cv2.VideoWriter(out_path, fourcc, fps, (w, h))
    if not writer.isOpened():
        log("错误：无法创建 MP4 编码器")
        return 1

    written = 0
    for i, img in enumerate(clips):
        base = to_bgr(img)
        for _ in range(int(durations[i] * fps)):
            writer.write(base); written += 1
        if i < len(clips) - 1:
            nxt = to_bgr(clips[i + 1])
            for k in range(1, cross + 1):
                a = k / (cross + 1.0)
                writer.write(cv2.addWeighted(base, 1 - a, nxt, a, 0))
                written += 1
    writer.release()

    # 清理无头浏览器 profile；帧目录按参数保留
    for p in glob.glob(os.path.join(ROOT, "_chrome_*")):
        if os.path.isdir(p):
            shutil.rmtree(p, ignore_errors=True)
    if not args.keep_frames and not args.reuse_frames:
        shutil.rmtree(FRAME_DIR, ignore_errors=True)

    size_mb = os.path.getsize(out_path) / 1024 / 1024
    log("已生成：%s" % out_path)
    log("帧数 %d | 时长约 %.1f 秒 | %d x %d | %.2f MB | 无音轨" %
        (written, written / fps, w, h, size_mb))
    return 0


if __name__ == "__main__":
    sys.exit(main())
