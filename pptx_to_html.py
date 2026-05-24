#!/usr/bin/env python3
"""
pptx_to_html.py — 将 PPT/PPTX 转换为全屏 HTML 演示文件

功能：
  · 完整保留每页视觉效果（矢量渲染）
  · 自动检测并嵌入视频（转码为 H.264/AAC，确保浏览器兼容）
  · 全屏无边距展示，支持：键盘翻页、点击翻页、手机滑动、缩略图导航
  · 所有资源内嵌为 base64，生成单一独立 .html 文件

依赖（Ubuntu/Debian）：
  · LibreOffice  : sudo apt install libreoffice
  · Poppler      : sudo apt install poppler-utils
  · FFmpeg       : sudo apt install ffmpeg
  · Python 3.8+  : 标准库即可，无需 pip 安装

用法：
  python3 pptx_to_html.py 演示文稿.pptx
  python3 pptx_to_html.py 演示文稿.pptx -o 输出.html
  python3 pptx_to_html.py 演示文稿.pptx --dpi 150 --quality 85
  python3 pptx_to_html.py 演示文稿.pptx --no-video   # 跳过视频处理
"""

import argparse
import base64
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path


# ──────────────────────────────────────────────
# 工具函数
# ──────────────────────────────────────────────

def run(cmd: list, **kw) -> subprocess.CompletedProcess:
    """运行命令，出错时打印并退出。"""
    result = subprocess.run(cmd, capture_output=True, text=True, **kw)
    if result.returncode != 0:
        print(f"[ERROR] 命令失败: {' '.join(str(c) for c in cmd)}", file=sys.stderr)
        print(result.stderr[-2000:] if result.stderr else "", file=sys.stderr)
        sys.exit(1)
    return result


def get_soffice_env() -> dict:
    """LibreOffice 在沙盒环境下需要特殊环境变量。"""
    env = os.environ.copy()
    env["SAL_USE_VCLPLUGIN"] = "svp"
    return env


def find_soffice() -> str:
    for name in ("soffice", "libreoffice"):
        path = shutil.which(name)
        if path:
            return path
    print("[ERROR] 未找到 LibreOffice (soffice)，请先安装：sudo apt install libreoffice", file=sys.stderr)
    sys.exit(1)


def require_tool(name: str, install_hint: str):
    if not shutil.which(name):
        print(f"[ERROR] 未找到 {name}，请先安装：{install_hint}", file=sys.stderr)
        sys.exit(1)


def b64(path: Path) -> str:
    return base64.b64encode(path.read_bytes()).decode()


def js_escape(s: str) -> str:
    return json.dumps(s)


# ──────────────────────────────────────────────
# 步骤 1：PPTX → PDF
# ──────────────────────────────────────────────

def pptx_to_pdf(pptx_path: Path, workdir: Path) -> Path:
    soffice = find_soffice()
    print(f"[1/4] 渲染幻灯片 → PDF …")
    env = get_soffice_env()
    subprocess.run(
        [soffice, "--headless", "--convert-to", "pdf",
         "--outdir", str(workdir), str(pptx_path.resolve())],
        env=env, capture_output=True
    )
    # LibreOffice 输出文件名 = 原文件名.pdf
    pdf = workdir / (pptx_path.stem + ".pdf")
    if not pdf.exists():
        # 有时 LibreOffice 会加奇怪前缀，找一下
        found = list(workdir.glob("*.pdf"))
        if not found:
            print("[ERROR] LibreOffice 未生成 PDF", file=sys.stderr)
            sys.exit(1)
        pdf = found[0]
    return pdf


# ──────────────────────────────────────────────
# 步骤 2：PDF → JPEG（每页一张）
# ──────────────────────────────────────────────

def pdf_to_images(pdf_path: Path, workdir: Path, dpi: int, quality: int) -> list[Path]:
    require_tool("pdftoppm", "sudo apt install poppler-utils")
    print(f"[2/4] 导出幻灯片图片 (dpi={dpi}) …")
    prefix = workdir / "slide"
    subprocess.run(
        ["pdftoppm", "-jpeg", f"-r", str(dpi),
         f"-jpegopt", f"quality={quality}",
         str(pdf_path), str(prefix)],
        capture_output=True
    )
    slides = sorted(workdir.glob("slide-*.jpg"))
    if not slides:
        slides = sorted(workdir.glob("slide*.jpg"))
    if not slides:
        print("[ERROR] pdftoppm 未生成图片", file=sys.stderr)
        sys.exit(1)
    print(f"       共 {len(slides)} 页")
    return slides


# ──────────────────────────────────────────────
# 步骤 3：提取视频并转码
# ──────────────────────────────────────────────

def extract_videos(pptx_path: Path, workdir: Path, skip_video: bool) -> dict[int, Path]:
    """
    返回 {slide_index_0based: h264_mp4_path}
    """
    if skip_video:
        return {}

    print("[3/4] 检测并提取视频 …")

    media_dir = workdir / "media"
    media_dir.mkdir(exist_ok=True)

    # 解压 PPTX（就是个 zip）
    unpack_dir = workdir / "pptx_unpack"
    with zipfile.ZipFile(pptx_path, "r") as z:
        z.extractall(unpack_dir)

    slides_dir = unpack_dir / "ppt" / "slides"
    rels_dir = slides_dir / "_rels"

    if not rels_dir.exists():
        print("       未找到关系文件，跳过视频")
        return {}

    # 遍历每个幻灯片的 .rels，找 mp4/video 关系
    video_map: dict[int, Path] = {}

    rels_files = sorted(rels_dir.glob("slide*.xml.rels"))
    for rels_file in rels_files:
        # 从文件名提取幻灯片编号（1-based）
        m = re.search(r"slide(\d+)\.xml\.rels", rels_file.name)
        if not m:
            continue
        slide_num = int(m.group(1))  # 1-based

        content = rels_file.read_text(encoding="utf-8", errors="ignore")

        # 找所有视频/media 关系
        targets = re.findall(
            r'Type="[^"]*(?:video|media)[^"]*"\s+Target="([^"]+)"',
            content, re.IGNORECASE
        )
        # 去重，只取 mp4
        mp4_targets = [t for t in targets if t.lower().endswith(".mp4")]
        if not mp4_targets:
            continue

        # 取第一个视频
        rel_target = mp4_targets[0]
        # 路径是相对于 ppt/slides/ 的，../media/mediaX.mp4
        src_path = (slides_dir / rel_target).resolve()
        if not src_path.exists():
            # 尝试直接在 media 目录查找
            media_name = Path(rel_target).name
            src_path = unpack_dir / "ppt" / "media" / media_name
        if not src_path.exists():
            print(f"       警告：找不到视频文件 {rel_target}")
            continue

        # 转码为 H.264 / AAC
        out_path = media_dir / f"slide{slide_num}.mp4"
        print(f"       转码视频 (第{slide_num}页): {src_path.name} → H.264 …")
        result = subprocess.run(
            ["ffmpeg", "-y", "-i", str(src_path),
             "-c:v", "libx264", "-preset", "fast", "-crf", "28",
             "-profile:v", "baseline", "-level", "3.0",
             "-c:a", "aac", "-b:a", "96k",
             "-movflags", "+faststart",
             str(out_path)],
            capture_output=True
        )
        if result.returncode != 0 or not out_path.exists():
            print(f"       警告：视频转码失败，跳过 (第{slide_num}页)")
            continue

        video_map[slide_num - 1] = out_path  # 转为 0-based

    if video_map:
        print(f"       发现 {len(video_map)} 个视频")
    else:
        print("       未发现视频")

    return video_map


# ──────────────────────────────────────────────
# 步骤 4：组装 HTML
# ──────────────────────────────────────────────

HTML_TEMPLATE = """\
<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, user-scalable=no">
<title>{title}</title>
<style>
  *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}

  html, body {{
    width: 100%; height: 100%;
    background: #000;
    overflow: hidden;
    font-family: 'Segoe UI', 'PingFang SC', 'Microsoft YaHei', sans-serif;
    touch-action: none;
    user-select: none;
  }}

  /* ── 全屏舞台 ── */
  #stage {{
    position: fixed;
    inset: 0;
    display: flex;
    align-items: center;
    justify-content: center;
    background: #000;
  }}
  #stage img, #stage video {{
    position: absolute;
    inset: 0;
    width: 100%;
    height: 100%;
    object-fit: contain;
    display: block;
  }}
  #stage video {{ display: none; }}
  #stage.show-video img {{ display: none; }}
  #stage.show-video video {{ display: block; }}

  /* ── HUD（自动隐藏）── */
  #hud {{
    position: fixed;
    inset: 0;
    pointer-events: none;
    transition: opacity 0.4s;
    z-index: 10;
  }}
  #hud.hidden {{ opacity: 0; }}

  /* 进度条 */
  #progressBar {{
    position: absolute;
    top: 0; left: 0;
    height: 3px;
    background: #2980b9;
    transition: width 0.25s;
    pointer-events: none;
  }}

  /* 页码 */
  #counter {{
    position: absolute;
    top: 10px;
    left: 50%;
    transform: translateX(-50%);
    background: rgba(0,0,0,0.55);
    color: #fff;
    font-size: 13px;
    padding: 4px 14px;
    border-radius: 20px;
    white-space: nowrap;
    pointer-events: none;
    backdrop-filter: blur(4px);
  }}

  /* 底部导航 */
  #navBar {{
    position: absolute;
    bottom: 22px;
    left: 50%;
    transform: translateX(-50%);
    display: flex;
    align-items: center;
    gap: 16px;
    pointer-events: all;
  }}
  .nav-btn {{
    background: rgba(0,0,0,0.6);
    border: 1px solid rgba(255,255,255,0.25);
    color: #fff;
    width: 44px; height: 44px;
    border-radius: 50%;
    font-size: 20px;
    cursor: pointer;
    display: flex; align-items: center; justify-content: center;
    transition: background 0.2s;
    backdrop-filter: blur(6px);
  }}
  .nav-btn:hover {{ background: rgba(41,128,185,0.75); }}
  .nav-btn:disabled {{ opacity: 0.25; cursor: default; }}

  /* 视频按钮（右下） */
  #videoBadge {{
    position: absolute;
    bottom: 22px;
    right: 22px;
    background: rgba(41,128,185,0.85);
    color: #fff;
    font-size: 13px;
    padding: 8px 16px;
    border-radius: 24px;
    cursor: pointer;
    display: none;
    align-items: center;
    gap: 6px;
    pointer-events: all;
    backdrop-filter: blur(6px);
    border: 1px solid rgba(255,255,255,0.2);
    transition: background 0.2s;
  }}
  #videoBadge:hover {{ background: rgba(41,128,185,1); }}
  #videoBadge svg {{ width: 16px; height: 16px; fill: #fff; flex-shrink: 0; }}

  /* 缩略图按钮（左下） */
  #thumbToggle {{
    position: absolute;
    bottom: 22px;
    left: 22px;
    background: rgba(0,0,0,0.6);
    border: 1px solid rgba(255,255,255,0.25);
    color: #fff;
    width: 44px; height: 44px;
    border-radius: 50%;
    font-size: 18px;
    cursor: pointer;
    display: flex; align-items: center; justify-content: center;
    pointer-events: all;
    backdrop-filter: blur(6px);
    transition: background 0.2s;
  }}
  #thumbToggle:hover {{ background: rgba(41,128,185,0.75); }}

  /* 缩略图面板 */
  #thumbPanel {{
    position: fixed;
    bottom: 0; left: 0; right: 0;
    background: rgba(10,10,20,0.92);
    backdrop-filter: blur(10px);
    padding: 12px 12px 24px;
    z-index: 20;
    transform: translateY(100%);
    transition: transform 0.3s ease;
    border-top: 1px solid rgba(255,255,255,0.1);
  }}
  #thumbPanel.open {{ transform: translateY(0); }}
  #thumbScroll {{
    display: flex;
    gap: 8px;
    overflow-x: auto;
    padding-bottom: 4px;
    scroll-behavior: smooth;
  }}
  #thumbScroll::-webkit-scrollbar {{ height: 4px; }}
  #thumbScroll::-webkit-scrollbar-thumb {{ background: #444; border-radius: 2px; }}
  .thumb {{
    flex: 0 0 auto;
    width: 120px;
    aspect-ratio: 16/9;
    border-radius: 5px;
    overflow: hidden;
    cursor: pointer;
    border: 2px solid transparent;
    position: relative;
    transition: border-color 0.15s, transform 0.15s;
  }}
  .thumb:hover {{ transform: scale(1.05); }}
  .thumb.active {{ border-color: #2980b9; }}
  .thumb img {{ width: 100%; height: 100%; object-fit: cover; display: block; }}
  .thumb-num {{
    position: absolute; bottom: 2px; right: 4px;
    font-size: 10px; color: rgba(255,255,255,0.75);
    background: rgba(0,0,0,0.5); padding: 1px 4px; border-radius: 3px;
  }}
  .thumb-vid {{
    position: absolute; top: 3px; left: 3px;
    background: rgba(41,128,185,0.85);
    border-radius: 3px; padding: 1px 5px;
    font-size: 9px; color: #fff;
  }}
  #thumbClose {{
    position: absolute; top: 8px; right: 12px;
    background: none; border: none;
    color: rgba(255,255,255,0.5);
    font-size: 20px; cursor: pointer; line-height: 1;
    pointer-events: all;
  }}
  #thumbClose:hover {{ color: #fff; }}

  /* 全屏提示（首次加载） */
  #fsHint {{
    position: fixed;
    top: 50%; left: 50%;
    transform: translate(-50%, -50%);
    background: rgba(0,0,0,0.7);
    color: #fff;
    padding: 16px 28px;
    border-radius: 12px;
    font-size: 14px;
    text-align: center;
    z-index: 50;
    backdrop-filter: blur(8px);
    transition: opacity 0.5s;
    pointer-events: none;
  }}
  #fsHint.fade {{ opacity: 0; }}
</style>
</head>
<body>

<div id="stage">
  <img id="mainImg" src="" alt="">
  <video id="mainVideo" playsinline preload="none"></video>
</div>

<div id="hud">
  <div id="progressBar"></div>
  <div id="counter">1 / {total}</div>

  <div id="navBar">
    <button class="nav-btn" id="prevBtn" onclick="go(-1)" title="上一页 (←)">&#8592;</button>
    <button class="nav-btn" id="nextBtn" onclick="go(1)"  title="下一页 (→)">&#8594;</button>
  </div>

  <div id="videoBadge" onclick="toggleVideo()">
    <svg id="vicon" viewBox="0 0 24 24"><path d="M8 5v14l11-7z"/></svg>
    <span id="vtext">播放视频</span>
  </div>

  <button id="thumbToggle" onclick="toggleThumb()" title="幻灯片概览">&#9783;</button>
</div>

<div id="thumbPanel">
  <button id="thumbClose" onclick="toggleThumb()">&#10005;</button>
  <div id="thumbScroll"></div>
</div>

<div id="fsHint">点击任意位置 · 左右箭头翻页 · 左下角查看全部幻灯片</div>

<script>
const SLIDES = {slides_json};

let cur = 0, vidActive = false, thumbOpen = false, hudTimer = null;
let touchX = 0, touchY = 0;

const stage    = document.getElementById('stage');
const mainImg  = document.getElementById('mainImg');
const mainVid  = document.getElementById('mainVideo');
const hud      = document.getElementById('hud');
const counter  = document.getElementById('counter');
const progress = document.getElementById('progressBar');
const badge    = document.getElementById('videoBadge');
const vicon    = document.getElementById('vicon');
const vtext    = document.getElementById('vtext');

function render(idx) {{
  if (idx < 0 || idx >= SLIDES.length) return;
  cur = idx;
  const s = SLIDES[idx];

  mainVid.pause();
  mainVid.src = '';
  vidActive = false;
  stage.classList.remove('show-video');

  mainImg.src = s.img;

  if (s.video) {{
    mainVid.src = s.video;
    badge.style.display = 'flex';
    vicon.innerHTML = '<path d="M8 5v14l11-7z"/>';
    vtext.textContent = '播放视频';
  }} else {{
    badge.style.display = 'none';
  }}

  const pct = ((idx + 1) / SLIDES.length * 100).toFixed(1);
  progress.style.width = pct + '%';
  counter.textContent  = (idx + 1) + ' / ' + SLIDES.length;
  document.getElementById('prevBtn').disabled = idx === 0;
  document.getElementById('nextBtn').disabled = idx === SLIDES.length - 1;

  document.querySelectorAll('.thumb').forEach((t, i) =>
    t.classList.toggle('active', i === idx));
  document.querySelectorAll('.thumb')[idx]
    ?.scrollIntoView({{inline: 'nearest', block: 'nearest'}});

  showHud();
}}

function go(dir) {{ render(cur + dir); }}

function toggleVideo() {{
  if (!SLIDES[cur].video) return;
  if (!vidActive) {{
    vidActive = true;
    stage.classList.add('show-video');
    mainVid.play();
    vicon.innerHTML = '<path d="M6 19h4V5H6v14zm8-14v14h4V5h-4z"/>';
    vtext.textContent = '关闭视频';
  }} else {{
    vidActive = false;
    stage.classList.remove('show-video');
    mainVid.pause();
    vicon.innerHTML = '<path d="M8 5v14l11-7z"/>';
    vtext.textContent = '播放视频';
  }}
}}

function showHud() {{
  hud.classList.remove('hidden');
  clearTimeout(hudTimer);
  hudTimer = setTimeout(() => {{
    if (!thumbOpen) hud.classList.add('hidden');
  }}, 3000);
}}

function toggleThumb() {{
  thumbOpen = !thumbOpen;
  document.getElementById('thumbPanel').classList.toggle('open', thumbOpen);
  if (thumbOpen) {{ hud.classList.remove('hidden'); clearTimeout(hudTimer); }}
  else showHud();
}}

// 构建缩略图
const grid = document.getElementById('thumbScroll');
SLIDES.forEach((s, i) => {{
  const d = document.createElement('div');
  d.className = 'thumb' + (i === 0 ? ' active' : '');
  d.onclick = () => {{ render(i); if (thumbOpen) toggleThumb(); }};
  d.innerHTML =
    `<img src="${{s.img}}" loading="lazy">` +
    `<span class="thumb-num">${{i + 1}}</span>` +
    (s.video ? '<span class="thumb-vid">▶ 视频</span>' : '');
  grid.appendChild(d);
}});

// 键盘
document.addEventListener('keydown', e => {{
  if (['ArrowRight','ArrowDown',' '].includes(e.key)) {{ e.preventDefault(); go(1); }}
  else if (['ArrowLeft','ArrowUp'].includes(e.key))   {{ e.preventDefault(); go(-1); }}
  else if (e.key === 'Escape') {{ if (thumbOpen) toggleThumb(); }}
  else if (e.key.toLowerCase() === 'f') document.documentElement.requestFullscreen?.();
}});

// 滑动
stage.addEventListener('touchstart', e => {{
  touchX = e.touches[0].clientX;
  touchY = e.touches[0].clientY;
}}, {{passive: true}});
stage.addEventListener('touchend', e => {{
  const dx = e.changedTouches[0].clientX - touchX;
  const dy = e.changedTouches[0].clientY - touchY;
  if (Math.abs(dx) > Math.abs(dy) && Math.abs(dx) > 40)
    dx < 0 ? go(1) : go(-1);
  else showHud();
}}, {{passive: true}});

// 点击舞台翻页
stage.addEventListener('click', e => {{
  if (e.target === stage || e.target === mainImg)
    e.clientX / window.innerWidth > 0.5 ? go(1) : go(-1);
}});

// 鼠标移动显示 HUD
document.addEventListener('mousemove', showHud);
document.addEventListener('click', showHud);

// 首次提示消失
const fsHint = document.getElementById('fsHint');
setTimeout(() => {{ fsHint.classList.add('fade'); }}, 3000);
setTimeout(() => {{ fsHint.style.display = 'none'; }}, 3600);

render(0);
</script>
</body>
</html>
"""


def build_html(title: str, slides: list[Path], video_map: dict[int, Path]) -> str:
    print("[4/4] 组装 HTML …")
    slide_objs = []
    for i, img_path in enumerate(slides):
        img_data = f"data:image/jpeg;base64,{b64(img_path)}"
        if i in video_map:
            vid_data = f"data:video/mp4;base64,{b64(video_map[i])}"
            slide_objs.append({"img": img_data, "video": vid_data})
        else:
            slide_objs.append({"img": img_data, "video": None})

    slides_json = json.dumps(slide_objs, ensure_ascii=False, separators=(",", ":"))

    return HTML_TEMPLATE.format(
        title=title,
        total=len(slides),
        slides_json=slides_json,
    )


# ──────────────────────────────────────────────
# 主程序
# ──────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="将 PPT/PPTX 转换为全屏 HTML 演示文件",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__
    )
    parser.add_argument("input", help="输入文件路径 (.pptx / .ppt)")
    parser.add_argument("-o", "--output", help="输出 HTML 路径（默认：同目录同名.html）")
    parser.add_argument("--dpi",     type=int, default=120, help="幻灯片图片分辨率 (默认 120)")
    parser.add_argument("--quality", type=int, default=82,  help="JPEG 质量 1-100 (默认 82)")
    parser.add_argument("--no-video", action="store_true",  help="跳过视频提取")
    args = parser.parse_args()

    pptx_path = Path(args.input).expanduser().resolve()
    if not pptx_path.exists():
        print(f"[ERROR] 文件不存在: {pptx_path}", file=sys.stderr)
        sys.exit(1)
    if pptx_path.suffix.lower() not in (".pptx", ".ppt"):
        print(f"[WARN] 文件扩展名非 .pptx/.ppt，尝试继续 …")

    out_path = Path(args.output) if args.output else pptx_path.with_suffix(".html")

    title = pptx_path.stem

    with tempfile.TemporaryDirectory(prefix="pptx2html_") as tmp:
        workdir = Path(tmp)

        # 1. 转 PDF
        pdf = pptx_to_pdf(pptx_path, workdir)

        # 2. PDF → 图片
        slides = pdf_to_images(pdf, workdir, args.dpi, args.quality)

        # 3. 提取视频
        video_map = extract_videos(pptx_path, workdir, args.no_video)

        # 4. 组装 HTML
        html = build_html(title, slides, video_map)

    out_path.write_text(html, encoding="utf-8")
    size_mb = out_path.stat().st_size / 1024 / 1024
    print(f"\n✅ 完成！输出文件：{out_path}  ({size_mb:.1f} MB)")
    print(f"   幻灯片数量：{len(slides)}")
    print(f"   嵌入视频数：{len(video_map)}")
    print(f"\n使用说明：")
    print(f"   · 用浏览器直接打开 {out_path.name} 即可")
    print(f"   · 左右箭头键 / 点击屏幕 / 手机左右滑动 翻页")
    print(f"   · 按 F 键进入浏览器全屏模式")
    print(f"   · 左下角 ⊹ 按钮查看所有幻灯片缩略图")


if __name__ == "__main__":
    main()
