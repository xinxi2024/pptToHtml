**基本用法**

```bash
python3 pptx_to_html.py 演示文稿.pptx
```

**完整选项**

```bash
# 指定输出路径
python3 pptx_to_html.py 演示文稿.pptx -o 输出文件.html

# 提高图片清晰度（文件会变大）
python3 pptx_to_html.py 演示文稿.pptx --dpi 150 --quality 90

# 不处理视频（更快）
python3 pptx_to_html.py 演示文稿.pptx --no-video
```

**依赖安装（Ubuntu/Debian）**

```bash
sudo apt install libreoffice poppler-utils ffmpeg
```

**脚本做了什么**

1. **LibreOffice** 将 PPT/PPTX 渲染成 PDF（完整保留字体、颜色、布局）
2. **pdftoppm** 把每页 PDF 导出为 JPEG 图片
3. **解析 PPTX 关系文件**找到视频，用 **FFmpeg** 转码为 H.264/AAC（解决黑屏问题）
4. 全部资源 base64 内嵌，输出单一独立 HTML 文件，无需网络、无需服务器

HTML 功能与之前一致：全屏无边距、自动隐藏控件、键盘/点击/滑动翻页、缩略图面板、视频一键播放。