这是一个功能强大的PPT转HTML工具。让我帮您正确使用它：

## 使用步骤：

### 1. 保存脚本文件
将您提供的代码保存为 `pptx_to_html.py` 文件：
```bash
# 在您的 pptTohtml 目录下创建这个文件
notepad pptx_to_html.py
```
然后将完整代码粘贴进去并保存。

### 2. 安装依赖工具

**在 Windows 上需要安装：**

#### LibreOffice（用于PPT转PDF）
- 下载地址：https://www.libreoffice.org/download/download-libreoffice/
- 安装后确保 `soffice.exe` 在系统PATH中

#### poppler-utils（用于PDF转图片）
- Windows 版本：https://github.com/oschwartz10612/poppler-windows/releases/
- 下载后解压，将 `bin` 目录添加到系统PATH

#### FFmpeg（用于视频转码）
- 下载地址：https://ffmpeg.org/download.html
- 解压后，将 `bin` 目录添加到系统PATH

### 3. 运行转换命令

```bash
# 基本用法（最简单）
python pptx_to_html.py "实习汇报ppt（5-20）.pptx"

# 指定输出文件名
python pptx_to_html.py "实习汇报ppt（5-20）.pptx" -o "我的汇报.html"

# 调整图片质量（DPI和压缩质量）
python pptx_to_html.py "实习汇报ppt（5-20）.pptx" --dpi 150 --quality 90

# 如果不想处理视频（加快速度）
python pptx_to_html.py "实习汇报ppt（5-20）.pptx" --no-video

# 查看帮助
python pptx_to_html.py -h
```

## 如果遇到问题，这里是简化版脚本（不依赖外部工具）：

```python
# simple_convert.py - 简化版，只提取文本
from pptx import Presentation
import sys
import os

if len(sys.argv) != 3:
    print("用法: python simple_convert.py <PPT文件> <输出HTML>")
    sys.exit(1)

ppt_file = sys.argv[1]
html_file = sys.argv[2]

if not os.path.exists(ppt_file):
    print(f"错误：找不到文件 {ppt_file}")
    sys.exit(1)

# 读取PPT
prs = Presentation(ppt_file)

# 生成HTML
html_content = '''<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>PPT转HTML</title>
    <style>
        body { font-family: Arial, sans-serif; max-width: 1200px; margin: 0 auto; padding: 20px; background: #f0f0f0; }
        .slide { background: white; margin: 20px 0; padding: 30px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }
        .slide-title { font-size: 24px; font-weight: bold; color: #333; border-bottom: 2px solid #4CAF50; padding-bottom: 10px; margin-bottom: 20px; }
        .slide-content { font-size: 16px; line-height: 1.6; }
        .slide-number { text-align: center; color: #999; margin-top: 20px; font-size: 12px; }
        p { margin: 10px 0; }
    </style>
</head>
<body>
'''

for idx, slide in enumerate(prs.slides, 1):
    html_content += f'<div class="slide">'
    html_content += f'<div class="slide-title">第 {idx} 页</div>'
    html_content += f'<div class="slide-content">'
    
    for shape in slide.shapes:
        if hasattr(shape, "text") and shape.text.strip():
            # 处理文本格式
            text = shape.text.replace('\n', '<br>')
            html_content += f'<p>{text}</p>'
    
    html_content += f'</div>'
    html_content += f'<div class="slide-number">第 {idx} 页 / 共 {len(prs.slides)} 页</div>'
    html_content += f'</div>'

html_content += '''</body>
</html>'''

# 保存文件
with open(html_file, 'w', encoding='utf-8') as f:
    f.write(html_content)

print(f"✅ 转换完成！")
print(f"📊 共 {len(prs.slides)} 页")
print(f"🌐 输出文件: {html_file}")
```

运行简化版：
```bash
# 先安装python-pptx
pip install python-pptx

# 运行转换
python simple_convert.py "实习汇报ppt（5-20）.pptx" "1.html"
```

## 推荐方案：

对于您的需求，我建议：

1. **如果只需要文本内容**：使用简化版（`simple_convert.py`），最简单快捷
2. **如果需要保持原有样式和图片**：使用您提供的完整版（需要安装依赖工具）
3. **如果PPT包含视频**：使用完整版并确保安装FFmpeg
