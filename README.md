# ppt-imgae

将模糊的 PPT 图片转化为清晰可编辑的 PPT 文件。

![Stars](https://img.shields.io/github/stars/SunmayK/ppt-imgae)
![Forks](https://img.shields.io/github/forks/SunmayK/ppt-imgae)
![Issues](https://img.shields.io/github/issues/SunmayK/ppt-imgae)

## 功能特点

- ✨ **OCR 文字识别** - 使用 PaddleOCR 本地识别，无需 API 密钥
- 🚀 **字体格式提取** - 识别并保存文字的字体、大小、颜色、样式等信息
- 🔒 **图像智能修复** - 使用 Stable Diffusion WebUI API 去除模糊文字
- 📦 **PPT 格式复现** - 使用 python-pptx 生成可编辑的 PPT，复原文字格式
- 🎯 **WPS 兼容** - 字体映射针对 WPS 进行了优化

## 工作流程

```
输入图片 → OCR 识别文字及格式 → SD Inpainting 去除文字 → 生成清晰 PPT
```

## 安装

### 1. 安装 Python 依赖

```bash
# 推荐 GPU 加速
pip install paddlepaddle-gpu paddleocr

# 或 CPU 版本
pip install paddlepaddle paddleocr

# 安装其他依赖
pip install python-pptx Pillow numpy requests toml
```

### 2. 配置 Stable Diffusion WebUI

确保 Stable Diffusion WebUI 已安装并启用 API：

```bash
# 启动时添加 --api 参数
python launch.py --api --api-server-only
```

## 使用方法

### 命令行使用

```bash
# 处理单张图片
python -m src.main --input slide1.png --output result.pptx

# 处理文件夹（每张图片对应一页）
python -m src.main --input ./slides/ --output presentation.pptx

# 指定 SD API 地址
python -m src.main --input slide.png --sd_api_url http://192.168.1.100:7860
```

### 命令行参数

| 参数 | 说明 | 默认值 |
|------|------|--------|
| `--input`, `-i` | 输入图片路径或文件夹 | 必填 |
| `--output`, `-o` | 输出 PPT 文件路径 | 与输入同名.pptx |
| `--sd_api_url` | SD WebUI API 地址 | http://127.0.0.1:7860 |
| `--config`, `-c` | 配置文件路径 | - |
| `--output_intermediate` | 保存中间结果 | False |
| `--verbose`, `-v` | 详细输出 | False |
| `--quiet`, `-q` | 静默模式 | False |

## OCR 提取的文字格式

| 属性 | 类型 | 说明 |
|------|------|------|
| `font_size` | int | 字体大小（磅） |
| `font_name` | str | 字体名称 |
| `font_color_rgb` | tuple | RGB 颜色 |
| `bold` / `italic` | bool | 粗体/斜体 |
| `alignment` | str | 对齐方式 |
| `text_type` | str | 文字类型 (title/body/bullet等) |

## 项目结构

```
ppt-imgae/
├── SKILL.md                 # Skill 定义文件
├── src/
│   ├── __init__.py
│   ├── main.py              # 主入口
│   ├── config.py            # 配置管理
│   ├── ocr_processor.py     # OCR 处理器
│   ├── image_inpainter.py   # 图像修复
│   └── ppt_generator.py     # PPT 生成器
├── requirements.txt
└── README.md
```

## 故障排除

### PaddleOCR 下载模型失败

模型会保存在 `~/.paddleocr/` 或当前目录，首次运行会自动下载。

### SD WebUI API 连接失败

1. 确保 SD WebUI 已启动并启用 `--api` 参数
2. 检查防火墙设置
3. 确认 API 地址正确

## 贡献

欢迎提交 Issue 和 Pull Request！
