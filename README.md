# ppt-imgae

skill which can convert image to ppt

## 功能特点

- **OCR 文字识别**: 使用 PaddleOCR 本地识别，无需 API 密钥
- **字体格式提取**: 识别并保存文字的字体、大小、颜色、样式等信息
- **图像智能修复**: 使用 Stable Diffusion WebUI API 去除模糊文字
- **PPT 格式复现**: 使用 python-pptx 生成可编辑的 PPT，复原文字格式
- **WPS 兼容**: 字体映射针对 WPS 进行了优化

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
```

## 项目结构

```
ppt-image-restorer/
├── SKILL.md                 # Skill 定义文件
├── src/
│   ├── __init__.py
│   ├── main.py              # 主入口
│   ├── config.py            # 配置管理
│   ├── ocr_processor.py     # OCR 处理器
│   ├── image_inpainter.py   # 图像修复
│   └── ppt_generator.py     # PPT 生成器
├── requirements.txt
├── README.md
└── config_example.toml
```

## 贡献

欢迎提交 Issue 和 Pull Request！
