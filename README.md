# PPT Image Restorer

将模糊的 PPT 图片转化为清晰可编辑的 PPT 文件。

## 功能特点

- **OCR 文字识别**: 使用 PaddleOCR 本地识别，无需 API 密钥
- **字体格式提取**: 识别并保存文字的字体、大小、颜色、样式等信息
- **图像智能修复**: 使用 Stable Diffusion WebUI API 去除模糊文字
- **PPT 格式复现**: 使用 python-pptx 生成可编辑的 PPT，复原文字格式
- **WPS 兼容**: 字体映射针对 WPS 进行了优化

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

### 3. 安装图像修复模型（可选）

推荐安装以下模型以获得更好的修复效果：
- `inpaint_only` - 专门用于修复的模型
- `inpaint` - 标准修复模型

## 使用方法

### 命令行使用

```bash
# 处理单张图片
python -m src.main --input slide1.png --output result.pptx

# 处理文件夹（每张图片对应一页）
python -m src.main --input ./slides/ --output presentation.pptx

# 指定 SD API 地址
python -m src.main --input slide.png --sd_api_url http://192.168.1.100:7860

# 保存中间结果（修复后的图片）
python -m src.main --input slide.png --output result.pptx --output_intermediate

# 静默模式
python -m src.main --input slide.png --quiet
```

### 命令行参数

| 参数 | 说明 | 默认值 |
|------|------|--------|
| `--input`, `-i` | 输入图片路径或文件夹 | 必填 |
| `--output`, `-o` | 输出 PPT 文件路径 | 与输入同名.pptx |
| `--sd_api_url` | SD WebUI API 地址 | http://127.0.0.1:7860 |
| `--config`, `-c` | 配置文件路径 | - |
| `--output_intermediate` | 保存中间结果 | False |
| `--intermediate_dir` | 中间结果目录 | ./output_intermediate |
| `--verbose`, `-v` | 详细输出 | False |
| `--quiet`, `-q` | 静默模式 | False |

### 配置文件

支持 TOML 或 JSON 格式的配置文件：

```toml
# config.toml 示例

# 全局设置
language = "zh"
output_intermediate = false

[ocr]
use_angle_cls = true
lang = "ch"
use_gpu = true

[inpaint]
sd_api_url = "http://127.0.0.1:7860"
prompt = "clean background, no text, preserve original style"
negative_prompt = "text, watermark, blurry, low quality"
steps = 20
cfg_scale = 7.0
denoising_strength = 0.75

[ppt]
slide_width = 10.0
slide_height = 5.625
default_font = "微软雅黑"
default_font_size = 18
```

### Python API

```python
from src.main import PPTImageRestorer
from src.config import ConfigManager, SkillConfig

# 创建配置
config = SkillConfig()
config.input_path = "slide.png"
config.output_path = "result.pptx"

# 创建处理器
restorer = PPTImageRestorer(config)

# 处理单张图片
result = restorer.process_single_image("slide.png", "result.pptx")

# 处理多张图片
results = restorer.process_multiple_images("./slides/", "presentation.pptx")

# 打印摘要
restorer.print_summary()

# 关闭资源
restorer.close()
```

## OCR 提取的文字格式信息

`TextFormat` 类包含以下格式信息：

| 属性 | 类型 | 说明 |
|------|------|------|
| `text` | str | 文字内容 |
| `font_size` | int | 字体大小（磅） |
| `font_name` | str | 字体名称 |
| `font_color_rgb` | tuple | RGB 颜色 |
| `font_color_hex` | str | 十六进制颜色 |
| `bold` | bool | 是否粗体 |
| `italic` | bool | 是否斜体 |
| `underline` | bool | 是否下划线 |
| `alignment` | str | 对齐方式 (left/center/right/justify) |
| `line_spacing` | float | 行间距 |
| `bullet` | str | 项目符号 |
| `bbox` | tuple | 边界框 [x1, y1, x2, y2] |
| `text_type` | str | 文字类型 (title/subtitle/body/bullet/footer/header) |

## 故障排除

### PaddleOCR 下载模型失败

```bash
# 手动下载模型
# 模型会保存在 ~/.paddleocr/ 或当前目录
```

### SD WebUI API 连接失败

1. 确保 SD WebUI 已启动并启用 `--api` 参数
2. 检查防火墙设置
3. 确认 API 地址正确

### 文字颜色识别不准确

颜色识别基于图像分析，可能存在误差。可以在生成的 PPT 中手动调整。

### PPT 在 WPS 中显示异常

- 确保系统安装了相应的字体
- 检查字体映射设置

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

## License

MIT License

## 贡献

欢迎提交 Issue 和 Pull Request！
