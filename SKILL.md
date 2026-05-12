# PPT 图片修复 Skill

## 名称

ppt-image-restorer

## 描述

将模糊的 PPT 图片转化为清晰可编辑的 PPT 文件。通过 OCR 识别图片中的文字及其格式（字体、字号、颜色、样式等），使用 Stable Diffusion 进行图像修复去除模糊文字，最后生成可编辑的 PPT 文件。

## 触发关键词

- ppt修复
- 图片转ppt
- 模糊文字修复
- ppt图片清晰化
- 文字格式提取

## 输入参数

| 参数名 | 类型 | 必填 | 默认值 | 说明 |
|--------|------|------|--------|------|
| input_path | string | 是 | - | 输入图片路径，支持单张图片或文件夹 |
| output_path | string | 否 | 与输入同名.pptx | 输出 PPT 文件路径 |
| sd_api_url | string | 否 | http://127.0.0.1:7860 | Stable Diffusion WebUI API 地址 |
| ppt_template | string | 否 | - | 可选的 PPT 模板路径（用于样式参考） |
| language | string | 否 | auto | 输出语言，auto/zh/en |

## 功能特点

### 1. OCR 文字识别
- 使用 PaddleOCR 本地识别，无需 API 密钥
- 提取文字内容、位置坐标、置信度
- **字体格式分析**：估算字体大小、颜色、样式（粗体/斜体）
- 段落布局分析：文字对齐方式、行间距

### 2. 图像修复
- 调用本地 Stable Diffusion WebUI API
- 智能去除模糊文字区域
- 保留原始背景和图片样式

### 3. PPT 生成
- 使用 python-pptx 生成可编辑 PPT
- 还原文字格式：字体、字号、颜色、加粗、斜体
- 还原段落格式：对齐方式、行间距、项目符号
- 每张图片对应一页幻灯片

## 使用示例

```bash
# 处理单张图片
python -m src.main --input "slide1.png" --output "result.pptx"

# 处理文件夹
python -m src.main --input "./slides/" --output "presentation.pptx"

# 指定 SD API 地址
python -m src.main --input "slide.png" --sd_api_url "http://192.168.1.100:7860"
```

## 依赖要求

- Python 3.8+
- PaddlePaddle / PaddlePaddle-GPU
- PaddleOCR
- python-pptx
- Pillow
- requests
- Stable Diffusion WebUI（需启用 --api 参数）

## 输出格式

生成标准 .pptx 文件，可直接用 WPS、Microsoft PowerPoint 打开编辑。

## 注意事项

1. 确保 Stable Diffusion WebUI 已启动并启用 API（添加 --api 参数）
2. PaddleOCR 首次运行会自动下载模型（约 10MB）
3. 图片中文字颜色通过颜色分析估算，可能存在误差
4. 建议输入图片分辨率不低于 720p 以获得最佳效果
