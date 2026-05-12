# ⚠️ Disclaimer

**This repository was completely created by an AI agent without any human review or verification.**

Use at your own risk.

---

# ppt-imgae

Convert blurry PPT images into clear, editable PPT files.

![Stars](https://img.shields.io/github/stars/SunmayK/ppt-imgae)
![Forks](https://img.shields.io/github/forks/SunmayK/ppt-imgae)
![Issues](https://img.shields.io/github/issues/SunmayK/ppt-imgae)

## Features

- ✨ **OCR Text Recognition** - Uses PaddleOCR for local recognition, no API key required
- 🚀 **Font Format Extraction** - Identifies and preserves font name, size, color, and style
- 🔒 **Image Inpainting** - Uses Stable Diffusion WebUI API to remove blurry text
- 📦 **PPT Format Reproduction** - Generates editable PPT using python-pptx
- 🎯 **WPS Compatible** - Font mapping optimized for WPS

## Workflow

```
Input Image → OCR Text & Format Recognition → SD Inpainting Remove Text → Generate Clear PPT
```

## Installation

### 1. Install Python Dependencies

```bash
# GPU accelerated (recommended)
pip install paddlepaddle-gpu paddleocr

# CPU version
pip install paddlepaddle paddleocr

# Other dependencies
pip install python-pptx Pillow numpy requests toml
```

### 2. Configure Stable Diffusion WebUI

Ensure Stable Diffusion WebUI is installed and API is enabled:

```bash
# Start with --api parameter
python launch.py --api --api-server-only
```

## Usage

### Command Line

```bash
# Process single image
python -m src.main --input slide1.png --output result.pptx

# Process folder (each image = one slide)
python -m src.main --input ./slides/ --output presentation.pptx

# Specify SD API address
python -m src.main --input slide.png --sd_api_url http://192.168.1.100:7860
```

### Command Line Arguments

| Argument | Description | Default |
|----------|-------------|---------|
| `--input`, `-i` | Input image path or folder | Required |
| `--output`, `-o` | Output PPT file path | Same as input with .pptx |
| `--sd_api_url` | SD WebUI API address | http://127.0.0.1:7860 |
| `--config`, `-c` | Config file path | - |
| `--output_intermediate` | Save intermediate results | False |
| `--verbose`, `-v` | Verbose output | False |
| `--quiet`, `-q` | Quiet mode | False |

## OCR Extracted Text Format

| Property | Type | Description |
|----------|------|-------------|
| `font_size` | int | Font size (points) |
| `font_name` | str | Font name |
| `font_color_rgb` | tuple | RGB color |
| `bold` / `italic` | bool | Bold/Italic |
| `alignment` | str | Text alignment |
| `text_type` | str | Text type (title/body/bullet etc.) |

## Project Structure

```
ppt-imgae/
├── SKILL.md                 # Skill definition file
├── src/
│   ├── __init__.py
│   ├── main.py              # Main entry
│   ├── config.py            # Configuration
│   ├── ocr_processor.py     # OCR processor
│   ├── image_inpainter.py   # Image inpainting
│   └── ppt_generator.py     # PPT generator
├── requirements.txt
└── README.md
```

## Troubleshooting

### PaddleOCR Model Download Failed

Models are saved to `~/.paddleocr/` or current directory. They will be downloaded automatically on first run.

### SD WebUI API Connection Failed

1. Ensure SD WebUI is running with `--api` parameter enabled
2. Check firewall settings
3. Verify API address is correct

## Contributing

Issues and Pull Requests are welcome!
