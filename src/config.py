# -*- coding: utf-8 -*-
"""
配置管理模块
统一管理配置，支持配置文件、命令行参数、环境变量
"""

import os
import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional, List


@dataclass
class OCRConfig:
    """OCR 配置"""
    use_angle_cls: bool = True
    lang: str = "ch"  # ch, en, japan, korean
    det_model_dir: Optional[str] = None
    rec_model_dir: Optional[str] = None
    cls_model_dir: Optional[str] = None
    use_gpu: bool = True
    show_log: bool = False


@dataclass
class InpaintConfig:
    """图像修复配置"""
    sd_api_url: str = "http://127.0.0.1:7860"
    inpaint_model: str = "inpaint"  # inpaint, inpaint_only, etc.
    prompt: str = "clean background, no text, preserve original style"
    negative_prompt: str = "text, watermark, blurry, low quality"
    steps: int = 20
    cfg_scale: float = 7.0
    denoising_strength: float = 0.75
    mask_blur: int = 2
    inpainting_fill: int = 1  # 0: original, 1: latent noise, 2: latent uniform


@dataclass
class PPTConfig:
    """PPT 生成配置"""
    slide_width: float = 10.0  # inches (standard 16:9)
    slide_height: float = 5.625
    default_font: str = "微软雅黑"
    default_font_size: int = 18
    default_font_color: str = "000000"
    margin_left: float = 0.5
    margin_right: float = 0.5
    margin_top: float = 0.5
    margin_bottom: float = 0.5


@dataclass
class SkillConfig:
    """Skill 全局配置"""
    input_path: str = ""
    output_path: str = ""
    ppt_template: Optional[str] = None
    language: str = "auto"  # auto, zh, en
    output_intermediate: bool = False  # 是否输出中间结果（修复后的图片）
    intermediate_dir: str = "./output_intermediate"
    verbose: bool = True
    
    # 子模块配置
    ocr: OCRConfig = field(default_factory=OCRConfig)
    inpaint: InpaintConfig = field(default_factory=InpaintConfig)
    ppt: PPTConfig = field(default_factory=PPTConfig)


class ConfigManager:
    """配置管理器"""
    
    def __init__(self, config_path: Optional[str] = None):
        self.config = SkillConfig()
        self.config_path = config_path
        
        # 尝试加载配置文件
        if config_path and os.path.exists(config_path):
            self._load_from_file(config_path)
        
        # 从环境变量加载
        self._load_from_env()
    
    def _load_from_file(self, path: str):
        """从配置文件加载"""
        path = Path(path)
        if path.suffix == ".toml":
            self._load_from_toml(path)
        elif path.suffix == ".json":
            self._load_from_json(path)
        else:
            raise ValueError(f"Unsupported config format: {path.suffix}")
    
    def _load_from_toml(self, path: Path):
        """从 TOML 文件加载（简化实现）"""
        try:
            import toml
            with open(path, 'r', encoding='utf-8') as f:
                data = toml.load(f)
            self._apply_dict(data)
        except ImportError:
            # 手动解析简单的 TOML
            self._parse_simple_toml(path)
    
    def _load_from_json(self, path: Path):
        """从 JSON 文件加载"""
        with open(path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        self._apply_dict(data)
    
    def _parse_simple_toml(self, path: Path):
        """简化的 TOML 解析（支持基本键值对）"""
        with open(path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        current_section = None
        for line in content.split('\n'):
            line = line.strip()
            if not line or line.startswith('#'):
                continue
            
            if line.startswith('[') and line.endswith(']'):
                current_section = line[1:-1].strip()
                continue
            
            if '=' in line:
                key, value = line.split('=', 1)
                key = key.strip()
                value = value.strip().strip('"').strip("'")
                
                if current_section == 'ocr':
                    if hasattr(self.config.ocr, key):
                        setattr(self.config.ocr, key, self._parse_value(value))
                elif current_section == 'inpaint':
                    if hasattr(self.config.inpaint, key):
                        setattr(self.config.inpaint, key, self._parse_value(value))
                elif current_section == 'ppt':
                    if hasattr(self.config.ppt, key):
                        setattr(self.config.ppt, key, self._parse_value(value))
                else:
                    if hasattr(self.config, key):
                        setattr(self.config, key, self._parse_value(value))
    
    def _apply_dict(self, data: dict):
        """应用字典配置"""
        for key, value in data.items():
            if key == 'ocr' and isinstance(value, dict):
                for ocr_key, ocr_value in value.items():
                    if hasattr(self.config.ocr, ocr_key):
                        setattr(self.config.ocr, ocr_key, self._parse_value(ocr_value))
            elif key == 'inpaint' and isinstance(value, dict):
                for inpaint_key, inpaint_value in value.items():
                    if hasattr(self.config.inpaint, inpaint_key):
                        setattr(self.config.inpaint, inpaint_key, self._parse_value(inpaint_value))
            elif key == 'ppt' and isinstance(value, dict):
                for ppt_key, ppt_value in value.items():
                    if hasattr(self.config.ppt, ppt_key):
                        setattr(self.config.ppt, ppt_key, self._parse_value(ppt_value))
            elif hasattr(self.config, key):
                setattr(self.config, key, self._parse_value(value))
    
    def _load_from_env(self):
        """从环境变量加载"""
        env_mappings = {
            'SD_API_URL': ('inpaint.sd_api_url',),
            'PPT_DEFAULT_FONT': ('ppt.default_font',),
            'OCR_LANG': ('ocr.lang',),
            'OCR_USE_GPU': ('ocr.use_gpu',),
        }
        
        for env_key, config_paths in env_mappings.items():
            value = os.environ.get(env_key)
            if value:
                for path in config_paths:
                    parts = path.split('.')
                    obj = self.config
                    for part in parts[:-1]:
                        obj = getattr(obj, part)
                    setattr(obj, parts[-1], self._parse_value(value))
    
    def _parse_value(self, value: str):
        """解析配置值"""
        # 布尔值
        if value.lower() in ('true', 'yes', '1'):
            return True
        if value.lower() in ('false', 'no', '0'):
            return False
        
        # 数字
        try:
            if '.' in value:
                return float(value)
            return int(value)
        except ValueError:
            return value
    
    def update_from_args(self, args):
        """从命令行参数更新配置"""
        arg_mappings = {
            'input_path': 'input_path',
            'output_path': 'output_path',
            'sd_api_url': ('inpaint', 'sd_api_url'),
            'ppt_template': 'ppt_template',
            'language': 'language',
        }
        
        for arg_key, config_path in arg_mappings.items():
            value = getattr(args, arg_key, None)
            if value is not None:
                if isinstance(config_path, tuple):
                    obj = self.config
                    for part in config_path[:-1]:
                        obj = getattr(obj, part)
                    setattr(obj, config_path[-1], value)
                else:
                    setattr(self.config, config_path, value)
    
    def save(self, path: str):
        """保存配置到文件"""
        data = {
            'input_path': self.config.input_path,
            'output_path': self.config.output_path,
            'language': self.config.language,
            'ocr': {
                'use_angle_cls': self.config.ocr.use_angle_cls,
                'lang': self.config.ocr.lang,
                'use_gpu': self.config.ocr.use_gpu,
            },
            'inpaint': {
                'sd_api_url': self.config.inpaint.sd_api_url,
                'prompt': self.config.inpaint.prompt,
                'negative_prompt': self.config.inpaint.negative_prompt,
                'steps': self.config.inpaint.steps,
                'cfg_scale': self.config.inpaint.cfg_scale,
                'denoising_strength': self.config.inpaint.denoising_strength,
            },
            'ppt': {
                'slide_width': self.config.ppt.slide_width,
                'slide_height': self.config.ppt.slide_height,
                'default_font': self.config.ppt.default_font,
                'default_font_size': self.config.ppt.default_font_size,
            }
        }
        
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
