# -*- coding: utf-8 -*-
"""
图像修复模块
使用 Stable Diffusion WebUI API 进行图像修复，去除文字区域
"""

import io
import base64
import json
import time
from dataclasses import dataclass
from typing import List, Tuple, Optional, Dict, Any
from pathlib import Path
import requests
from PIL import Image, ImageDraw, ImageFilter

from .config import InpaintConfig


@dataclass
class InpaintResult:
    """图像修复结果"""
    cleaned_image: Image.Image
    original_size: Tuple[int, int]
    mask_used: Optional[Image.Image] = None
    success: bool = True
    error_message: Optional[str] = None


class MaskGenerator:
    """掩码生成器 - 生成文字区域的掩码"""
    
    def __init__(self, mask_expand: int = 2, mask_blur: int = 3):
        """
        Args:
            mask_expand: 掩码扩展像素，让掩码略大于文字区域
            mask_blur: 掩码边缘模糊半径
        """
        self.mask_expand = mask_expand
        self.mask_blur = mask_blur
    
    def generate_mask(self, image: Image.Image, 
                      regions: List[Tuple[int, int, int, int]]) -> Image.Image:
        """生成掩码图像
        
        Args:
            image: 原始图片
            regions: 文字区域列表 [(x1, y1, x2, y2), ...]
        
        Returns:
            掩码图像（白色=保留区域，黑色=修复区域）
        """
        width, height = image.size
        
        # 创建掩码（全白=保留，黑色=需要修复）
        mask = Image.new('L', (width, height), 255)
        draw = ImageDraw.Draw(mask)
        
        for bbox in regions:
            x1, y1, x2, y2 = bbox
            
            # 扩展掩码区域
            x1 = max(0, x1 - self.mask_expand)
            y1 = max(0, y1 - self.mask_expand)
            x2 = min(width, x2 + self.mask_expand)
            y2 = min(height, y2 + self.mask_expand)
            
            # 绘制矩形
            draw.rectangle([x1, y1, x2, y2], fill=0)
        
        # 模糊边缘以获得更自然的修复效果
        if self.mask_blur > 0:
            mask = mask.filter(ImageFilter.GaussianBlur(self.mask_blur))
        
        return mask
    
    def generate_mask_from_regions(self, image: Image.Image,
                                     regions: List[Tuple[int, int, int, int]]) -> Image.Image:
        """从区域列表生成掩码"""
        return self.generate_mask(image, regions)


class ImageInpainter:
    """图像修复主类 - 调用 Stable Diffusion WebUI API"""
    
    def __init__(self, config: Optional[InpaintConfig] = None):
        self.config = config or InpaintConfig()
        self.mask_generator = MaskGenerator()
        self._session = requests.Session()
    
    def _check_api_available(self) -> Tuple[bool, str]:
        """检查 SD WebUI API 是否可用"""
        try:
            response = self._session.get(
                f"{self.config.sd_api_url}/sdapi/v1/ping",
                timeout=5
            )
            if response.status_code == 200:
                return True, "OK"
            return False, f"API 返回状态码: {response.status_code}"
        except requests.exceptions.ConnectionError:
            return False, "无法连接到 SD WebUI API，请确保 SD WebUI 已启动并启用 --api 参数"
        except Exception as e:
            return False, f"检查 API 时出错: {str(e)}"
    
    def _image_to_base64(self, image: Image.Image) -> str:
        """将 PIL Image 转换为 base64 字符串"""
        buffer = io.BytesIO()
        image.save(buffer, format='PNG')
        return base64.b64encode(buffer.getvalue()).decode('utf-8')
    
    def _base64_to_image(self, base64_str: str) -> Image.Image:
        """将 base64 字符串转换为 PIL Image"""
        image_data = base64.b64decode(base64_str)
        return Image.open(io.BytesIO(image_data))
    
    def inpaint(self, image: Image.Image, 
                mask_regions: List[Tuple[int, int, int, int]],
                custom_prompt: Optional[str] = None,
                custom_negative_prompt: Optional[str] = None) -> InpaintResult:
        """执行图像修复
        
        Args:
            image: 原始图片
            mask_regions: 需要修复的区域列表
            custom_prompt: 自定义提示词（可选）
            custom_negative_prompt: 自定义负面提示词（可选）
        
        Returns:
            InpaintResult: 修复结果
        """
        # 检查 API 可用性
        available, message = self._check_api_available()
        if not available:
            return InpaintResult(
                cleaned_image=image,
                original_size=image.size,
                success=False,
                error_message=message,
            )
        
        # 生成掩码
        mask = self.mask_generator.generate_mask_from_regions(image, mask_regions)
        
        # 准备 API 请求
        prompt = custom_prompt or self.config.prompt
        negative_prompt = custom_negative_prompt or self.config.negative_prompt
        
        payload = {
            "init_images": [self._image_to_base64(image)],
            "mask": self._image_to_base64(mask),
            "prompt": prompt,
            "negative_prompt": negative_prompt,
            "steps": self.config.steps,
            "cfg_scale": self.config.cfg_scale,
            "denoising_strength": self.config.denoising_strength,
            "inpainting_fill": self.config.inpainting_fill,
            "mask_blur": self.config.mask_blur,
            "inpaint_full_res": False,
            "inpaint_full_res_padding": 0,
            "inpainting_mask_invert": 0,  # 0=白色区域修复, 1=黑色区域修复
        }
        
        try:
            # 发送请求
            response = self._session.post(
                f"{self.config.sd_api_url}/sdapi/v1/img2img",
                json=payload,
                timeout=120,
            )
            
            if response.status_code != 200:
                return InpaintResult(
                    cleaned_image=image,
                    original_size=image.size,
                    mask_used=mask,
                    success=False,
                    error_message=f"API 返回错误: {response.status_code}",
                )
            
            # 解析响应
            result_data = response.json()
            if 'images' in result_data and len(result_data['images']) > 0:
                cleaned_image = self._base64_to_image(result_data['images'][0])
                
                # 确保尺寸一致
                if cleaned_image.size != image.size:
                    cleaned_image = cleaned_image.resize(image.size, Image.LANCZOS)
                
                return InpaintResult(
                    cleaned_image=cleaned_image,
                    original_size=image.size,
                    mask_used=mask,
                    success=True,
                )
            else:
                return InpaintResult(
                    cleaned_image=image,
                    original_size=image.size,
                    mask_used=mask,
                    success=False,
                    error_message="API 响应中未找到图像数据",
                )
        
        except requests.exceptions.Timeout:
            return InpaintResult(
                cleaned_image=image,
                original_size=image.size,
                mask_used=mask,
                success=False,
                error_message="API 请求超时，请尝试降低 steps 参数或检查 SD WebUI 状态",
            )
        except Exception as e:
            return InpaintResult(
                cleaned_image=image,
                original_size=image.size,
                mask_used=mask,
                success=False,
                error_message=f"图像修复时出错: {str(e)}",
            )
    
    def inpaint_with_fallback(self, image: Image.Image,
                              mask_regions: List[Tuple[int, int, int, int]],
                              fallback_method: str = "blur") -> InpaintResult:
        """使用备用方法的图像修复
        
        当 SD API 不可用时，使用本地方法进行简单修复
        
        Args:
            image: 原始图片
            mask_regions: 需要修复的区域
            fallback_method: 备用方法 "blur" | "surround" | "solid"
        """
        # 生成掩码
        mask = self.mask_generator.generate_mask_from_regions(image, mask_regions)
        
        if fallback_method == "blur":
            # 方法1：模糊修复
            result = image.copy()
            for bbox in mask_regions:
                x1, y1, x2, y2 = bbox
                region = result.crop((x1, y1, x2, y2))
                blurred = region.filter(ImageFilter.GaussianBlur(radius=5))
                result.paste(blurred, (x1, y1))
        
        elif fallback_method == "surround":
            # 方法2：用周围像素填充
            result = image.copy()
            for bbox in mask_regions:
                x1, y1, x2, y2 = bbox
                # 扩展区域以获取更多上下文
                expand = 20
                x1_src = max(0, x1 - expand)
                y1_src = max(0, y1 - expand)
                x2_src = min(image.width, x2 + expand)
                y2_src = min(image.height, y2 + expand)
                
                # 使用区域上方的像素（通常更相似）
                src_region = result.crop((x1_src, y1_src, x2_src, y2_src))
                src_region = src_region.resize((x2 - x1, y2 - y1), Image.LANCZOS)
                result.paste(src_region, (x1, y1))
        
        else:  # solid
            # 方法3：使用纯色填充
            result = image.copy()
            # 分析周围颜色
            avg_color = self._get_average_surrounding_color(image, mask_regions)
            draw = ImageDraw.Draw(result)
            for bbox in mask_regions:
                draw.rectangle(bbox, fill=avg_color)
        
        return InpaintResult(
            cleaned_image=result,
            original_size=image.size,
            mask_used=mask,
            success=True,
            error_message=f"使用备用方法 ({fallback_method}) 进行修复",
        )
    
    def _get_average_surrounding_color(self, image: Image.Image,
                                       region: Tuple[int, int, int, int]) -> Tuple[int, int, int]:
        """获取区域周围像素的平均颜色"""
        x1, y1, x2, y2 = region
        expand = 10
        
        # 获取区域周围的一圈像素
        borders = []
        if y1 > expand:
            borders.extend(image.crop((x1, y1 - expand, x2, y1)).getdata())
        if y2 < image.height - expand:
            borders.extend(image.crop((x1, y2, x2, y2 + expand)).getdata())
        if x1 > expand:
            borders.extend(image.crop((x1 - expand, y1, x1, y2)).getdata())
        if x2 < image.width - expand:
            borders.extend(image.crop((x2, y1, x2 + expand, y2)).getdata())
        
        if borders:
            import numpy as np
            borders_array = np.array(borders)[:, :3]  # 只取 RGB
            avg = borders_array.mean(axis=0).astype(int)
            return tuple(avg.tolist())
        
        return (255, 255, 255)  # 默认白色
    
    def process_image_file(self, image_path: str,
                          mask_regions: List[Tuple[int, int, int, int]],
                          output_path: Optional[str] = None,
                          use_fallback: bool = True) -> InpaintResult:
        """处理图片文件
        
        Args:
            image_path: 输入图片路径
            mask_regions: 需要修复的区域
            output_path: 可选的输出路径
            use_fallback: API 不可用时是否使用备用方法
        
        Returns:
            InpaintResult: 修复结果
        """
        image = Image.open(image_path)
        
        result = self.inpaint(image, mask_regions)
        
        if not result.success and use_fallback:
            # 使用备用方法
            result = self.inpaint_with_fallback(image, mask_regions)
        
        # 保存结果
        if output_path and result.success:
            result.cleaned_image.save(output_path)
        
        return result
    
    def close(self):
        """关闭会话"""
        self._session.close()
