# -*- coding: utf-8 -*-
"""
PPT 图片修复 Skill - 主入口
将模糊的 PPT 图片转化为清晰可编辑的 PPT 文件
"""

import argparse
import sys
import os
from pathlib import Path
from typing import List, Optional, Tuple
from dataclasses import dataclass

from PIL import Image

from .config import ConfigManager, SkillConfig
from .ocr_processor import OCRProcessor, OCRResult
from .image_inpainter import ImageInpainter, InpaintResult
from .ppt_generator import PPTGeneratorPipeline, PPTGenerationResult


@dataclass
class ProcessingResult:
    """单张图片处理结果"""
    image_path: str
    success: bool
    ocr_result: Optional[OCRResult] = None
    inpaint_result: Optional[InpaintResult] = None
    ppt_result: Optional[PPTGenerationResult] = None
    error_message: Optional[str] = None


class PPTImageRestorer:
    """PPT 图片修复主类"""
    
    def __init__(self, config: SkillConfig):
        self.config = config
        self.ocr_processor = OCRProcessor(config.ocr)
        self.inpainter = ImageInpainter(config.inpaint)
        self.pipeline = PPTGeneratorPipeline(config.ppt)
        self.results: List[ProcessingResult] = []
    
    def process_single_image(self, image_path: str, 
                            output_ppt_path: Optional[str] = None) -> ProcessingResult:
        """处理单张图片
        
        Args:
            image_path: 输入图片路径
            output_ppt_path: 输出 PPT 路径（可选）
        
        Returns:
            ProcessingResult: 处理结果
        """
        result = ProcessingResult(image_path=image_path, success=False)
        
        try:
            if self.config.verbose:
                print(f"[*] 正在处理: {image_path}")
            
            # 步骤 1: OCR 识别
            if self.config.verbose:
                print(f"    [1/3] OCR 识别文字...")
            
            ocr_result = self.ocr_processor.process_image(image_path)
            result.ocr_result = ocr_result
            
            if self.config.verbose:
                print(f"    识别到 {len(ocr_result.texts)} 个文字区域")
            
            # 步骤 2: 图像修复
            if self.config.verbose:
                print(f"    [2/3] 图像修复去除文字...")
            
            mask_regions = self.ocr_processor.get_mask_regions(ocr_result)
            original_image = Image.open(image_path)
            
            inpaint_result = self.inpainter.inpaint(original_image, mask_regions)
            result.inpaint_result = inpaint_result
            
            if not inpaint_result.success:
                if self.config.verbose:
                    print(f"    SD API 不可用，使用备用方法...")
                inpaint_result = self.inpainter.inpaint_with_fallback(
                    original_image, mask_regions
                )
            
            # 保存中间结果
            if self.config.output_intermediate and inpaint_result.success:
                self._save_intermediate(image_path, inpaint_result.cleaned_image)
            
            # 步骤 3: 生成 PPT
            if self.config.verbose:
                print(f"    [3/3] 生成可编辑 PPT...")
            
            if output_ppt_path is None:
                output_ppt_path = self._get_default_output_path(image_path)
            
            ppt_result = self.pipeline.create_from_ocr_and_inpaint(
                image=original_image,
                ocr_result=ocr_result,
                cleaned_image=inpaint_result.cleaned_image,
                output_path=output_ppt_path,
            )
            result.ppt_result = ppt_result
            
            if ppt_result.success:
                result.success = True
                if self.config.verbose:
                    print(f"    完成! 输出: {output_ppt_path}")
            else:
                result.error_message = ppt_result.error_message
                if self.config.verbose:
                    print(f"    PPT 生成失败: {ppt_result.error_message}")
        
        except Exception as e:
            result.error_message = str(e)
            if self.config.verbose:
                print(f"    错误: {e}")
        
        return result
    
    def process_multiple_images(self, input_path: str,
                               output_path: Optional[str] = None) -> List[ProcessingResult]:
        """处理多张图片
        
        Args:
            input_path: 输入路径（文件或文件夹）
            output_path: 输出 PPT 路径（文件夹模式时为目录）
        
        Returns:
            处理结果列表
        """
        image_paths = self._collect_images(input_path)
        
        if not image_paths:
            print(f"[!] 未找到图片文件: {input_path}")
            return []
        
        if self.config.verbose:
            print(f"[*] 找到 {len(image_paths)} 张图片")
        
        # 如果是文件夹模式，为每张图片生成单独的 PPT
        if os.path.isdir(input_path):
            return self._process_as_slideshow(image_paths, output_path)
        
        # 单文件模式
        if len(image_paths) == 1:
            result = self.process_single_image(image_paths[0], output_path)
            return [result]
        
        # 多文件模式：创建包含所有图片的 PPT
        return self._process_as_slideshow(image_paths, output_path)
    
    def _process_as_slideshow(self, image_paths: List[str],
                             output_path: Optional[str] = None) -> List[ProcessingResult]:
        """作为幻灯片处理多张图片"""
        slides_data = []
        
        for image_path in image_paths:
            try:
                # OCR
                ocr_result = self.ocr_processor.process_image(image_path)
                image = Image.open(image_path)
                
                # 图像修复
                mask_regions = self.ocr_processor.get_mask_regions(ocr_result)
                inpaint_result = self.inpainter.inpaint(image, mask_regions)
                
                if not inpaint_result.success:
                    inpaint_result = self.inpainter.inpaint_with_fallback(
                        image, mask_regions
                    )
                
                slides_data.append({
                    'image': inpaint_result.cleaned_image,
                    'ocr_result': ocr_result,
                })
                
                result = ProcessingResult(
                    image_path=image_path,
                    success=True,
                    ocr_result=ocr_result,
                    inpaint_result=inpaint_result,
                )
                self.results.append(result)
                
            except Exception as e:
                result = ProcessingResult(
                    image_path=image_path,
                    success=False,
                    error_message=str(e),
                )
                self.results.append(result)
        
        # 生成统一的 PPT
        if output_path is None:
            # 使用第一张图片的名称
            first_name = Path(image_paths[0]).stem
            output_path = f"{first_name}_slideshow.pptx"
        
        if self.config.verbose:
            print(f"[*] 生成多页 PPT: {output_path}")
        
        ppt_result = self.pipeline.create_multi_slide(slides_data, output_path)
        
        # 更新结果
        for result in self.results:
            result.ppt_result = ppt_result
        
        return self.results
    
    def _collect_images(self, input_path: str) -> List[str]:
        """收集输入路径下的所有图片"""
        supported_formats = {'.png', '.jpg', '.jpeg', '.bmp', '.tiff', '.webp'}
        
        if os.path.isfile(input_path):
            if Path(input_path).suffix.lower() in supported_formats:
                return [input_path]
            return []
        
        if os.path.isdir(input_path):
            images = []
            for ext in supported_formats:
                images.extend(Path(input_path).glob(f"*{ext}"))
                images.extend(Path(input_path).glob(f"*{ext.upper()}"))
            return [str(p) for p in sorted(images)]
        
        return []
    
    def _get_default_output_path(self, image_path: str) -> str:
        """获取默认输出路径"""
        stem = Path(image_path).stem
        return f"{stem}_restored.pptx"
    
    def _save_intermediate(self, image_path: str, image: Image.Image):
        """保存中间结果"""
        Path(self.config.intermediate_dir).mkdir(parents=True, exist_ok=True)
        
        stem = Path(image_path).stem
        output = os.path.join(
            self.config.intermediate_dir,
            f"{stem}_cleaned.png"
        )
        image.save(output)
        
        if self.config.verbose:
            print(f"    保存中间结果: {output}")
    
    def close(self):
        """关闭资源"""
        self.ocr_processor.close()
        self.inpainter.close()
    
    def print_summary(self):
        """打印处理摘要"""
        total = len(self.results)
        success = sum(1 for r in self.results if r.success)
        failed = total - success
        
        print("\n" + "=" * 50)
        print("处理摘要")
        print("=" * 50)
        print(f"总计: {total} 张图片")
        print(f"成功: {success}")
        print(f"失败: {failed}")
        
        if failed > 0:
            print("\n失败详情:")
            for r in self.results:
                if not r.success:
                    print(f"  - {r.image_path}: {r.error_message}")


def parse_args():
    """解析命令行参数"""
    parser = argparse.ArgumentParser(
        description="PPT 图片修复 - 将模糊的 PPT 图片转化为清晰可编辑的 PPT 文件",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  python -m src.main --input slide.png --output result.pptx
  python -m src.main --input ./slides/ --output presentation.pptx
  python -m src.main --input slide.png --sd_api_url http://192.168.1.100:7860
        """
    )
    
    parser.add_argument(
        '--input', '-i',
        required=True,
        help='输入图片路径或文件夹路径'
    )
    
    parser.add_argument(
        '--output', '-o',
        help='输出 PPT 文件路径（默认：与输入同名.pptx）'
    )
    
    parser.add_argument(
        '--sd_api_url',
        default='http://127.0.0.1:7860',
        help='Stable Diffusion WebUI API 地址（默认：http://127.0.0.1:7860）'
    )
    
    parser.add_argument(
        '--config', '-c',
        help='配置文件路径（TOML 或 JSON 格式）'
    )
    
    parser.add_argument(
        '--output_intermediate',
        action='store_true',
        help='保存中间结果（修复后的图片）'
    )
    
    parser.add_argument(
        '--intermediate_dir',
        default='./output_intermediate',
        help='中间结果保存目录（默认：./output_intermediate）'
    )
    
    parser.add_argument(
        '--verbose', '-v',
        action='store_true',
        help='显示详细输出'
    )
    
    parser.add_argument(
        '--quiet', '-q',
        action='store_true',
        help='静默模式（仅显示结果）'
    )
    
    return parser.parse_args()


def main():
    """主函数"""
    args = parse_args()
    
    # 创建配置管理器
    config_manager = ConfigManager(args.config)
    config_manager.update_from_args(args)
    
    # 应用命令行参数
    config = config_manager.config
    config.input_path = args.input
    if args.output:
        config.output_path = args.output
    config.inpaint.sd_api_url = args.sd_api_url
    config.output_intermediate = args.output_intermediate
    config.intermediate_dir = args.intermediate_dir
    config.verbose = not args.quiet
    
    # 创建处理器
    restorer = PPTImageRestorer(config)
    
    try:
        # 执行处理
        results = restorer.process_multiple_images(
            input_path=args.input,
            output_path=args.output,
        )
        
        # 打印摘要
        restorer.print_summary()
        
        # 返回状态码
        if any(r.success for r in results):
            sys.exit(0)
        else:
            sys.exit(1)
    
    finally:
        restorer.close()


if __name__ == '__main__':
    main()
