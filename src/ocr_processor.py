# -*- coding: utf-8 -*-
"""
OCR 处理器模块
使用 PaddleOCR 识别图片文字及其格式信息
"""

import io
import os
from dataclasses import dataclass, field
from typing import List, Tuple, Optional, Dict, Any
from pathlib import Path
from PIL import Image
import numpy as np

from .config import OCRConfig


@dataclass
class TextFormat:
    """文字格式信息（用于 PPT 复现）"""
    # 基础格式
    text: str = ""
    font_size: int = 18  # 估算的字体大小（磅）
    font_name: str = "微软雅黑"
    font_color_rgb: Tuple[int, int, int] = (0, 0, 0)  # RGB 元组
    font_color_hex: str = "000000"  # 十六进制颜色
    
    # 样式
    bold: bool = False
    italic: bool = False
    underline: bool = False
    
    # 段落格式
    alignment: str = "left"  # left, center, right, justify
    line_spacing: float = 1.5  # 行间距倍数
    bullet: Optional[str] = None  # 项目符号，如 "•", "1.", "a)"
    
    # 位置信息
    bbox: Tuple[int, int, int, int] = (0, 0, 0, 0)  # [x1, y1, x2, y2]
    center_x: float = 0.0  # 相对于图片宽度的比例 0-1
    center_y: float = 0.0  # 相对于图片高度的比例 0-1
    
    # 识别置信度
    confidence: float = 0.0
    
    # 文字类型
    text_type: str = "body"  # title, subtitle, body, bullet, footer, header


@dataclass
class OCRResult:
    """OCR 识别结果"""
    image_path: str
    image_width: int
    image_height: int
    texts: List[TextFormat] = field(default_factory=list)
    raw_results: List[Dict[str, Any]] = field(default_factory=list)  # 原始识别结果
    page_layout: Dict[str, Any] = field(default_factory=dict)  # 页面布局分析
    
    @property
    def title_texts(self) -> List[TextFormat]:
        """获取标题类文字"""
        return [t for t in self.texts if t.text_type == "title"]
    
    @property
    def body_texts(self) -> List[TextFormat]:
        """获取正文类文字"""
        return [t for t in self.texts if t.text_type in ("body", "bullet")]
    
    @property
    def all_text(self) -> str:
        """获取所有文字内容"""
        return " ".join([t.text for t in self.texts])


class ImageAnalyzer:
    """图片分析器 - 分析文字区域颜色和样式"""
    
    def __init__(self):
        self.color_cache = {}
    
    def extract_region_color(self, image: Image.Image, bbox: Tuple[int, int, int, int]) -> Tuple[int, int, int]:
        """提取文字区域的平均颜色"""
        x1, y1, x2, y2 = bbox
        cropped = image.crop((x1, y1, x2, y2))
        
        # 转换为 numpy 数组
        img_array = np.array(cropped)
        
        # 计算平均颜色（排除白色背景）
        # 转换为 RGB
        if len(img_array.shape) == 3:
            # 计算非白色像素的平均颜色
            mask = np.all(img_array[:, :, :3] < 240, axis=2)
            if np.any(mask):
                avg_color = img_array[:, :, :3][mask].mean(axis=0)
                return tuple(int(c) for c in avg_color[::-1])  # BGR -> RGB
            else:
                # 默认返回黑色
                return (0, 0, 0)
        return (0, 0, 0)
    
    def estimate_font_size(self, bbox: Tuple[int, int, int, int], image_height: int) -> int:
        """根据文字区域高度估算字体大小
        
        估算逻辑：
        - 标准 PPT 字号范围：12-72 磅
        - 假设图片高度对应 5.625 英寸（标准 PPT 高度）
        - 1 英寸 = 72 磅
        """
        x1, y1, x2, y2 = bbox
        text_height = y2 - y1
        
        # 假设标准 PPT 高度为 5.625 英寸，72 磅/英寸
        # 图片高度对应的字体大小
        estimated_size = int((text_height / image_height) * 72)
        
        # 限制在合理范围内
        return max(12, min(72, estimated_size))
    
    def is_likely_bold(self, image: Image.Image, bbox: Tuple[int, int, int, int]) -> bool:
        """通过分析文字笔画粗细估算是否粗体"""
        x1, y1, x2, y2 = bbox
        width = x2 - x1
        height = y2 - y1
        
        if width == 0 or height == 0:
            return False
        
        cropped = image.crop((x1, y1, x2, y2))
        img_array = np.array(cropped)
        
        # 分析边缘密度
        gray = np.mean(img_array[:, :, :3], axis=2) if len(img_array.shape) == 3 else img_array
        
        # 计算梯度（边缘）
        grad_x = np.abs(np.diff(gray, axis=1))
        grad_y = np.abs(np.diff(gray, axis=0))
        
        # 平均边缘强度
        avg_edge = (np.mean(grad_x) * width + np.mean(grad_y) * height) / (width + height)
        
        # 边缘强度大于阈值认为可能是粗体
        return avg_edge > 30
    
    def detect_bullet(self, text: str, bbox: Tuple[int, int, int, int], 
                      prev_bbox: Optional[Tuple[int, int, int, int]] = None) -> Optional[str]:
        """检测是否为项目符号"""
        # 常见项目符号
        bullets = ['•', '●', '○', '■', '□', '▶', '►', '✓', '✔', '✗', '✿', '★', '☆']
        
        # 检查开头字符
        for bullet in bullets:
            if text.startswith(bullet):
                return bullet
        
        # 检查数字编号
        import re
        if re.match(r'^\d+[\.\)]\s', text):
            match = re.match(r'^(\d+)[\.\)]\s', text)
            return f"{match.group(1)}."
        
        if re.match(r'^[a-z][\.\)]\s', text, re.IGNORECASE):
            match = re.match(r'^([a-z])[\.\)]\s', text, re.IGNORECASE)
            return f"{match.group(1)}."
        
        # 检查与上一行的缩进关系
        if prev_bbox:
            prev_x1 = prev_bbox[0]
            curr_x1 = bbox[0]
            # 如果当前位置比上一行更靠右，可能是列表
            if curr_x1 > prev_x1 + 10:
                return "•"
        
        return None


class OCRProcessor:
    """OCR 处理器主类"""
    
    def __init__(self, config: Optional[OCRConfig] = None):
        self.config = config or OCRConfig()
        self.analyzer = ImageAnalyzer()
        self._ocr_engine = None
    
    def _init_engine(self):
        """初始化 PaddleOCR 引擎"""
        if self._ocr_engine is None:
            from paddleocr import PaddleOCR
            self._ocr_engine = PaddleOCR(
                use_angle_cls=self.config.use_angle_cls,
                lang=self.config.lang,
                use_gpu=self.config.use_gpu,
                show_log=self.config.show_log,
                det_model_dir=self.config.det_model_dir,
                rec_model_dir=self.config.rec_model_dir,
                cls_model_dir=self.config.cls_model_dir,
            )
        return self._ocr_engine
    
    def process_image(self, image_path: str) -> OCRResult:
        """处理单张图片，识别文字和格式"""
        # 加载图片
        image = Image.open(image_path)
        image_width, image_height = image.size
        
        # 初始化引擎并识别
        ocr = self._init_engine()
        raw_results = ocr.ocr(image_path, cls=True)
        
        # 解析结果
        texts = []
        sorted_results = self._sort_by_layout(raw_results)
        
        prev_bbox = None
        for item in sorted_results:
            bbox = item['bbox']
            text = item['text']
            confidence = item['confidence']
            
            if not text.strip():
                continue
            
            # 创建格式对象
            fmt = TextFormat(
                text=text,
                bbox=bbox,
                confidence=confidence,
                center_x=(bbox[0] + bbox[2]) / 2 / image_width,
                center_y=(bbox[1] + bbox[3]) / 2 / image_height,
            )
            
            # 分析颜色
            fmt.font_color_rgb = self.analyzer.extract_region_color(image, bbox)
            fmt.font_color_hex = '{:02X}{:02X}{:02X}'.format(*fmt.font_color_rgb[::-1])
            
            # 估算字体大小
            fmt.font_size = self.analyzer.estimate_font_size(bbox, image_height)
            
            # 检测粗体
            fmt.bold = self.analyzer.is_likely_bold(image, bbox)
            
            # 检测项目符号
            fmt.bullet = self.analyzer.detect_bullet(text, bbox, prev_bbox)
            if fmt.bullet:
                # 移除项目符号
                fmt.text = self._remove_bullet(text, fmt.bullet)
                fmt.text_type = "bullet"
            
            # 判断文字类型
            fmt.text_type = self._classify_text_type(fmt, image_height)
            
            # 分析对齐方式
            fmt.alignment = self._detect_alignment(bbox, image_width)
            
            texts.append(fmt)
            prev_bbox = bbox
        
        # 页面布局分析
        page_layout = self._analyze_page_layout(texts, image_width, image_height)
        
        return OCRResult(
            image_path=image_path,
            image_width=image_width,
            image_height=image_height,
            texts=texts,
            raw_results=raw_results[0] if raw_results else [],
            page_layout=page_layout,
        )
    
    def _sort_by_layout(self, raw_results) -> List[Dict[str, Any]]:
        """按阅读顺序排序识别结果"""
        if not raw_results or not raw_results[0]:
            return []
        
        items = []
        for line in raw_results[0]:
            bbox = [int(x) for x in line[0]]
            text = line[1][0]
            confidence = float(line[1][1])
            items.append({
                'bbox': tuple(bbox),
                'text': text,
                'confidence': confidence,
            })
        
        # 按 Y 坐标分组，然后按 X 坐标排序
        # 简单实现：先按 Y 排序，同一行按 X 排序
        items.sort(key=lambda x: (x['bbox'][1], x['bbox'][0]))
        
        return items
    
    def _classify_text_type(self, fmt: TextFormat, image_height: int) -> str:
        """分类文字类型"""
        y_ratio = fmt.center_y
        font_size = fmt.font_size
        bold = fmt.bold
        
        # 顶部区域（标题区）
        if y_ratio < 0.15:
            if font_size >= 28 or bold:
                return "title"
            return "subtitle"
        
        # 底部区域（页脚）
        if y_ratio > 0.9:
            return "footer"
        
        # 大字体或粗体可能是副标题
        if font_size >= 24 and y_ratio < 0.3:
            return "subtitle"
        
        return "body"
    
    def _detect_alignment(self, bbox: Tuple[int, int, int, int], image_width: int) -> str:
        """检测文字对齐方式"""
        x1, _, x2, _ = bbox
        width = x2 - x1
        
        # 居中判断
        center = (x1 + x2) / 2
        if abs(center - image_width / 2) < image_width * 0.1:
            return "center"
        
        # 右对齐判断
        if x2 > image_width * 0.85:
            return "right"
        
        # 左对齐
        return "left"
    
    def _remove_bullet(self, text: str, bullet: str) -> str:
        """移除文字开头的项目符号"""
        import re
        
        # 处理数字/字母编号
        if re.match(r'^\d+[\.\)]\s', text):
            return re.sub(r'^\d+[\.\)]\s', '', text)
        if re.match(r'^[a-z][\.\)]\s', text, re.IGNORECASE):
            return re.sub(r'^[a-z][\.\)]\s', '', text, flags=re.IGNORECASE)
        
        # 处理符号
        if text.startswith(bullet):
            return text[len(bullet):].lstrip()
        
        return text
    
    def _analyze_page_layout(self, texts: List[TextFormat], 
                            image_width: int, image_height: int) -> Dict[str, Any]:
        """分析页面布局"""
        if not texts:
            return {}
        
        # 统计各区域文字
        title_count = sum(1 for t in texts if t.text_type == "title")
        body_count = len([t for t in texts if t.text_type in ("body", "bullet")])
        
        # 分析列布局
        x_positions = [t.bbox[0] for t in texts]
        columns = 1
        if x_positions:
            x_min, x_max = min(x_positions), max(x_positions)
            if x_max - x_min > image_width * 0.5:
                columns = 2
        
        return {
            'columns': columns,
            'has_title': title_count > 0,
            'has_bullets': any(t.bullet for t in texts),
            'body_text_count': body_count,
        }
    
    def get_mask_regions(self, result: OCRResult) -> List[Tuple[int, int, int, int]]:
        """获取所有文字区域（用于图像修复的掩码）"""
        return [text.bbox for text in result.texts]
    
    def close(self):
        """关闭 OCR 引擎，释放资源"""
        if self._ocr_engine:
            self._ocr_engine = None
