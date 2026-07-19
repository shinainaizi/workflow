#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
活动信息提取模块 - 从策划方案、现场照片、回忆记录中提取结构化信息
"""

import os
import re
from pathlib import Path
from typing import Dict, Optional, List, Any
import docx
from llm_client import LLMClient

class ActivityExtractor:
    """活动信息提取器"""
    
    def __init__(self, config: Dict[str, Any]):
        self.config = config
        self.llm = LLMClient(config)
    
    def extract(self, scheme_path: str, photos_dir: Optional[str] = None, notes_path: Optional[str] = None) -> Dict[str, Any]:
        """
        从各种来源提取活动信息
        
        Args:
            scheme_path: 策划方案文件路径
            photos_dir: 照片目录
            notes_path: 回忆记录文件路径
            
        Returns:
            结构化活动数据
        """
        # 1. 从方案提取基础信息
        scheme_data = self._extract_from_scheme(scheme_path)
        
        # 2. 从回忆记录提取补充信息
        notes_data = {}
        if notes_path and os.path.exists(notes_path):
            notes_data = self._extract_from_notes(notes_path)
        
        # 3. 从照片提取现场信息（如果启用）
        photo_data = {}
        if photos_dir and os.path.exists(photos_dir):
            photo_data = self._extract_from_photos(photos_dir)
        
        # 4. 合并信息
        merged = self._merge_data(scheme_data, notes_data, photo_data)
        
        # 5. 使用AI生成活动描述
        merged['activity_content'] = self._generate_content_description(merged)
        
        return merged
    
    def _extract_from_scheme(self, scheme_path: str) -> Dict[str, Any]:
        """从策划方案提取结构化信息"""
        # 读取Word文档
        doc = docx.Document(scheme_path)
        full_text = '\n'.join([p.text for p in doc.paragraphs])
        
        # 使用AI提取关键信息
        system_prompt = "你是一位社区活动策划专家，擅长从活动方案中提取结构化信息。请严格按照JSON格式输出。"
        
        prompt = f"""
请从以下活动方案中提取关键信息，以JSON格式返回：

活动方案内容：
{full_text}

请提取以下字段（如方案中未提及则留空）：
{{
    "activity_theme": "活动主题",
    "activity_date": "活动日期",
    "activity_time": "活动时间",
    "activity_location": "活动地点",
    "participants": "参与对象",
    "expected_participants": "预期参与人数",
    "activity_goal": "活动目的",
    "activity_content": "活动内容概述",
    "activity_flow": ["环节1", "环节2"],
    "staff": "工作人员",
    "notes": "注意事项"
}}

注意：
- 日期格式为 YYYY-MM-DD
- 时间格式为 HH:MM-HH:MM
- 只返回JSON，不要其他说明文字
"""
        
        result = self.llm.extract_json(prompt, system_prompt)
        
        if not result:
            # 降级方案：手动提取
            result = self._manual_extract(full_text)
        
        return result
    
    def _manual_extract(self, text: str) -> Dict[str, Any]:
        """手动提取方案信息（降级方案）"""
        data = {
            'activity_theme': '',
            'activity_date': '',
            'activity_time': '',
            'activity_location': '',
            'participants': '',
            'expected_participants': '',
            'activity_goal': '',
            'activity_content': '',
            'activity_flow': [],
            'staff': '',
            'notes': ''
        }
        
        # 正则匹配
        # 活动主题
        theme_match = re.search(r'(?:活动主题|主题)[：:]\s*(.+)', text)
        if theme_match:
            data['activity_theme'] = theme_match.group(1).strip()
        
        # 活动日期
        date_match = re.search(r'(?:活动时间|时间)[：:]\s*(\d{1,2}[月/]\d{1,2}[日/])', text)
        if date_match:
            data['activity_date'] = date_match.group(1).strip()
        
        # 时间
        time_match = re.search(r'(\d{1,2}:\d{2})\s*[-~]\s*(\d{1,2}:\d{2})', text)
        if time_match:
            data['activity_time'] = f"{time_match.group(1)}-{time_match.group(2)}"
        
        # 地点
        location_match = re.search(r'(?:活动地点|地点)[：:]\s*(.+)', text)
        if location_match:
            data['activity_location'] = location_match.group(1).strip()
        
        # 参与对象
        participants_match = re.search(r'(?:参与对象|对象)[：:]\s*(.+)', text)
        if participants_match:
            data['participants'] = participants_match.group(1).strip()
        
        # 活动目的
        goal_match = re.search(r'(?:活动目的|目的)[：:]\s*(.+)', text)
        if goal_match:
            data['activity_goal'] = goal_match.group(1).strip()
        
        return data
    
    def _extract_from_notes(self, notes_path: str) -> Dict[str, Any]:
        """从回忆记录提取信息"""
        with open(notes_path, 'r', encoding='utf-8') as f:
            notes_text = f.read()
        
        system_prompt = "你是一位社区活动记录员，擅长从活动回忆记录中提取关键信息。请严格按照JSON格式输出。"
        
        prompt = f"""
请从以下活动回忆记录中提取关键信息，以JSON格式返回：

回忆记录：
{notes_text}

请提取以下字段：
{{
    "actual_participants": "实际参与人数",
    "participant_feedback": "参与者反馈",
    "highlights": "活动亮点",
    "problems": "存在的问题",
    "improvements": "改进建议",
    "actual_date": "实际活动日期",
    "actual_time": "实际活动时间",
    "scene_description": "现场情况描述"
}}

注意：只返回JSON，不要其他说明文字
"""
        
        result = self.llm.extract_json(prompt, system_prompt)
        
        if not result:
            return {}
        
        return result
    
    def _extract_from_photos(self, photos_dir: str) -> Dict[str, Any]:
        """从照片提取现场信息"""
        photos_dir = Path(photos_dir)
        image_files = []
        
        for ext in ['*.jpg', '*.jpeg', '*.png']:
            image_files.extend(photos_dir.glob(ext))
        
        if not image_files:
            return {}
        
        # 如果启用了照片理解，使用多模态模型
        if self.config.get('qianfan', {}).get('enable_photo_understanding', False):
            image_paths = [str(f) for f in image_files[:4]]  # 最多分析4张照片
            
            prompt = """
请分析这些活动现场照片，提取以下信息：
- 现场参与人数（估算）
- 现场布置情况
- 活动氛围
- 是否有特殊事件发生

以JSON格式返回：
{
    "estimated_participants": "估算参与人数",
    "venue_setup": "场地布置描述",
    "atmosphere": "活动氛围",
    "special_events": "特殊事件",
    "photos_count": "照片数量"
}
"""
            
            result = self.llm.chat_with_images(prompt, image_paths)
            
            # 尝试解析结果
            try:
                import json
                return json.loads(result)
            except:
                return {"photo_description": result, "photos_count": len(image_files)}
        
        else:
            # 未启用照片理解，返回基本信息
            return {
                "photos_count": len(image_files),
                "photos_description": f"现场照片共{len(image_files)}张，请在回忆记录中描述照片内容"
            }
    
    def _merge_data(self, scheme_data: Dict, notes_data: Dict, photo_data: Dict) -> Dict[str, Any]:
        """合并所有来源的数据"""
        merged = {
            # 基础信息（以方案为准）
            'activity_theme': scheme_data.get('activity_theme', ''),
            'activity_date': notes_data.get('actual_date', scheme_data.get('activity_date', '')),
            'activity_time': notes_data.get('actual_time', scheme_data.get('activity_time', '')),
            'activity_location': scheme_data.get('activity_location', ''),
            'participants': scheme_data.get('participants', '社区居民'),
            'target_group': scheme_data.get('participants', ''),
            
            # 人数信息
            'expected_participants': scheme_data.get('expected_participants', ''),
            'actual_participants': notes_data.get('actual_participants', ''),
            'estimated_participants': photo_data.get('estimated_participants', ''),
            
            # 活动内容
            'activity_goal': scheme_data.get('activity_goal', ''),
            'activity_flow': scheme_data.get('activity_flow', []),
            'highlights': notes_data.get('highlights', ''),
            'problems': notes_data.get('problems', ''),
            'improvements': notes_data.get('improvements', ''),
            
            # 反馈
            'participant_feedback': notes_data.get('participant_feedback', ''),
            'scene_description': notes_data.get('scene_description', ''),
            
            # 工作人员
            'staff': scheme_data.get('staff', '社区工作人员'),
            'notes': scheme_data.get('notes', ''),
            
            # 照片信息
            'photos_count': photo_data.get('photos_count', 0),
            'photos_description': photo_data.get('photos_description', ''),
        }
        
        # 补充社区信息
        community = self.config.get('community', {})
        if not merged['activity_location']:
            merged['activity_location'] = community.get('venue', '')
        if not merged['participants']:
            merged['participants'] = community.get('default_target', '社区居民')
        
        return merged
    
    def _generate_content_description(self, data: Dict[str, Any]) -> str:
        """使用AI生成活动内容描述"""
        system_prompt = "你是一位社区活动记录员，擅长撰写简洁、专业的活动记录。"
        
        # 构建上下文
        context = f"""
活动主题：{data.get('activity_theme', '')}
活动日期：{data.get('activity_date', '')} {data.get('activity_time', '')}
活动地点：{data.get('activity_location', '')}
参与对象：{data.get('participants', '')}
实际参与人数：{data.get('actual_participants', data.get('estimated_participants', '未知'))}
活动目的：{data.get('activity_goal', '')}
活动流程：{'、'.join(data.get('activity_flow', []))}
活动亮点：{data.get('highlights', '')}
存在问题：{data.get('problems', '')}
改进建议：{data.get('improvements', '')}
现场描述：{data.get('scene_description', '')}
参与者反馈：{data.get('participant_feedback', '')}
"""
        
        prompt = f"""
请根据以下活动信息，撰写一段简洁的活动内容描述（200-400字），用于社区活动台账。

要求：
1. 包含活动时间、地点、主题、参与人数
2. 简述活动流程和主要内容
3. 提及活动效果和参与者反馈
4. 适当提及后续计划
5. 语言简洁专业，适合社区台账记录
6. 不要出现"我"字，用第三人称

{context}

请直接输出描述文字，不要添加标题。
"""
        
        result = self.llm.chat(prompt, system_prompt)
        
        return result.strip()
