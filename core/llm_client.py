#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
大模型调用客户端 - 统一封装百度千帆API
"""

import os
import json
import base64
from typing import Optional, List, Dict, Any
import qianfan

class LLMClient:
    """大模型客户端，支持文本和视觉理解"""
    
    def __init__(self, config: Dict[str, Any]):
        """
        初始化客户端
        
        Args:
            config: 配置字典，包含qianfan.api_key和qianfan.secret_key
        """
        self.config = config
        self.api_key = config.get('qianfan', {}).get('api_key', '')
        self.secret_key = config.get('qianfan', {}).get('secret_key', '')
        self.text_model = config.get('qianfan', {}).get('text_model', 'ERNIE-Speed')
        self.vision_model = config.get('qianfan', {}).get('vision_model', 'ERNIE-4.0-Turbo-VL')
        self.enable_photo = config.get('qianfan', {}).get('enable_photo_understanding', False)
        
        # 初始化千帆客户端
        if self.api_key and self.api_key != 'your-api-key-here':
            os.environ['QIANFAN_AK'] = self.api_key
            os.environ['QIANFAN_SK'] = self.secret_key
            qianfan.AK(self.api_key)
            qianfan.SK(self.secret_key)
    
    def chat(self, prompt: str, system_prompt: str = "", model: Optional[str] = None) -> str:
        """
        纯文本对话
        
        Args:
            prompt: 用户输入
            system_prompt: 系统提示词
            model: 指定模型，默认使用配置中的text_model
            
        Returns:
            模型生成的文本
        """
        try:
            model_name = model or self.text_model
            
            # 使用千帆SDK
            chat_comp = qianfan.ChatCompletion()
            
            messages = []
            if system_prompt:
                messages.append({"role": "system", "content": system_prompt})
            messages.append({"role": "user", "content": prompt})
            
            resp = chat_comp.do(
                model=model_name,
                messages=messages,
                temperature=0.7,
                top_p=0.8,
            )
            
            return resp["body"]["result"]
            
        except Exception as e:
            print(f"[LLM调用失败] {e}")
            # 降级方案：返回空字符串，让上层处理
            return ""
    
    def chat_with_images(self, prompt: str, image_paths: List[str], system_prompt: str = "") -> str:
        """
        带图片的多模态对话
        
        Args:
            prompt: 文本提示
            image_paths: 图片文件路径列表
            system_prompt: 系统提示词
            
        Returns:
            模型生成的文本
        """
        if not self.enable_photo:
            return "[照片理解已禁用] 请在config.yaml中启用enable_photo_understanding并配置vision_model"
        
        try:
            chat_comp = qianfan.ChatCompletion()
            
            # 构建消息内容
            content = []
            content.append({"type": "text", "text": prompt})
            
            for img_path in image_paths:
                if os.path.exists(img_path):
                    with open(img_path, 'rb') as f:
                        img_base64 = base64.b64encode(f.read()).decode('utf-8')
                    content.append({
                        "type": "image_url",
                        "image_url": {
                            "url": f"data:image/jpeg;base64,{img_base64}"
                        }
                    })
            
            messages = []
            if system_prompt:
                messages.append({"role": "system", "content": system_prompt})
            messages.append({"role": "user", "content": content})
            
            resp = chat_comp.do(
                model=self.vision_model,
                messages=messages,
                temperature=0.7,
            )
            
            return resp["body"]["result"]
            
        except Exception as e:
            print(f"[多模态调用失败] {e}")
            return ""
    
    def extract_json(self, prompt: str, system_prompt: str = "") -> Optional[Dict]:
        """
        调用LLM并解析返回结果为JSON
        
        Args:
            prompt: 提示词
            system_prompt: 系统提示
            
        Returns:
            解析后的字典，失败返回None
        """
        response = self.chat(prompt, system_prompt)
        
        try:
            # 尝试直接解析
            return json.loads(response)
        except json.JSONDecodeError:
            # 尝试从文本中提取JSON块
            try:
                # 查找 ```json ... ``` 或 ``` ... ``` 格式
                import re
                json_match = re.search(r'```(?:json)?\s*(.*?)\s*```', response, re.DOTALL)
                if json_match:
                    return json.loads(json_match.group(1))
                
                # 尝试查找大括号包裹的内容
                json_match = re.search(r'\{.*\}', response, re.DOTALL)
                if json_match:
                    return json.loads(json_match.group(0))
                    
            except json.JSONDecodeError:
                pass
        
        return None
