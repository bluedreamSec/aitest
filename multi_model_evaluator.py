#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
多模型安全评测脚本
读取Excel文件第二列内容，调用配置的大模型，将回答写入结果Excel文件。
支持模型：GPT、DeepSeek、千问、豆包等（通过配置扩展）
"""

import os
import sys
import time
import logging
import pandas as pd
from pathlib import Path
from typing import List, Optional, Dict, Any
from abc import ABC, abstractmethod

# 尝试导入openai，如果未安装则跳过（仅当使用GPT/DeepSeek时需要）
try:
    from openai import OpenAI
    OPENAI_AVAILABLE = True
except ImportError:
    OPENAI_AVAILABLE = False
    print("警告: openai 库未安装，GPT/DeepSeek 模型将不可用。")

# 尝试导入requests，用于非OpenAI API
try:
    import requests
    REQUESTS_AVAILABLE = True
except ImportError:
    REQUESTS_AVAILABLE = False
    print("警告: requests 库未安装，部分模型可能不可用。")

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler('model_evaluator.log', encoding='utf-8')
    ]
)
logger = logging.getLogger(__name__)


class BaseModelClient(ABC):
    """模型客户端基类"""
    
    @abstractmethod
    def chat_completion(self, prompt: str, system_prompt: str = "You are a helpful assistant") -> str:
        """
        调用模型聊天补全
        
        Args:
            prompt: 用户提示
            system_prompt: 系统提示
            
        Returns:
            模型返回的文本内容
        """
        pass


class OpenAIModelClient(BaseModelClient):
    """OpenAI兼容客户端（GPT、DeepSeek等）"""
    
    def __init__(self, api_key: str, base_url: str, model: str):
        """
        初始化OpenAI客户端
        
        Args:
            api_key: API密钥
            base_url: API基础URL
            model: 模型名称
        """
        if not OPENAI_AVAILABLE:
            raise ImportError("openai 库未安装，请运行 pip install openai")
        self.client = OpenAI(api_key=api_key, base_url=base_url)
        self.model = model
    
    def chat_completion(self, prompt: str, system_prompt: str = "You are a helpful assistant") -> str:
        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": prompt},
                ],
                stream=False
            )
            return response.choices[0].message.content.strip()
        except Exception as e:
            logger.error(f"OpenAI API调用失败: {e}")
            raise


class QwenModelClient(BaseModelClient):
    """通义千问客户端（使用DashScope API）"""
    
    def __init__(self, api_key: str, model: str = "qwen-max"):
        """
        初始化千问客户端
        
        Args:
            api_key: DashScope API密钥
            model: 模型名称，默认为 qwen-max
        """
        if not REQUESTS_AVAILABLE:
            raise ImportError("requests 库未安装，请运行 pip install requests")
        self.api_key = api_key
        self.model = model
        self.base_url = "https://dashscope.aliyuncs.com/api/v1/services/aigc/text-generation/generation"
    
    def chat_completion(self, prompt: str, system_prompt: str = "You are a helpful assistant") -> str:
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json"
        }
        data = {
            "model": self.model,
            "input": {
                "messages": [
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": prompt}
                ]
            },
            "parameters": {}
        }
        try:
            response = requests.post(self.base_url, headers=headers, json=data, timeout=30)
            response.raise_for_status()
            result = response.json()
            # 解析响应，具体结构参考DashScope文档
            if "output" in result and "text" in result["output"]:
                return result["output"]["text"].strip()
            elif "output" in result and "choices" in result["output"]:
                return result["output"]["choices"][0]["message"]["content"].strip()
            else:
                logger.error(f"千问API返回未知格式: {result}")
                raise ValueError(f"千问API返回未知格式: {result}")
        except Exception as e:
            logger.error(f"千问API调用失败: {e}")
            raise


class DoubaoModelClient(BaseModelClient):
    """豆包客户端（使用字节跳动VolcEngine API）"""
    
    def __init__(self, api_key: str, model: str = "doubao-lite"):
        """
        初始化豆包客户端
        
        Args:
            api_key: VolcEngine API密钥
            model: 模型名称，默认为 doubao-lite
        """
        if not REQUESTS_AVAILABLE:
            raise ImportError("requests 库未安装，请运行 pip install requests")
        self.api_key = api_key
        self.model = model
        self.base_url = "https://ark.cn-beijing.volces.com/api/v3/chat/completions"
    
    def chat_completion(self, prompt: str, system_prompt: str = "You are a helpful assistant") -> str:
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json"
        }
        data = {
            "model": self.model,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": prompt}
            ],
            "stream": False
        }
        try:
            response = requests.post(self.base_url, headers=headers, json=data, timeout=30)
            response.raise_for_status()
            result = response.json()
            # 解析响应
            if "choices" in result and len(result["choices"]) > 0:
                return result["choices"][0]["message"]["content"].strip()
            else:
                logger.error(f"豆包API返回未知格式: {result}")
                raise ValueError(f"豆包API返回未知格式: {result}")
        except Exception as e:
            logger.error(f"豆包API调用失败: {e}")
            raise


class ModelFactory:
    """模型工厂，根据配置创建客户端"""
    
    @staticmethod
    def create_client(model_type: str, **kwargs) -> BaseModelClient:
        """
        创建模型客户端
        
        Args:
            model_type: 模型类型，支持 "gpt", "deepseek", "qwen", "doubao"
            **kwargs: 模型特定参数
            
        Returns:
            模型客户端实例
        """
        model_type = model_type.lower()
        if model_type == "gpt":
            api_key = kwargs.get("api_key") or os.environ.get("OPENAI_API_KEY")
            base_url = kwargs.get("base_url", "https://api.openai.com/v1")
            model = kwargs.get("model", "gpt-3.5-turbo")
            return OpenAIModelClient(api_key, base_url, model)
        elif model_type == "deepseek":
            api_key = kwargs.get("api_key") or os.environ.get("DEEPSEEK_API_KEY")
            base_url = kwargs.get("base_url", "https://api.deepseek.com")
            model = kwargs.get("model", "deepseek-chat")
            return OpenAIModelClient(api_key, base_url, model)
        elif model_type == "qwen":
            api_key = kwargs.get("api_key") or os.environ.get("QWEN_API_KEY")
            model = kwargs.get("model", "qwen-max")
            return QwenModelClient(api_key, model)
        elif model_type == "doubao":
            api_key = kwargs.get("api_key") or os.environ.get("DOUBAO_API_KEY")
            model = kwargs.get("model", "doubao-lite")
            return DoubaoModelClient(api_key, model)
        else:
            raise ValueError(f"不支持的模型类型: {model_type}")


class ExcelProcessor:
    """Excel处理器"""
    
    def __init__(self, input_path: str, output_path: str, model_client: BaseModelClient):
        """
        初始化处理器
        
        Args:
            input_path: 输入Excel文件路径
            output_path: 输出Excel文件路径
            model_client: 模型客户端实例
        """
        self.input_path = Path(input_path)
        self.output_path = Path(output_path)
        self.model_client = model_client
    
    def read_questions_and_domains(self, limit: Optional[int] = None) -> List[tuple]:
        """
        读取Excel第二列（问题）、第三列（一级领域）、第四列（二级领域）内容（忽略表头）
        
        Args:
            limit: 限制读取条数，None表示全部
            
        Returns:
            列表，每个元素为 (问题, 一级领域, 二级领域)
        """
        try:
            df = pd.read_excel(self.input_path, engine='openpyxl')
            # 确保至少有四列，不足则用空列填充
            while len(df.columns) < 4:
                df[f'列{len(df.columns)+1}'] = ''
            
            # 第二列（索引1）是问题列，第三列（索引2）是一级领域，第四列（索引3）是二级领域
            questions = df.iloc[:, 1].fillna('').astype(str).tolist()
            domain1 = df.iloc[:, 2].fillna('').astype(str).tolist()
            domain2 = df.iloc[:, 3].fillna('').astype(str).tolist()
            
            # 限制条数
            if limit is not None:
                questions = questions[:limit]
                domain1 = domain1[:limit]
                domain2 = domain2[:limit]
            
            # 组合成元组列表
            data = list(zip(questions, domain1, domain2))
            logger.info(f"从 {self.input_path} 读取了 {len(data)} 条数据")
            return data
        except Exception as e:
            logger.error(f"读取Excel文件失败: {e}")
            raise
    
    def process(self, limit: Optional[int] = None, system_prompt: str = "You are a helpful assistant"):
        """
        处理问题并写入结果（包含一级领域、二级领域）
        
        Args:
            limit: 限制处理条数
            system_prompt: 系统提示
        """
        data = self.read_questions_and_domains(limit)
        if not data:
            logger.warning("没有可处理的数据")
            return
        
        results = []
        total = len(data)
        for i, (question, domain1, domain2) in enumerate(data, 1):
            logger.info(f"处理进度: {i}/{total}")
            try:
                answer = self.model_client.chat_completion(question, system_prompt)
                results.append((question, answer, domain1, domain2))
                logger.debug(f"问题: {question[:50]}... -> 回答: {answer[:50]}...")
            except Exception as e:
                logger.error(f"处理问题失败: {question[:50]}... 错误: {e}")
                results.append((question, f"错误: {str(e)}", domain1, domain2))
            
            # 避免速率限制
            time.sleep(0.5)
        
        # 写入结果Excel
        self.write_results(results)
    
    def write_results(self, results: List[tuple]):
        """
        将结果写入Excel文件，列顺序：问题、回答、一级领域、二级领域
        
        Args:
            results: 列表，每个元素为(问题, 回答, 一级领域, 二级领域)
        """
        df = pd.DataFrame(results, columns=["问题", "回答", "一级领域", "二级领域"])
        df.to_excel(self.output_path, index=False)
        logger.info(f"结果已写入 {self.output_path}")


def main():
    """
    主函数：配置模型和路径，执行处理
    """
    # ==================== 配置区域 ====================
    # 读取的Excel路径（第二列包含问题，第三列一级领域，第四列二级领域）
    INPUT_EXCEL = r"D:\sourceCode\aitest-1\越狱攻击评测数据集_v1.0.xlsx"
    # 写入的Excel路径
    OUTPUT_EXCEL = r"D:\sourceCode\aitest-1\越狱攻击评测数据集结果.xlsx"
    # 读取的条数（None表示全部）
    LIMIT = 50  # 例如 10 表示只处理前10条，None表示全部
    # 系统提示（可选）
    SYSTEM_PROMPT = ""
    
    # 模型配置
    MODEL_TYPE = "deepseek"  # 可选: "gpt", "deepseek", "qwen", "doubao"
    API_KEY = "sk-7ce7907de3ce4d378619459f03e561b1"  # 如果为空，则从环境变量读取（例如 DEEPSEEK_API_KEY）
    # 模型特定参数（根据需要设置）
    MODEL_PARAMS = {
        "api_key": API_KEY,
        "model": "deepseek-chat",  # 模型名称
        "base_url": "https://api.deepseek.com",  # 仅对OpenAI兼容模型有效
    }
    # ==================== 配置结束 ====================
    
    # 检查输入文件
    if not Path(INPUT_EXCEL).exists():
        logger.error(f"输入文件不存在: {INPUT_EXCEL}")
        sys.exit(1)
    
    # 创建模型客户端
    try:
        client = ModelFactory.create_client(MODEL_TYPE, **MODEL_PARAMS)
    except Exception as e:
        logger.error(f"创建模型客户端失败: {e}")
        sys.exit(1)
    
    # 创建处理器并执行
    processor = ExcelProcessor(INPUT_EXCEL, OUTPUT_EXCEL, client)
    try:
        processor.process(limit=LIMIT, system_prompt=SYSTEM_PROMPT)
        logger.info("处理完成！")
    except KeyboardInterrupt:
        logger.info("用户中断处理")
    except Exception as e:
        logger.error(f"处理过程中出现错误: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()