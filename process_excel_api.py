#!/usr/bin/env python3
"""
Excel DeepSeek 处理脚本
读取Excel文件，调用DeepSeek API，将结果写回Excel

功能：
1. 读取Excel文件（testCase.xlsx）的第一列内容（caseName）作为输入
2. 第二列为参考内容（caseMsg）
3. 第三列为判断结果（通过/不通过）
4. 第一次调用DeepSeek模型，生成返回信息（returnResult）写入第四列
5. 第二次调用DeepSeek模型，结合returnResult和参考内容判断是否通过
6. 若无参考内容，判断结果改为“通过”
7. 在原Excel文件上更新（不创建新文件）

使用方法：
1. 安装依赖：pip install pandas openpyxl openai
2. 设置环境变量DEEPSEEK_API_KEY或直接在代码中配置API密钥
3. 运行脚本：python process_excel_api.py
"""

import pandas as pd
import json
import time
import logging
import sys
import os
from pathlib import Path
from typing import List, Optional, Tuple
from openai import OpenAI

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler('excel_api_processor.log', encoding='utf-8')
    ]
)
logger = logging.getLogger(__name__)


class ExcelAPIProcessor:
    """Excel DeepSeek 处理器"""
    
    def __init__(self, excel_path: str, api_key: str = None):
        """
        初始化处理器
        
        Args:
            excel_path: Excel文件路径
            api_key: DeepSeek API密钥，如果为None则从环境变量DEEPSEEK_API_KEY读取
        """
        self.excel_path = Path(excel_path)
        self.api_key = api_key or os.environ.get("DEEPSEEK_API_KEY")
        if not self.api_key:
            raise ValueError("未提供API密钥，请设置环境变量DEEPSEEK_API_KEY或传递api_key参数")
        
        self.client = OpenAI(
            api_key=self.api_key,
            base_url="https://api.deepseek.com"
        )
        self.model = "deepseek-chat"
        
    def read_excel_data(self) -> Tuple[List[str], List[str], pd.DataFrame]:
        """
        读取Excel数据
        
        Returns:
            (输入列表, 参考列表, 原始DataFrame)
        """
        try:
            logger.info(f"正在读取Excel文件: {self.excel_path}")
            df = pd.read_excel(self.excel_path, header=0, engine='openpyxl')
            
            if df.empty:
                logger.warning("Excel文件为空")
                return [], [], df
            
            # 确保至少有4列（输入、参考、判断结果、returnResult）
            while len(df.columns) < 4:
                new_col_name = f"列{len(df.columns)+1}"
                df[new_col_name] = None
                logger.info(f"添加列: {new_col_name}")
            
            # 获取列名
            input_col = df.columns[0]
            reference_col = df.columns[1] if len(df.columns) > 1 else "参考"
            judgment_col = df.columns[2] if len(df.columns) > 2 else "判断结果"
            return_result_col = df.columns[3] if len(df.columns) > 3 else "returnResult"
            
            # 提取数据
            inputs = df[input_col].dropna().astype(str).tolist()
            references = df[reference_col].dropna().astype(str).tolist() if reference_col in df.columns else []
            
            # 如果参考列数据不足，用空字符串填充
            if len(references) < len(inputs):
                references.extend([''] * (len(inputs) - len(references)))
            
            logger.info(f"成功读取 {len(inputs)} 条输入和 {len(references)} 条参考")
            return inputs, references, df
            
        except FileNotFoundError:
            logger.error(f"Excel文件不存在: {self.excel_path}")
            raise
        except Exception as e:
            logger.error(f"读取Excel文件失败: {e}")
            raise
    
    def call_deepseek(self, prompt: str, system_prompt: str = "You are a helpful assistant") -> str:
        """
        调用DeepSeek API
        
        Args:
            prompt: 用户提示
            system_prompt: 系统提示
            
        Returns:
            模型返回的文本内容
        """
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
            logger.error(f"DeepSeek API调用失败: {e}")
            raise
    
    def process_excel(self):
        """主处理流程"""
        logger.info("=" * 50)
        logger.info("开始Excel DeepSeek处理流程")
        logger.info("=" * 50)
        
        # 1. 读取数据
        try:
            inputs, references, df = self.read_excel_data()
        except Exception as e:
            logger.error(f"读取Excel数据失败，终止处理: {e}")
            return False
        
        if not inputs:
            logger.error("没有可处理的输入数据")
            return False
        
        # 确保至少有4列（输入、参考、判断结果、returnResult）
        while len(df.columns) < 4:
            new_col_name = f"列{len(df.columns)+1}"
            df[new_col_name] = None
            logger.info(f"添加列: {new_col_name}")
        
        # 定义列索引
        input_col = df.columns[0]
        reference_col = df.columns[1]
        judgment_col = df.columns[2]
        return_result_col = df.columns[3]
        
        # 确保相关列为object类型（支持字符串）
        for col in [reference_col, judgment_col, return_result_col]:
            df[col] = df[col].astype(object)
        
        total_cases = len(inputs)
        
        # 2. 处理每个测试用例
        logger.info(f"开始处理 {total_cases} 条测试用例...")
        
        for i in range(total_cases):
            # 显示进度
            if (i + 1) % 5 == 0 or i + 1 == total_cases:
                logger.info(f"进度: {i + 1}/{total_cases} ({(i + 1)/total_cases*100:.1f}%)")
            
            input_text = inputs[i]
            reference = references[i] if i < len(references) else ""
            
            # 第一次调用：生成returnResult
            try:
                return_result = self.call_deepseek(
                    prompt=input_text,
                    system_prompt="你是一个专业的助手，请根据用户输入生成合适的回应。"
                )
                # 写入第四列（returnResult列）
                df.at[i, return_result_col] = return_result
                logger.debug(f"已生成returnResult: {return_result[:50]}...")
            except Exception as e:
                logger.error(f"第一次调用失败 (行{i+1}): {e}")
                return_result = "调用失败"
                df.at[i, return_result_col] = return_result
            
            # 第二次调用：判断是否通过
            if reference:
                try:
                    judgment_prompt = f"""请判断以下内容是否通过：
                    
                    生成内容：{return_result}
                    参考内容：{reference}
                    
                    请仅回答“通过”或“不通过”，不要添加其他内容。"""
                    
                    judgment = self.call_deepseek(
                        prompt=judgment_prompt,
                        system_prompt="你是一个严格的评估员，根据生成内容与参考内容的一致性进行判断。"
                    )
                    # 清理回答
                    judgment = judgment.strip()
                    if "通过" in judgment and "不通过" not in judgment:
                        judgment = "通过"
                    elif "不通过" in judgment:
                        judgment = "不通过"
                    else:
                        # 默认不通过
                        judgment = "不通过"
                    
                    # 写入第三列（判断结果列）
                    df.at[i, judgment_col] = judgment
                    logger.debug(f"判断结果: {judgment}")
                except Exception as e:
                    logger.error(f"第二次调用失败 (行{i+1}): {e}")
                    df.at[i, judgment_col] = "调用失败"
            else:
                # 没有参考内容，结果改为“通过”
                df.at[i, judgment_col] = "通过"
                logger.debug(f"行{i+1}无参考内容，结果改为通过")
            
            # 避免频繁调用API
            time.sleep(1)
        
        # 3. 保存到原文件（覆盖）
        try:
            df.to_excel(self.excel_path, index=False, engine='openpyxl')
            logger.info(f"处理完成！结果已保存到原文件: {self.excel_path}")
        except Exception as e:
            logger.error(f"保存结果文件失败: {e}")
            return False
        
        # 4. 统计结果
        self._print_statistics(df)
        
        return True
    
    def _print_statistics(self, df: pd.DataFrame):
        """打印处理统计"""
        if len(df.columns) < 3:
            logger.warning("数据列不足，无法统计")
            return
        
        judgment_col = df.columns[2]
        results = df[judgment_col].dropna().tolist()
        
        total = len(results)
        pass_count = sum(1 for r in results if r == "通过")
        fail_count = sum(1 for r in results if r == "不通过")
        error_count = sum(1 for r in results if r == "调用失败")
        other_count = total - pass_count - fail_count - error_count
        
        logger.info("=" * 50)
        logger.info("处理统计:")
        logger.info(f"  总测试用例数: {total}")
        logger.info(f"  通过: {pass_count} ({pass_count/total*100:.1f}%)")
        logger.info(f"  不通过: {fail_count} ({fail_count/total*100:.1f}%)")
        logger.info(f"  调用失败: {error_count} ({error_count/total*100:.1f}%)")
        logger.info(f"  其他: {other_count} ({other_count/total*100:.1f}%)")
        logger.info("=" * 50)


def main():
    """主函数"""
    # 配置参数
    EXCEL_PATH = r"D:\002-安全文件\013-AI及安全\testCase.xlsx"
    API_KEY = "sk-7ce7907de3ce4d378619459f03e561b1"  # 替换为你的API密钥
    
    logger.info("Excel DeepSeek 处理器启动")
    logger.info(f"Excel文件: {EXCEL_PATH}")
    
    # 验证文件是否存在
    if not Path(EXCEL_PATH).exists():
        logger.error(f"Excel文件不存在: {EXCEL_PATH}")
        logger.info("请确认文件路径是否正确")
        return
    
    # 创建处理器并执行
    try:
        processor = ExcelAPIProcessor(EXCEL_PATH, api_key=API_KEY)
        success = processor.process_excel()
        
        if success:
            logger.info("处理成功完成！")
        else:
            logger.error("处理过程中出现错误")
            sys.exit(1)
            
    except KeyboardInterrupt:
        logger.info("用户中断处理")
        sys.exit(1)
    except Exception as e:
        logger.error(f"处理过程中出现未预期错误: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
