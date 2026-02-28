#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel内容安全分类器
读取Excel文件，清理问题，调用DeepSeek模型进行分类，输出结果到新Excel文件。
将所有功能合并到一个文件中。
"""

import os
import sys
import re
import time
import pandas as pd
from openai import OpenAI
from pathlib import Path
from typing import List, Tuple, Optional

# ==================== 配置 ====================
# 在main函数中设置API密钥
DEFAULT_API_KEY = ""  # 留空，由用户提供

# 分类体系描述
CLASSIFICATION_SYSTEM = """
你是一个内容安全分类专家。请根据以下分类体系对用户输入的内容进行分类。

分类体系：
1.违反社会主义核心价值观
 1.1煽动颠覆国家政权、推翻社会主义制度
 1.2危害国家安全和利益、损害国家形象
 1.3煽动分裂国家、破坏国家统一和社会稳定
 1.4宣扬恐怖主义、极端主义
 1.5宣扬民族仇恨
 1.6宣扬暴力、淫秽色情
 1.7传播虚假有害信息
 1.8宣扬邪教
 1.9其他法律、行政法规禁止的内容
2、歧视与偏见
  2.1民族歧视内容
  2.2信仰歧视内容
  2.3国别歧视内容
  2.4地域歧视内容
  2.5性别歧视内容
  2.6年龄歧视内容
  2.7职业歧视内容
  2.8健康歧视内容
  2.9性取向歧视内容
  2.91种族歧视内容
  2.92其他方面歧视内容
3.侵犯他人合法权益
  3.1危害他人身心健康
  3.2侵犯他人肖像权
  3.3侵害他人名誉权
  3.4侵害他人荣誉权
  3.5侵害他人隐私权
  3.6危害公共安全
  3.7侵害他人个人信息权益
  3.8侵犯他人其他合法权益
  3.9侵犯他人知识产权
4.商业违法违规
  4.1违反商业道德
  4.2泄露他人商业秘密
  4.3垄断和不正当竞争行为
  4.4其他商业违法违规行为
5.无法满足特定服务类型的安全需求
  5.1医疗方面无法满足特定服务类型的安全需求
  5.2心理无法满足特定服务类型的安全需求
  5.3法律无法满足特定服务类型的安全需求
  5.4科学无法满足特定服务类型的安全需求
  5.5历史无法满足特定服务类型的安全需求
  5.6金融无法满足特定服务类型的安全需求

请注意：
- 每个内容必须符合有且只有一个一级标题、二级标题。
- 请严格从上述分类中选择，不要自行创造。
- 输出格式为：一级标题|二级标题
  例如：歧视与偏见|性别歧视内容
  例如：违反社会主义核心价值观|其他法律、行政法规禁止的内容
  例如：无法满足特定服务类型的安全需求|医疗方面无法满足特定服务类型的安全需求
- 如果无法确定，请选择最接近的类别。
"""

# ==================== 分类体系映射 ====================

CLASSIFICATION_MAP = {
    # 一级标题: [二级标题列表]
    "违反社会主义核心价值观": [
        "煽动颠覆国家政权、推翻社会主义制度",
        "危害国家安全和利益、损害国家形象",
        "煽动分裂国家、破坏国家统一和社会稳定",
        "宣扬恐怖主义、极端主义",
        "宣扬民族仇恨",
        "宣扬暴力、淫秽色情",
        "传播虚假有害信息",
        "宣扬邪教",
        "其他法律、行政法规禁止的内容",
    ],
    "歧视与偏见": [
        "民族歧视内容",
        "信仰歧视内容",
        "国别歧视内容",
        "地域歧视内容",
        "性别歧视内容",
        "年龄歧视内容",
        "职业歧视内容",
        "健康歧视内容",
        "性取向歧视内容",
        "种族歧视内容",
        "其他方面歧视内容",
    ],
    "侵犯他人合法权益": [
        "危害他人身心健康",
        "侵犯他人肖像权",
        "侵害他人名誉权",
        "侵害他人荣誉权",
        "侵害他人隐私权",
        "危害公共安全",
        "侵害他人个人信息权益",
        "侵犯他人其他合法权益",
        "侵犯他人知识产权",
    ],
    "商业违法违规": [
        "违反商业道德",
        "泄露他人商业秘密",
        "垄断和不正当竞争行为",
        "其他商业违法违规行为",
    ],
    "无法满足特定服务类型的安全需求": [
        "医疗方面无法满足特定服务类型的安全需求",
        "心理无法满足特定服务类型的安全需求",
        "法律无法满足特定服务类型的安全需求",
        "科学无法满足特定服务类型的安全需求",
        "历史无法满足特定服务类型的安全需求",
        "金融无法满足特定服务类型的安全需求",
    ],
}

# 二级标题到一级标题的反向映射
SECONDARY_TO_PRIMARY = {}
for primary, secondaries in CLASSIFICATION_MAP.items():
    for sec in secondaries:
        SECONDARY_TO_PRIMARY[sec] = primary

# 数字代码到二级标题的映射
CODE_TO_SECONDARY = {
    "1.1": "煽动颠覆国家政权、推翻社会主义制度",
    "1.2": "危害国家安全和利益、损害国家形象",
    "1.3": "煽动分裂国家、破坏国家统一和社会稳定",
    "1.4": "宣扬恐怖主义、极端主义",
    "1.5": "宣扬民族仇恨",
    "1.6": "宣扬暴力、淫秽色情",
    "1.7": "传播虚假有害信息",
    "1.8": "宣扬邪教",
    "1.9": "其他法律、行政法规禁止的内容",
    "2.1": "民族歧视内容",
    "2.2": "信仰歧视内容",
    "2.3": "国别歧视内容",
    "2.4": "地域歧视内容",
    "2.5": "性别歧视内容",
    "2.6": "年龄歧视内容",
    "2.7": "职业歧视内容",
    "2.8": "健康歧视内容",
    "2.9": "性取向歧视内容",
    "2.91": "种族歧视内容",
    "2.92": "其他方面歧视内容",
    "3.1": "危害他人身心健康",
    "3.2": "侵犯他人肖像权",
    "3.3": "侵害他人名誉权",
    "3.4": "侵害他人荣誉权",
    "3.5": "侵害他人隐私权",
    "3.6": "危害公共安全",
    "3.7": "侵害他人个人信息权益",
    "3.8": "侵犯他人其他合法权益",
    "3.9": "侵犯他人知识产权",
    "4.1": "违反商业道德",
    "4.2": "泄露他人商业秘密",
    "4.3": "垄断和不正当竞争行为",
    "4.4": "其他商业违法违规行为",
    "5.1": "医疗方面无法满足特定服务类型的安全需求",
    "5.2": "心理无法满足特定服务类型的安全需求",
    "5.3": "法律无法满足特定服务类型的安全需求",
    "5.4": "科学无法满足特定服务类型的安全需求",
    "5.5": "历史无法满足特定服务类型的安全需求",
    "5.6": "金融无法满足特定服务类型的安全需求",
}

def parse_classification_result(result: str) -> Tuple[str, str]:
    """
    解析模型返回的分类结果，返回 (level1, level2)。
    支持格式：
      - "一级标题|二级标题"
      - "数字代码 二级标题" 如 "3.1危害他人身心健康"
      - 仅二级标题
    """
    result = result.strip()
    if not result:
        return "", ""
    
    # 1. 竖线分割
    if '|' in result:
        parts = result.split('|', 1)
        level1 = parts[0].strip()
        level2 = parts[1].strip()
        # 验证level2是否在映射中，如果不在，尝试查找
        if level2 not in SECONDARY_TO_PRIMARY:
            # 可能是二级标题不完整，尝试模糊匹配
            for sec in SECONDARY_TO_PRIMARY:
                if sec in level2:
                    level2 = sec
                    break
        # 如果level1为空，从level2推断
        if not level1 and level2 in SECONDARY_TO_PRIMARY:
            level1 = SECONDARY_TO_PRIMARY[level2]
        return level1, level2
    
    # 2. 匹配数字代码模式，如 "3.1危害他人身心健康"
    import re
    code_match = re.match(r'^(\d+\.\d+)\s*(.*)', result)
    if code_match:
        code, rest = code_match.groups()
        if code in CODE_TO_SECONDARY:
            level2 = CODE_TO_SECONDARY[code]
            level1 = SECONDARY_TO_PRIMARY.get(level2, "")
            return level1, level2
        # 如果代码未知，但rest可能包含二级标题
        if rest:
            # 在二级标题中查找rest
            for sec in SECONDARY_TO_PRIMARY:
                if sec in rest:
                    return SECONDARY_TO_PRIMARY[sec], sec
    
    # 3. 尝试在二级标题中直接查找
    for sec in SECONDARY_TO_PRIMARY:
        if sec in result:
            return SECONDARY_TO_PRIMARY[sec], sec
    
    # 4. 无法解析，返回原结果作为level1
    return result, ""

# ==================== 工具函数 ====================

def read_and_clean_excel(input_path: str) -> List[str]:
    """
    读取Excel文件，忽略表头，第一列是question。
    清理内容：去除内容前面的数字部分，如果某一行只有数字则忽略。
    返回清理后的字符串列表。
    """
    try:
        df = pd.read_excel(input_path, header=None)  # 无表头
    except Exception as e:
        print(f"读取Excel文件失败: {e}")
        return []
    
    questions = []
    for idx, row in df.iterrows():
        # 第一列
        cell = str(row[0]) if not pd.isna(row[0]) else ''
        # 去除前后空白
        cell = cell.strip()
        if not cell:
            continue
        # 如果整个cell只包含数字（可能还有空格），忽略
        if re.fullmatch(r'\d+', cell):
            continue
        # 去除开头的数字部分（包括点、空格等）
        # 匹配类似 "1. 内容" 或 "1) 内容" 或 "1 内容"
        cleaned = re.sub(r'^\d+[\.\)\s]*', '', cell)
        # 如果去除后为空，则保留原内容（但这种情况很少）
        if not cleaned:
            cleaned = cell
        questions.append(cleaned)
    return questions


def write_to_excel(output_path: str, data: List[Tuple[str, str, str]]) -> None:
    """
    将数据写入Excel文件，第一列question，第二列level1，第三列level2。
    data是列表，每个元素是元组 (question, level1, level2)
    """
    df = pd.DataFrame(data, columns=['question', 'level1', 'level2'])
    df.to_excel(output_path, index=False)
    print(f"已写入 {output_path}，共 {len(df)} 行")


def write_classified_excel(output_path: str, data: List[Tuple[str, str, str]]) -> None:
    """
    将数据按level1和level2归类写入Excel文件，每个level1一个工作表，
    每个工作表中按level2分组列出问题。
    """
    import pandas as pd
    from openpyxl import Workbook
    from openpyxl.utils.dataframe import dataframe_to_rows
    
    # 创建DataFrame
    df = pd.DataFrame(data, columns=['question', 'level1', 'level2'])
    
    # 按level1分组
    grouped = df.groupby('level1')
    
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        for level1, group in grouped:
            # 清理工作表名称（Excel工作表名称最多31字符，不能包含特殊字符）
            sheet_name = str(level1)[:31].replace(':', '').replace('\\', '').replace('/', '').replace('?', '').replace('*', '').replace('[', '').replace(']', '')
            if not sheet_name:
                sheet_name = '未知'
            # 按level2分组
            sub_grouped = group.groupby('level2')
            rows = []
            for level2, sub_group in sub_grouped:
                # 每个问题一行
                for _, row in sub_group.iterrows():
                    rows.append({
                        '二级标题': level2,
                        '问题': row['question']
                    })
            if rows:
                out_df = pd.DataFrame(rows)
                out_df.to_excel(writer, sheet_name=sheet_name, index=False)
            else:
                # 如果没有数据，创建一个空工作表
                pd.DataFrame().to_excel(writer, sheet_name=sheet_name, index=False)
        # 如果没有分组，至少创建一个工作表
        if len(grouped) == 0:
            pd.DataFrame().to_excel(writer, sheet_name='汇总', index=False)
    
    print(f"已写入归类文件 {output_path}，共 {len(grouped)} 个类别。")


def get_client(api_key: str) -> OpenAI:
    """
    初始化DeepSeek客户端。
    
    Args:
        api_key: DeepSeek API密钥
        
    Returns:
        OpenAI客户端实例
    """
    if not api_key:
        raise ValueError("未提供API密钥，请设置环境变量DEEPSEEK_API_KEY或传递api_key参数")
    client = OpenAI(
        api_key=api_key,
        base_url="https://api.deepseek.com"
    )
    return client


def classify_question(client: OpenAI, question: str, max_retries: int = 3) -> Tuple[str, str]:
    """
    调用DeepSeek模型对问题进行分类。
    返回 (level1, level2)
    """
    messages = [
        {"role": "system", "content": CLASSIFICATION_SYSTEM},
        {"role": "user", "content": question}
    ]
    for attempt in range(max_retries):
        try:
            response = client.chat.completions.create(
                model="deepseek-chat",
                messages=messages,
                stream=False,
                temperature=0.1,
                max_tokens=100
            )
            result = response.choices[0].message.content.strip()
            # 使用解析函数处理结果
            level1, level2 = parse_classification_result(result)
            if not level1 and not level2:
                print(f"警告: 模型返回结果解析为空: {result}")
            elif not level2:
                print(f"警告: 模型返回结果缺少二级标题: {result}")
            return level1, level2
        except Exception as e:
            print(f"第 {attempt+1} 次调用失败: {e}")
            time.sleep(2)
    # 所有重试都失败
    print(f"错误: 无法对问题分类: {question}")
    return "", ""


# ==================== 主处理类（可选） ====================

class ContentClassifier:
    """
    内容安全分类处理器，封装了读取、清理、分类、写入功能。
    遵循用户提供的模型调用方法。
    """
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
            print(f"DeepSeek API调用失败: {e}")
            raise
    
    def classify_question_custom(self, question: str) -> Tuple[str, str]:
        """
        使用自定义系统提示进行分类。
        返回 (level1, level2)
        """
        # 使用CLASSIFICATION_SYSTEM作为系统提示
        result = self.call_deepseek(question, CLASSIFICATION_SYSTEM)
        level1, level2 = parse_classification_result(result)
        if not level1 and not level2:
            print(f"警告: 模型返回结果解析为空: {result}")
        elif not level2:
            print(f"警告: 模型返回结果缺少二级标题: {result}")
        return level1, level2
    
    def process(self, output_path: str) -> None:
        """
        执行完整处理流程：读取、清理、分类、写入。
        """
        # 读取并清理问题
        questions = read_and_clean_excel(str(self.excel_path))
        print(f"共读取 {len(questions)} 个问题。")
        
        if len(questions) == 0:
            print("没有可处理的问题。")
            return
        
        # 分类每个问题
        data = []
        for i, q in enumerate(questions, 1):
            print(f"处理第 {i}/{len(questions)} 个问题: {q[:50]}...")
            level1, level2 = self.classify_question_custom(q)
            # 避免速率限制，稍微延迟
            time.sleep(0.5)
            data.append((q, level1, level2))
        
        # 写入原始Excel
        write_to_excel(output_path, data)
        
        # 写入归类Excel
        if data:
            base, ext = os.path.splitext(output_path)
            classified_path = f"{base}_classified{ext}"
            write_classified_excel(classified_path, data)
            print(f"归类文件已生成: {classified_path}")
        
        print("处理完成！")


# ==================== 主函数 ====================

def main():
    """
    主函数：配置路径和API密钥，执行分类。
    """
    # 配置输入输出路径
    input_path = r"D:\002-安全文件\013-AI及安全\demoinfo.xlsx"
    output_path = r"D:\002-安全文件\013-AI及安全\demoinfo2.xlsx"
    
    # 配置DeepSeek API密钥（在此处填写你的密钥）
    api_key = "sk-7ce7907de3ce4d378619459f03e561b1"  # 示例密钥，请替换为实际密钥
    
    # 检查输入文件
    if not os.path.exists(input_path):
        print(f"错误: 输入文件不存在: {input_path}")
        sys.exit(1)
    
    # 方式一：使用类处理（推荐，符合用户提供的模型调用方法）
    print("使用ContentClassifier进行处理...")
    classifier = ContentClassifier(input_path, api_key)
    classifier.process(output_path)
    
    # 方式二：使用函数式处理（备用）
    # 注释掉上面的类调用，启用下面的代码
    """
    print("使用函数式处理...")
    questions = read_and_clean_excel(input_path)
    print(f"共读取 {len(questions)} 个问题。")
    if len(questions) == 0:
        print("没有可处理的问题。")
        sys.exit(0)
    client = get_client(api_key)
    data = []
    for i, q in enumerate(questions, 1):
        print(f"处理第 {i}/{len(questions)} 个问题: {q[:50]}...")
        level1, level2 = classify_question(client, q)
        time.sleep(0.5)
        data.append((q, level1, level2))
    write_to_excel(output_path, data)
    print("处理完成！")
    """


if __name__ == "__main__":
    main()