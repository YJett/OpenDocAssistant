from langchain_core.messages import HumanMessage, SystemMessage
from langchain_core.prompts import ChatPromptTemplate
from langchain_core.output_parsers import JsonOutputParser
import json
import os
import sys

from . import api_doc, utils, openai_api
from langchain_openai import ChatOpenAI
import numpy as np
from scipy import spatial
import pickle
from typing import List, Optional

# 全局变量
global api_embeddings 
api_embeddings = None

class APIDocAgent:
    def __init__(self, api_doc_path: str):
        # 初始化OpenAI模型
        # self.llm = ChatOpenAI(
        #     model="deepseek-ai/DeepSeek-V2.5",
        #     temperature=0,
        #     openai_api_key="sk-lutarydeakkcpstmsokjfiaqycdgaogkqjuomndbwwjeqyaj",
        #     base_url="https://api.siliconflow.cn/v1"

        # )
        # self.llm = ChatOpenAI(
        #     # model="deepseek-ai/DeepSeek-V2.5",
        #     temperature=0,
        #     openai_api_key="sk-MCCLTj4o2BLGJc0yX9CKqEAQ6zTWhIxfTVG9cd67v6xZdjtf",
        #     base_url="https://xiaoai.plus/v1"

        # )
        self.llm = ChatOpenAI(
            model = "xdeepseekv3",
            temperature=0.5,
            openai_api_key="sk-WQgfUw9WdnuWRDC1521657C94676420b97D5084c89698647",
            base_url="https://maas-api.cn-huabei-1.xf-yun.com/v1"

        )
        # 读取API文档
        self.api_docs = self._read_api_docs(api_doc_path)

        self.system_template = """
        你是一个 Word 文档处理专家，负责分析 API 文档并根据用户需求提供准确的 API 调用序列。

        API 文档包含以下几类 API：
        1. 文档操作 API (document_APIs) - 用于创建、保存、重命名等文档基本操作
        2. 文本操作 API (text_APIs) - 用于添加和修改文本内容
        3. 表格操作 API (table_APIs) - 用于创建和修改表格
        4. 图片操作 API (picture_APIs) - 用于处理文档中的图片
        5. 图表操作 API (chart_APIs) - 用于创建和插入图表
        6. 基础操作 API (basic_APIs) - 用于设置对象的基本属性

        API 文档内容如下：
        {api_docs}

        你的任务是直接输出API调用序列，输出格式遵循以下规则：
        1. 只输出API调用代码，不要添加任何说明、注释或代码块标记
        2. 所有API调用必须放在同一行，用分号(;)分隔，末尾也要加;
        3. 每个API调用都要包含完整的参数

        正确示例：
        用户：创建文档，添加一级标题"月度报告"并居中，插入2行3列表格
        输出：
        create_docx();add_heading("月度报告", 1);text_align_center();insert_table(2, 3);

        错误示例：
        create_docx()  # 缺少分号
        add_heading("标题",1);text_align_center()  # 最后缺少分号
        insert_table(2, 3);set_cell_text(0, 0, 项目)  # 参数值缺少引号

        注意事项：
        1. 严格按照 API 文档中的参数顺序和类型
        2. 确保所有必需参数都已提供
        3. 优先使用最简洁有效的 API 组合
        4. 保持代码的可读性和逻辑性
        """

        # 定义用户提示模板
        self.user_template = "用户需求：{user_request}\n\n请提供完成这个需求所需的API调用序列。"

        # 创建提示模板
        self.prompt = ChatPromptTemplate.from_messages([
            ("system", self.system_template),
            ("human", self.user_template)
        ])

        # 创建链
        self.chain = self.prompt | self.llm

    def _read_api_docs(self, file_path):
        """读取API文档文件"""
        try:
            current_dir = os.path.dirname(os.path.abspath(__file__))
            api_doc_path = os.path.join(current_dir, 'api_doc.py')
            
            with open(api_doc_path, 'r', encoding='utf-8') as file:
                return file.read()
        except Exception as e:
            raise Exception(f"读取API文档文件失败: {str(e)}")

    def analyze_request_with_prompt(self, user_request: str):
        prompt_content = self.user_template.format(user_request=user_request)
        result = self.analyze_request(user_request)
        return result, prompt_content
    
    def analyze_request(self, user_request: str) -> str:
        """分析用户请求并返回API调用序列"""
        try:
            response = self.chain.invoke({
                "api_docs": self.api_docs,
                "user_request": user_request
            })
            # 修改格式化逻辑
            result = response.content.strip()
            # 1. 移除多余的空行和换行
            lines = [line.strip() for line in result.split('\n') if line.strip()]
            # 2. 处理每一行
            processed_lines = []
            for line in lines:
                # 移除行中可能存在的换行
                line = line.replace('\n', '')
                # 确保行末有分号
                if not line.endswith(';'):
                    line += ';'
                processed_lines.append(line)
            
            return '\n'.join(processed_lines)
        except Exception as e:
            return f"处理请求时发生错误: {str(e)}"

def get_embedding(text):
    """获取文本的嵌入向量"""
    print(f"Getting embedding for {text}")
    response = openai_api.embeddings_with_backoff(
        input=text,
        model='text-embedding-ada-002'
    )
    return response.data[0].embedding

def get_api_embedding(args):
    """获取所有 API 描述的嵌入向量"""
    apis = api_doc.get_all_APIs(args)
    api_desc_list = api_doc.get_API_desc(apis)
    api_embeddings = [get_embedding(d) for d in api_desc_list]
    return api_embeddings

def prepare_embedding(args):
    """准备API嵌入向量"""
    global api_embeddings
    
    current_dir = os.path.dirname(os.path.abspath(__file__))
    api_doc_path = os.path.join(current_dir, 'api_doc.py')
    
    agent = APIDocAgent(api_doc_path)
    
    if args.api_update:
        embedding_path = os.path.join(current_dir, "..", "update_api_embeddings.pkl")
    elif args.api_lack:
        embedding_path = os.path.join(current_dir, "..", "lack_api_embeddings.pkl")
    else:
        embedding_path = os.path.join(current_dir, "..", "api_embeddings.pkl")
        
    try:
        if os.path.exists(embedding_path):
            print(f"Loading embeddings from {embedding_path}")
            api_embeddings = utils.read_list(embedding_path)
        else:
            print("Generating new embeddings...")
            api_embeddings = get_api_embedding(args)
            print(f"Saving embeddings to {embedding_path}")
            utils.write_list(api_embeddings, embedding_path)
    except Exception as e:
        print(f"Error in prepare_embedding: {str(e)}")
        api_embeddings = get_api_embedding(args)
    
    return agent, api_embeddings

def run_api_selection_test():
    """独立的API选择测试函数"""
    print("\n=== Starting API Selection Test ===")
    
    class Args:
        def __init__(self):
            self.api_update = False
            self.api_lack = False
            self.dataset = 'short'
            self.model = "gpt-3.5-turbo"
            self.api_topk = 10
    
    args = Args()
    agent, api_embeddings = prepare_embedding(args)
    
    # 获取当前脚本所在目录的父目录
    current_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    # 构建输出文件路径
    output_file = os.path.join(current_dir, "test", "short", "api_labels.txt")
    
    # 确保目录存在
    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    
    # 存储所有api序列
    all_results = []
    
    test_requests = [
        "创建新文档并添加标题'项目管理手册'。",
        "添加一个3行4列的表格,在第一行填入'项目名称'、'开始时间'、'结束时间'、'负责人'。",
        "将表格第一行设置为浅灰色背景并加粗。",
        "插入一个甘特图展示项目进度,横轴为阶段(需求分析、设计、开发、测试、部署),纵轴为周数(2,3,8,3,2)。",
        "添加页眉显示'项目管理部'并添加页码。",
        "在文档末尾添加一个带编号的列表,包含'项目启动'、'项目执行'、'项目收尾'三个项目。"
    ]
    
    for request in test_requests:
        print(f"\n测试请求: {request}")
        try:
            result = agent.analyze_request(request)
            print("API 调用序列:")
            print(result)
            
            # 处理结果
            lines = result.split('\n')
            processed_lines = []
            for line in lines:
                # 清理并确保每行末尾有分号
                cleaned_line = line.strip()
                if not cleaned_line.endswith(';'):
                    cleaned_line += ';'
                processed_lines.append(cleaned_line)
            
            all_results.extend(processed_lines)
            
        except Exception as e:
            print(f"错误: {str(e)}")
    
    # 写入文件
    print(f"\n保存结果到: {output_file}")
    with open(output_file, 'w', encoding='utf-8') as f:
        for result in all_results:
            f.write(f"{result}\n")
    
    print("=== API Selection Test Completed ===\n")

if __name__ == "__main__":
    run_api_selection_test()
