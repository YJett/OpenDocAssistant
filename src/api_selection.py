from langchain_core.messages import HumanMessage, SystemMessage
from langchain_core.prompts import ChatPromptTemplate
from langchain_core.output_parsers import JsonOutputParser
import json
import os
import sys

# 添加项目根目录到 Python 路径
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src import api_doc, utils, openai_api
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
        # 修改初始化OpenAI模型，使用本地配置
        self.llm = ChatOpenAI(
            model="gpt-3.5-turbo",
            temperature=0,
            openai_api_key=openai_api.openai.api_key,
            openai_api_base=openai_api.openai.api_base
        )

        # 读取API文档
        self.api_docs = self._read_api_docs(api_doc_path)

        # 定义输出解析器
        self.output_parser = JsonOutputParser()

        # 定义系统提示模板，使用双大括号来转义JSON式示例
        self.system_template = """你是一个 Word 文档处理专家，负责分析 API 文档并根据用户需求提供准确的 API 调用序列。

API 文档包含以下几类 API：
1. 文档操作 API (document_APIs) - 用于创建、保存、重命名等文档基本操作
2. 文本操作 API (text_APIs) - 用于添加和修改文本内容
3. 表格操作 API (table_APIs) - 用于创建和修改表格
4. 图片操作 API (picture_APIs) - 用于处理文档中的图片
5. 图表操作 API (chart_APIs) - 用于创建和插入图表
6. 基础操作 API (basic_APIs) - 用于设置对象的基本属性

API 文档内容如下：
{api_docs}

你的任务是：
1. 理解用户的需求
2. 根据需求选择合适的 API 类别
3. 分析哪些 API 需要按顺序调用来完成用户需求
4. 对每个 API 调用，明确指出需要的参数和值

请以 JSON 格式返回，格式如下：
{{
    "steps": [
        {{
            "step": 1,
            "api_category": "API 类别",
            "api_name": "API 名称",
            "parameters": {{
                "参数名": "参数值或描述"
            }},
            "description": "这一步的作用说明"
        }}
    ]
}}"""

        # 定义用户提示模板
        self.user_template = """用户需求：{user_request}

                                请提供完成这个需求所需的API调用序列。"""

        # 创建提示模板
        self.prompt = ChatPromptTemplate.from_messages([
            ("system", self.system_template),
            ("human", self.user_template)
        ])

        # 创建链
        self.chain = (
                self.prompt
                | self.llm
                | self.output_parser
        )

    def _read_api_docs(self, file_path):
        """读取API文档文件"""
        try:
            # 使用绝对路径
            current_dir = os.path.dirname(os.path.abspath(__file__))
            api_doc_path = os.path.join(current_dir, 'api_doc.py')
            
            with open(api_doc_path, 'r', encoding='utf-8') as file:
                return file.read()
        except Exception as e:
            raise Exception(f"读取API文档文件失败: {str(e)}")

    def analyze_request(self, user_request: str) -> dict:
        """分析用户请求并返回API调用序列"""
        try:
            response = self.chain.invoke({
                "api_docs": self.api_docs,
                "user_request": user_request
            })
            return response
        except Exception as e:
            return {"error": f"处理请求时发生错误: {str(e)}"}

    @staticmethod
    def format_output(result: dict) -> str:
        """格式化输出结果"""
        if "error" in result:
            return f"错误: {result['error']}"

        output = "API 调用序列：\n\n"
        for step in result["steps"]:
            output += f"步骤 {step['step']}:\n"
            output += f"- API: {step['api_name']}\n"
            output += "- 参数:\n"
            for param, value in step['parameters'].items():
                output += f"  * {param}: {value}\n"
            output += f"- 说明: {step['description']}\n\n"

        return output

def get_embedding(text):
    """获取文本的嵌入向量"""
    print(f"Getting embedding for {text}")
    response = openai_api.embeddings_with_backoff(
        input=text,
        model='text-embedding-ada-002'
    )
    embeddings = response['data'][0]['embedding']
    return embeddings

def get_api_embedding(args):
    """获取所有 API 描述的嵌入向量"""
    apis = api_doc.get_all_APIs(args)
    api_desc_list = api_doc.get_API_desc(apis)
    api_embeddings = [get_embedding(d) for d in api_desc_list]
    return api_embeddings

def prepare_embedding(args):
    """准备API嵌入向量"""
    global api_embeddings
    
    # 使用绝对路径
    current_dir = os.path.dirname(os.path.abspath(__file__))
    api_doc_path = os.path.join(current_dir, 'api_doc.py')
    
    # 创建API文档代理
    agent = APIDocAgent(api_doc_path)
    
    # 准备API嵌入向量
    if args.api_update:
        embedding_path = os.path.join(current_dir, "..", "update_api_embeddings.pkl")
    elif args.api_lack:
        embedding_path = os.path.join(current_dir, "..", "lack_api_embeddings.pkl")
    else:
        embedding_path = os.path.join(current_dir, "..", "api_embeddings.pkl")
        
    # 尝试从缓存加载嵌入向量
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
        # 如果出错，生成新的嵌入向量
        api_embeddings = get_api_embedding(args)
    
    return agent, api_embeddings

def main():
    """测试 API 选择功能"""
    # 创建测试参数
    class Args:
        def __init__(self):
            self.api_update = False
            self.api_lack = False
            self.dataset = 'short'
            self.model = "gpt-3.5-turbo"
            self.api_topk = 10
    
    args = Args()
    
    # 准备嵌入向量
    prepare_embedding(args)
    
    # 创建 API 文档代理
    agent = APIDocAgent("src/api_doc.py")  # 替换为你的 API 文档路径
    
    # 测试一些用户请求
    test_requests = [
        "创建一个新文档并添加标题",
        "在文档中插入一个2x3的表格",
        "添加一个居中的段落，内容是'Hello World'"
    ]
    
    for request in test_requests:
        print(f"\n测试请求: {request}")
        try:
            result = agent.analyze_request(request)
            print("API 选结果:")
            print(json.dumps(result, indent=2, ensure_ascii=False))
        except Exception as e:
            print(f"处理请求时出错: {str(e)}")

if __name__ == "__main__":
    # main()
    pass
