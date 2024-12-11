import argparse
from docx import Document
from src import openai_api
import os
import pickle
import numpy as np
from src import api_doc, utils
import spacy
from typing import List, Optional

from scipy import spatial

# 全局变量
global api_embeddings 
global K
api_embeddings = None
K = None

def cosine_similarity(a, b):
    """计算两个向量的余弦相似度"""
    return 1 - spatial.distance.cosine(a, b)

def get_topk(scores, k=10):
    """获取前 k 个得分最高的索引"""
    sorted_idx = sorted(range(len(scores)), key=lambda k: scores[k], reverse=True)
    topk_idx = sorted_idx[:k]
    return topk_idx

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

def get_dynamic_k(query):
    """让大模型决定应该返回的 K 值"""
    prompt = f"""
    用户的查询是：{query}
    请根据查询的复杂度和可能需要的功能数量，返回一个建议的 API 数量（K）。
    K 的范围是 1 到 10。只返回一个整数即可。
    """
    response = openai_api.chat_with_backoff(
        messages=[{"role": "system", "content": "你是一个帮助选择 API 的助手。"},
                  {"role": "user", "content": prompt}],
        model="gpt-4"
    )
    try:
        k = int(response['choices'][0]['message']['content'].strip())
        return max(1, min(k, 10))  # 确保 K 在 1 到 10 的范围内
    except ValueError:
        return 5  # 如果模型返回异常，默认使用 5

def select_api(query):
    """根据查询动态选择 API"""
    global api_embeddings
    # 动态获取 K
    K = get_dynamic_k(query)
    print(f"Dynamic K determined by the model: {K}")
    
    # 获取查询的嵌入向量
    query_embedding = get_embedding(query)
    # 计算相似度得分
    scores = [cosine_similarity(query_embedding, e) for e in api_embeddings]
    print(f"Scores: {scores[:5]}...")  # 打印前 5 个相似度得分
    # 获取最相关的 K 个 API
    topk_idx = get_topk(scores, k=K)
    return topk_idx

def get_selected_apis(instruction, args):
    """主逻辑：根据指令选取 API"""
    args = argparse.Namespace(**args)
    # 使用动态 K 的 API 选择
    topk_idx = select_api(instruction)
    apis = api_doc.get_all_APIs(args)
    selected_apis = [apis[x] for x in topk_idx]
    must_apis = api_doc.get_must_APIs(args)
    ans = []
    cnt = 0
    for api in selected_apis:
        if api not in must_apis:
            ans.append(api)
            cnt += 1
            if cnt == len(selected_apis) - len(must_apis):
                break
    for api in ans:
        print(f"Selected API: {api.name}, Description: {api.description}, Parameters: {api.parameters}")
    return ans

def prepare_embedding(args):
    """准备 API 嵌入向量"""
    args = argparse.Namespace(**args)
    global api_embeddings, K
    if args.api_update:
        embedding_path = f"update_api_embeddings.pkl"
    elif args.api_lack:
        embedding_path = f"lack_api_embeddings.pkl" 
    else:
        embedding_path = f"api_embeddings.pkl"
    if os.path.exists(embedding_path):
        api_embeddings = utils.read_list(embedding_path)
    else:
        api_embeddings = get_api_embedding(args)
        utils.write_list(api_embeddings, embedding_path)

if __name__ == '__main__':
    query = "帮我创建一个饼图和折线图"

    # 设置参数
    args = {
        'api_update': False,
        'api_lack': False,
        # 'api_topk': 2,  # 不再直接使用
        'dataset': 'long'
    }

    # 获取相关的 API 嵌入
    prepare_embedding(args)

    # 获取选中的 API
    selected_apis = get_selected_apis(query, args)

    # 输出选中的 API
    print("Selected APIs:", selected_apis)
