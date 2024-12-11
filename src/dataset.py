import os
import json
from docx import Document
import sys
import logging

# 设置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# 添加项目根目录到 Python 路径
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src import utils

def load_data(path, dataset, args):
    """加载文本格式的数据集"""
    try:
        # 获取项目根目录
        project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        
        # 构建文件路径
        if args.robust:
            instruction_path = os.path.join(project_root, 'test', dataset, f'robust{args.robust_num}.txt')
        elif args.noisy:
            instruction_path = os.path.join(project_root, 'test', dataset, 'noisy.txt')
        else:
            instruction_path = os.path.join(project_root, 'test', dataset, 'instructions.txt')
            
        label_path = os.path.join(project_root, 'test', dataset, 'api_labels.txt')
        
        logger.info(f"Loading instructions from: {instruction_path}")
        logger.info(f"Loading labels from: {label_path}")
        
        instructions = []
        labels = []

        # 读取指令文件
        with open(instruction_path, 'r') as f:
            lines = f.read()
            dialogues = lines.strip().split('\n\n')
            for dialogue in dialogues:
                ins = dialogue.split('\n')
                ins = [x.strip() for x in ins if x != '']
                instructions.append(ins)
            
        # 读取标签文件
        with open(label_path, 'r') as f:
            lines = f.read()
            dialogue_labels = lines.strip().split('\n\n')
            for dialogue_label in dialogue_labels:
                apis = dialogue_label.split('\n')
                api_list = [x.strip(';').split(';') for x in apis]
                labels.append(api_list)
        
        return instructions, labels
        
    except Exception as e:
        logger.error(f"Error loading data: {e}")
        logger.error(f"Current working directory: {os.getcwd()}")
        raise

def load_data_json(path, dataset):
    """加载JSON格式的数据集
    
    Args:
        path: 数据路径
        dataset: 数据集类型 ('short'/'long')
    Returns:
        instructions: 用户指令列表
        labels: API序列标签列表
        doc_labels: 文档标签列表
    """
    instructions = []
    labels = []
    doc_labels = []

    # 根据数据集类型选择目录
    dir = 'Create_new_docs' if dataset == 'short' else 'Edit_Word_template'
    read_path = os.path.join(path, dir)
    
    # 遍历所有JSON文件
    for json_file in sorted(os.listdir(read_path), key=lambda x: int(x.split('_')[1].split('.')[0])):
        if json_file.endswith('.json'):
            print(f"Processing {json_file}")
            with open(os.path.join(read_path, json_file), 'r') as f:
                session_data = f.read().strip().split('\n')
                
                session_instructions = []
                session_api_labels = []
                session_doc_labels = []
                
                for line in session_data:
                    if not line:
                        continue
                        
                    data = json.loads(line)
                    turn = data['Turn']
                    instruction = data['User instruction']
                    api_sequence = data['Feasible API sequence']
                    label_file = data['Label File']
                    
                    print(f"Turn: {turn}")
                    print(f"Instruction: {instruction}")
                    print(f"API Sequence: {api_sequence}")
                    print(f"Label File: {label_file}")
                    
                    session_instructions.append(instruction)
                    session_api_labels.append(api_sequence)
                    session_doc_labels.append(label_file)
                
                instructions.append(session_instructions)
                labels.append(session_api_labels)
                doc_labels.append(session_doc_labels)
    
    return instructions, labels, doc_labels

def load_docx(path, dataset, args):
    """加载Word文档数据集
    
    Args:
        path: Word文档根目录路径
        dataset: 数据集类型 ('short'/'long')
        args: 参数配置
    Returns:
        paragraphs: 段落文本列表
        styles: 样式信息列表
    """
    paragraphs = []
    styles = []
    
    # 根据参数选择不同的文档路径
    if args.robust:
        doc_path = os.path.join(path, dataset, f'robust{args.robust_num}.docx')
    elif args.noisy:
        doc_path = os.path.join(path, dataset, 'noisy.docx')
    else:
        doc_path = os.path.join(path, dataset, 'document.docx')

    # 读取Word文档
    doc = Document(doc_path)
    
    # 提取段落文本和样式
    for paragraph in doc.paragraphs:
        paragraphs.append(paragraph.text)
        style = {
            'alignment': paragraph.alignment,
            'style': paragraph.style.name,
            'font': paragraph.runs[0].font.name if paragraph.runs else None,
            'size': paragraph.runs[0].font.size if paragraph.runs else None,
            'bold': paragraph.runs[0].font.bold if paragraph.runs else None,
            'italic': paragraph.runs[0].font.italic if paragraph.runs else None
        }
        styles.append(style)

    return paragraphs, styles

if __name__ == '__main__':
    # 测试代码
    # class Args:
    #     def __init__(self):
    #         self.robust = False
    #         self.noisy = False
    #         self.robust_num = 0
    
    # args = Args()
    
    # # 测试文本数据加载
    # instructions, labels = load_data("data", "short", args)
    # print("\nText data test:")
    # print(f"Loaded {len(instructions)} dialogues")
    
    # # 测试JSON数据加载
    # instructions, labels, doc_labels = load_data_json("json_data", "short")
    # print("\nJSON data test:")
    # print(f"Loaded {len(instructions)} dialogues")
    
    # # 测试Word文档加载
    # paragraphs, styles = load_docx("docs", "short", args)
    # print("\nWord document test:")
    # print(f"Loaded {len(paragraphs)} paragraphs")
    load_data_json("json_file", "short")