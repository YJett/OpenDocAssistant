# 导入所需的库和模块
import pickle  # 用于序列化和反序列化Python对象
import os
from src import dataset # 从src目录导入
from src import api_doc # 从src目录导入
from sacremoses import MosesTokenizer  # 用于文本分词
import tiktoken  # OpenAI的分词器
import json     # 用于JSON数据处理

# 将列表写入二进制文件
def write_list(lst, filename):
    with open(filename, "wb") as f:
        pickle.dump(lst, f)

# 从二进制文件读取列表
def read_list(filename):
    with open(filename, 'rb') as f:
        lst = pickle.load(f)
    return lst

# 将列表中的字符串逐行写入文本文件
def write_lines(lst, path):
    with open(path, 'w') as file:
        for s in lst:
            file.write(s)
            file.write('\n')

# 从文本文件读取内容到列表中
def read_lines(path):
    data = []
    with open(path, 'r') as file:
        for line in file:
            data.append(line.strip())
    
    return data

# 创建目录(如果不存在)
def makedir(path):
    if not os.path.exists(path):
        os.makedirs(path)

# 将多个列表合并一个列表
def merge_list(lst):
    merged = []
    for l in lst:
        merged.extend(l)
    return merged

# 从API调用列表中提取图片名称
def get_picture_name(labels):
    # labels: all api list
    picture_name = []
    for api in labels:
        if "insert_picture(" in api:
            try:
                # 处理双引号情况
                start_idx = api.index('insert_picture("')
                end_idx = api.index('")')
                s = api[start_idx+len('insert_picture("'):end_idx]
            except:
                # 处理单引号情况
                start_idx = api.index("insert_picture('")
                end_idx = api.index("')")
                s = api[start_idx+len("insert_picture('"):end_idx]
            picture_name.append(s)
    return picture_name

# 获取数据集中所有图片名称列表
def get_picture_name_list(args):
    instruction_data, label_data = dataset.load_data_v2(args.data_path)
    label_list = []
    for x in label_data:
        for xx in x:
            label_list.extend(xx)
    labels = get_picture_name(label_list)
    write_lines(labels, f"{args.data_path}/pictures.txt")
    return labels

# 从代码文本中解析API调用
def parse_api(codes):
    apis = []
    # 尝试多种代码块标记格式解析
    # 1. 处理<code></code>格式
    start = "<code>"
    end = "</code>"
    while start in codes and end in codes:
        start_index = codes.index(start)
        end_index = codes.index(end)
        code = codes[start_index + len(start): end_index]
        codes = codes[end_index+len(end):]
        lines = code.strip().split('\n')
        for line in lines:
            parsed = line.strip().split(';')
            for x in parsed:
                if len(x)!=0 and x[-1]==')':
                    apis.append(x.strip())
    # 2. 处理<code></code格式(缺少>的情况)
    if len(apis)==0:
        start = "<code>"
        end = "</code"
        while start in codes and end in codes:
            start_index = codes.index(start)
            end_index = codes.index(end)
            code = codes[start_index + len(start): end_index]
            codes = codes[end_index + len(end):]
            lines = code.strip().split('\n')
            for line in lines:
                parsed = line.strip().split(';')
                for x in parsed:
                    if len(x) != 0 and x[-1] == ')':
                        apis.append(x.strip())
    # 3. 处理```scss```格式(使用rindex)
    if len(apis)==0:
        start="```scss"
        end="```"
        while start in codes and end in codes:
            start_index = codes.index(start)
            end_index = codes.rindex(end)
            code = codes[start_index + len(start): end_index]
            codes = codes[end_index + len(end):]
            lines = code.strip().split('\n')
            for line in lines:
                parsed = line.strip().split(';')
                for x in parsed:
                    if len(x) != 0 and x[-1] == ')':
                        apis.append(x.strip())
    # 4. 处理```scss```格式(使用index)
    if len(apis)==0:
        start="```scss"
        end="```"
        while start in codes and end in codes:
            start_index = codes.index(start)
            end_index = codes.index("```", start_index + len(start))
            code = codes[start_index + len(start): end_index]
            codes = codes[end_index + len(end):]
            lines = code.strip().split('\n')
            for line in lines:
                parsed = line.strip().split(';')
                for x in parsed:
                    if len(x) != 0 and x[-1] == ')':
                        apis.append(x.strip())
    # 5. 处理```python```格式
    if len(apis)==0:
        start="```python"
        end="```"
        while start in codes and end in codes:
            start_index = codes.index(start)
            end_index = codes.index("```", start_index + len(start))
            code = codes[start_index + len(start): end_index]
            codes = codes[end_index + len(end):]
            lines = code.strip().split('\n')
            for line in lines:
                parsed = line.strip().split(';')
                for x in parsed:
                    if len(x) != 0 and x[-1] == ')':
                        apis.append(x.strip())
    # 6. 处理普通```格式
    if len(apis)==0:
        start="```"
        end="```"
        while start in codes and end in codes:
            start_index = codes.index(start)
            end_index = codes.index(end, start_index + len(start))
            code = codes[start_index + len(start): end_index]
            codes = codes[end_index + len(end):]
            lines = code.strip().split('\n')
            for line in lines:
                parsed = line.strip().split(';')
                for x in parsed:
                    if len(x) != 0 and x[-1] == ')':
                        apis.append(x.strip())

    return apis

# 根据参数生成实验名称
def prepare_exp_name(args):
    """根据参数生成实验名称"""
    # 如果指定了exp_name，直接返回
    if hasattr(args, 'exp_name') and args.exp_name:
        return args.exp_name
        
    # 否则生成默认的实验名称
    name = ""
    if args.tf:
        name += "tf_"
    if args.sess:
        name += "sess_"
    name += args.model + "_"
    if args.planning:
        name += "p"
    if args.api_selection:
        name += f"a{args.api_topk}"
    if args.content_selection:
        name += "c"
    if not (args.planning or args.api_selection or args.content_selection):
        name += "n"
    return name

# 获取文本的token数量
def get_tokens(text):
    tokenizer = MosesTokenizer()
    tokens = tokenizer.tokenize(text)
    return len(tokens)

# 计算API调用的数量
def calc_api_cost(path):
    all_apis = api_doc.get_all_APIs()
    api_names = [x.name for x in all_apis]
    line = open(path, "r").read()
    lines = line.split('\n')
    cnt = 0
    for l in lines:
        l = l.strip(';')
        if l.endswith(')') and l.split('(')[0] in api_names:
            cnt += 1
    return cnt

# 检查prompt的token数是否超出模型限制
def check_token(model, prompt):
    # 设置不同模型的token限制
    if model == 'gpt4':
        max_token_limit = 8191
    elif model == 'text3':
        max_token_limit = 3095
    elif 'Llama' in model:
        max_token_limit = 2800
    elif 'WizardLM':
        max_token_limit= 1200
    else:
        max_token_limit = 4095
    # 选择合适的编码器
    if model == 'text3':
        encoding_model = tiktoken.get_encoding('p50k_base')
    else:
        encoding_model = tiktoken.get_encoding('cl100k_base')
    num_tokens = len(encoding_model.encode(prompt))
    exceeded = num_tokens - max_token_limit
    return exceeded if exceeded > 0 else 0

# 截断文本到指定token数
def get_token(text, trunc_num, model):
    if model == 'text3':
        encoding_model = tiktoken.get_encoding('p50k_base')
    else:
        encoding_model = tiktoken.get_encoding('cl100k_base')
    encoded = encoding_model.encode(text)[:-trunc_num]
    truncated = encoding_model.decode(encoded)
    return truncated

# 检查断点续传
def checkpoint(mode,args,idx,step):
    if not args.resume: 
        return 0
    # 检查数据准备阶段
    if mode == 'prepare' and args.prepare:
        if args.api_lack:
            if os.path.exists(os.path.join(args.save_path,str(idx),str(step),"lack_after_label.docx")):
                print(f"Prepare data Exists {idx}/{step}!")
                return 1
        else:
            if os.path.exists(os.path.join(args.save_path,str(idx),str(step),"after_label.docx")):
                print(f"Prepare data Exists {idx}/{step}!")
                return 1    
    # 检查会话模式
    if mode == 'sess' and args.sess and os.path.exists(os.path.join(args.save_path,str(idx),str(step),f"{args.exp_name}_after_pred.docx")):
        print(f"Sess data Exists {idx}!")
        return 1
    # 检查转换模式
    if mode == 'tf' and args.tf and os.path.exists(os.path.join(args.save_path,str(idx),str(step),f"{args.exp_name}_after_pred.docx")):
        print(f"Tf data Exists {idx}!")
        return 1
    return 0

# 按数字顺序排序文件列表
def sorted_list(path):
    """按数字顺序排序目录中的文件
    
    Args:
        path: 目录路径
    Returns:
        排序后的文件列表
    """
    def get_number(filename):
        try:
            # 尝试从文件名中提取数字
            return int(filename.split('_')[-1].split('.')[0])
        except (ValueError, IndexError):
            # 如果无法提取数字，返回-1或其他默认值
            return -1
    
    # 使用修改后的排序键函数
    return sorted(os.listdir(path), key=get_number)

# 解析训练数据JSON文件
def parse_train_json(path):
    """解析训练数据的JSON文件
    
    Args:
        path: JSON文件路径
        
    Returns:
        turns: 包含每个训练样本信息的列表
    """
    turns = []
    try:
        with open(path, 'r', encoding='utf-8') as f:
            content = f.read().strip()
            if not content:  # 如果文件为空
                return turns
                
            try:
                # 尝试解析为JSON数组
                data = json.loads(content)
                if isinstance(data, list):
                    # 处理JSON数组
                    for item in data:
                        turns.append([
                            item.get('Turn'),
                            item.get('User instruction'),
                            item.get('Feasible API sequence'),
                            item.get('Base File'),
                            item.get('Label File'),
                            item.get('API Lack Base File'),
                            item.get('API Lack Label File')
                        ])
                elif isinstance(data, dict):
                    # 处理单个JSON对象
                    turns.append([
                        data.get('Turn'),
                        data.get('User instruction'),
                        data.get('Feasible API sequence'),
                        data.get('Base File'),
                        data.get('Label File'),
                        data.get('API Lack Base File'),
                        data.get('API Lack Label File')
                    ])
            except json.JSONDecodeError as e:
                print(f"Error parsing JSON in {path}: {str(e)}")
                
    except Exception as e:
        print(f"Error reading file {path}: {str(e)}")
        
    return turns

# 解析测试数据JSON文件
def parse_test_json(path):
    """解析测试数据的JSON文件
    
    Args:
        path: JSON文件路径
        
    Returns:
        turns: 包含每个测试样本信息的列表
    """
    turns = []
    try:
        with open(path, 'r', encoding='utf-8') as f:
            content = f.read().strip()
            if not content:  # 如果文件为空
                return turns
                
            try:
                # 尝试解析为JSON数组或对象
                data = json.loads(content)
                if isinstance(data, list):
                    # 处理JSON数组
                    for item in data:
                        turns.append([
                            item.get('Turn'),
                            item.get('User instruction'),
                            item.get('Feasible API sequence'),
                            item.get('Reply', ''),  # 使用空字符串作为默认值
                            item.get('Pred API sequence', []),  # 使用空列表作为默认值
                            item.get('Pred File', ''),
                            item.get('Label File', ''),
                            item.get('Prompt File', '')
                        ])
                elif isinstance(data, dict):
                    # 处理单个JSON对象
                    turns.append([
                        data.get('Turn'),
                        data.get('User instruction'),
                        data.get('Feasible API sequence'),
                        data.get('Reply', ''),
                        data.get('Pred API sequence', []),
                        data.get('Pred File', ''),
                        data.get('Label File', ''),
                        data.get('Prompt File', '')
                    ])
            except json.JSONDecodeError as e:
                print(f"Error parsing JSON in {path}: {str(e)}")
                
    except Exception as e:
        print(f"Error reading file {path}: {str(e)}")
        
    return turns

# 主函数用于测试
if __name__ == '__main__':
    api_text = """<code>
# Set the height of pictures in paragraph 1 as 5
move_to_paragraph(1);
choose_picture(0);
set_height(5);
choose_picture(1);
set_height(5);
choose_picture(2);
set_height(5);

# Set the heading text as 'The Art of Young' and underline it
move_to_heading(1);
delete_text();
insert_text('The Art of Young');
set_font_underline();
</code>"""
    parse_api(api_text)