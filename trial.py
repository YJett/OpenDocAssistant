import json

# 定义JSON文件路径
json_path = 'Word_test_input/Create_new_docs/session_0.json'

def parse_train_json(path):
    """解析训练数据的JSON文件
    
    Args:
        path: JSON文件路径
        
    Returns:
        turns: 包含每个训练样本信息的列表,每个样本包含:
            - turn_id: 对话轮次ID
            - instruction: 用户指令
            - label_api: 标准API序列
            - base_ppt_path: 基础PPT文件路径
            - label_ppt_path: 标准结果PPT路径
            - api_lack_label_ppt_path: API缺失情况下的标准PPT路径
    """
    turns = []
    with open(path, 'r', encoding='utf-8') as file:
        content = file.read()
        # 确保内容不为空
        if not content.strip():
            return turns
            
        # 尝试解析整个文件作为一个JSON对象
        try:
            data = json.loads(content)
            if isinstance(data, list):
                # 如果是JSON数组
                for item in data:
                    turn_id = item.get('Turn')
                    instruction = item.get('User instruction')
                    label_api = item.get('Feasible API sequence')
                    base_ppt_path = item.get('Base File')
                    label_ppt_path = item.get('Label File')
                    api_lack_label_ppt_path = item.get('API Lack Label File')
                    turns.append([turn_id, instruction, label_api, base_ppt_path, label_ppt_path, api_lack_label_ppt_path])
            else:
                # 单个JSON对象
                turn_id = data.get('Turn')
                instruction = data.get('User instruction')
                label_api = data.get('Feasible API sequence')
                base_ppt_path = data.get('Base File')
                label_ppt_path = data.get('Label File')
                api_lack_label_ppt_path = data.get('API Lack Label File')
                turns.append([turn_id, instruction, label_api, base_ppt_path, label_ppt_path, api_lack_label_ppt_path])
        except json.JSONDecodeError:
            # 尝试按行解析JSON
            for line in content.splitlines():
                if line.strip():
                    try:
                        data = json.loads(line)
                        turn_id = data.get('Turn')
                        instruction = data.get('User instruction')
                        label_api = data.get('Feasible API sequence')
                        base_ppt_path = data.get('Base File')
                        label_ppt_path = data.get('Label File')
                        api_lack_label_ppt_path = data.get('API Lack Label File')
                        turns.append([turn_id, instruction, label_api, base_ppt_path, label_ppt_path, api_lack_label_ppt_path])
                    except json.JSONDecodeError:
                        continue
    return turns

def parse_test_json(path):
    """解析测试数据的JSON文件
    
    Args:
        path: JSON文件路径
        
    Returns:
        turns: 包含每个测试样本信息的列表,每个样本包含:
            - turn_id: 对话轮次ID
            - instruction: 用户指令
            - label_api: 标准API序列
            - reply: 模型回复
            - pred_api: 预测的API序列
            - pred_ppt_path: 预测生成的PPT路径
            - label_ppt_path: 标准结果PPT路径
            - prompt_path: 提示词文件路径
    """
    turns = []
    with open(path, 'r', encoding='utf-8') as file:
        content = file.read()
        # 确保内容不为空
        if not content.strip():
            return turns
            
        # 尝试解析整个文件作为一个JSON对象
        try:
            data = json.loads(content)
            if isinstance(data, list):
                # 如果是JSON数组
                for item in data:
                    turn_id = item.get('Turn')
                    instruction = item.get('User instruction')
                    label_api = item.get('Feasible API sequence')
                    reply = item.get('Reply')
                    pred_api = item.get('Pred API sequence')
                    pred_ppt_path = item.get('Pred File')
                    label_ppt_path = item.get('Label File')
                    prompt_path = item.get('Prompt File')
                    turns.append([turn_id, instruction, label_api, reply, pred_api, pred_ppt_path, label_ppt_path, prompt_path])
            else:
                # 单个JSON对象
                turn_id = data.get('Turn')
                instruction = data.get('User instruction')
                label_api = data.get('Feasible API sequence')
                reply = data.get('Reply')
                pred_api = data.get('Pred API sequence')
                pred_ppt_path = data.get('Pred File')
                label_ppt_path = data.get('Label File')
                prompt_path = data.get('Prompt File')
                turns.append([turn_id, instruction, label_api, reply, pred_api, pred_ppt_path, label_ppt_path, prompt_path])
        except json.JSONDecodeError:
            # 尝试按行解析JSON
            for line in content.splitlines():
                if line.strip():
                    try:
                        data = json.loads(line)
                        turn_id = data.get('Turn')
                        instruction = data.get('User instruction')
                        label_api = data.get('Feasible API sequence')
                        reply = data.get('Reply')
                        pred_api = data.get('Pred API sequence')
                        pred_ppt_path = data.get('Pred File')
                        label_ppt_path = data.get('Label File')
                        prompt_path = data.get('Prompt File')
                        turns.append([turn_id, instruction, label_api, reply, pred_api, pred_ppt_path, label_ppt_path, prompt_path])
                    except json.JSONDecodeError:
                        continue
    return turns

# 测试代码:解析训练数据并打印第一个样本的信息
turns = parse_train_json(json_path)
print(len(turns))
for x in turns[0]:
    print(x)