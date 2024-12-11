from src import utils
from tqdm import tqdm
from src import api_doc

# 获取所有Word API列表
all_apis = api_doc.document_APIs + api_doc.text_APIs + api_doc.table_APIs + api_doc.picture_APIs + api_doc.chart_APIs + api_doc.basic_APIs
all_apis = [x.name for x in all_apis]
set_name = 'Create_new_doc'  # 修改为Word相关的目录名

def check(pred):
    """
    检查预测的API是否都在已定义的API列表中
    
    Args:
        pred: 预测的API列表
    Returns:
        bool: 所有API都合法返回1，否则返回0
    """
    for x in pred:
        strx = x.split('(')[0]
        if not strx in all_apis:
            return 0
    return 1

def check_trigger(pred, label):
    """
    检查特定API的触发条件
    
    Args:
        pred: 预测的API列表
        label: 标准答案API列表
    Returns:
        bool: 满足触发条件返回1，否则返回0
    """
    # 检查表格相关操作的触发
    if 'add_table_row' in '\n'.join(pred) and 'set_cell_text' in label:
        return 1
    if 'add_table_row' in label and 'set_cell_text' in '\n'.join(pred):
        return 1
    # 检查段落相关操作的触发
    if 'add_paragraph' in '\n'.join(pred) and 'add_text' in label:
        return 1
    if 'add_paragraph' in label and 'add_text' in '\n'.join(pred):
        return 1
    return 0

# 遍历不同模型的结果进行检查
for model in ['Baichuan-13b', 'Baichuan-13B-Chat', 'Llama-2-13b-chat-hf', 'WizardLM-13b']:
    cnt = 0
    print(model)
    for sess_id, session_path in enumerate(utils.sorted_list(f'Word_test_output/{set_name}')):
        if not session_path.startswith(f'tf_lack_{model}'):
            continue
            
        session = utils.parse_test_json(f'Word_test_output/{set_name}/{session_path}')
        
        for turn_id, turn in enumerate(session):
            turn_id, instruction, label_api, reply, pred_api, pred_doc_path, label_doc_path, prompt_path = turn
            
            # 检查表格和段落相关的API调用
            if ('add_table_row' in '\n'.join(pred_api) or 'set_cell_text' in '\n'.join(pred_api) or
                'add_paragraph' in '\n'.join(pred_api) or 'add_text' in '\n'.join(pred_api)):
                
                # 跳过非法API调用
                if not check(pred_api):
                    continue
                    
                # 跳过不满足触发条件的调用
                if not check_trigger(pred_api, label_api):
                    continue
                    
                print(pred_api)
                print(label_api)
                print('')
                cnt += 1
    print(cnt)

"""
主要修改说明：
1. 修改了API列表，使用Word相关的API
2. 修改了目录名从PPT相关改为Word相关
3. 修改了check_trigger函数中的触发条件，增加了Word文档特有的操作组合：
   - 表格操作：add_table_row 和 set_cell_text
   - 段落操作：add_paragraph 和 add_text
4. 保持了原有的统计和检查逻辑
"""