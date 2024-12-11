from src import utils
from tqdm import tqdm
from src import (
    word_executor,
    modeling,
    evaluate,
    content_selection,
    word_reader,
    openai_api,
    prompt_factor,
    dataset,
    api_selection
)
import argparse
import os
from tqdm import tqdm
import jsonlines

def regenerate(word_assistant, args):
    """重新生成文档"""
    # 设置数据集名称
    set_name = 'Create_new_docs' if args.dataset == 'short' else 'Edit_Word_template'
    
    # 创建输出目录
    utils.makedir(f'Word_Pred_File1/{set_name}')
    utils.makedir(f'Word_test_output1/{set_name}')
    
    # 遍历测试输出文件
    for session_path in utils.sorted_list(f'Word_test_output/{set_name}'):
        if not session_path.startswith(args.exp_name):
            continue
            
        sess_id = int(session_path.split('_')[-1].replace('.json',''))
        session = utils.parse_test_json(f'Word_test_output/{set_name}/{session_path}')
        chat_history = []
        
        for turn_id, turn in tqdm(enumerate(session)):
            print(f"{sess_id}/{turn_id}")
            turn_id, instruction, label_api, reply, pred_api, pred_doc_path, label_doc_path, prompt_path = turn

            if args.tf:
                base_doc_path = f"Word_Base_File/{set_name}/{sess_id}_{turn_id}.docx"
                api_lack_base_doc_path = f"Word_Base_File/{set_name}_API_lack/{sess_id}_{turn_id}.docx"
                
                if args.api_lack:
                    word_assistant.load_docx(api_lack_base_doc_path)
                    label_file = api_lack_label_doc_path
                else:
                    word_assistant.load_docx(base_doc_path)
                    label_file = label_doc_path

                apis = pred_api
                word_assistant.api_executor(apis, test=True)
                word_executor.save_docx(f'Word_Pred_File1/{set_name}/{args.exp_name}_{sess_id}_{turn_id}.docx')
                
                with jsonlines.open(f"Word_test_output1/{set_name}/{args.exp_name}_session_{sess_id}.json", mode='a') as writer:
                    data = {
                        'Turn': turn_id,
                        'User instruction': instruction,
                        'Feasible API sequence': label_api,
                        'Reply': reply,
                        'Pred API sequence': apis,
                        'Pred File': f'Word_Pred_File1/{set_name}/{args.exp_name}_{sess_id}_{turn_id}.docx',
                        'Label File': label_file,
                        'Prompt File': f'Word_Prompt_File/{set_name}/{args.exp_name}_{sess_id}_{turn_id}.txt'
                    }
                    writer.write(data)
            
            elif args.sess:
                base_doc_path = f"Word_Base_File/{set_name}/{sess_id}_{turn_id}.docx"
                api_lack_base_doc_path = f"Word_Base_File/{set_name}_API_lack/{sess_id}_{turn_id}.docx"
                
                if args.api_lack:
                    if turn_id == 0:
                        word_assistant.load_docx(api_lack_base_doc_path)
                    label_file = api_lack_label_doc_path
                else:
                    if turn_id == 0:
                        word_assistant.load_docx(base_doc_path)
                    label_file = label_doc_path

                apis = pred_api
                word_assistant.api_executor(apis, test=True)
                word_executor.save_docx(f'Word_Pred_File1/{set_name}/{args.exp_name}_{sess_id}_{turn_id}.docx')

                with jsonlines.open(f"Word_test_output1/{set_name}/{args.exp_name}_session_{sess_id}.json", mode='a') as writer:
                    data = {
                        'Turn': turn_id,
                        'User instruction': instruction,
                        'Feasible API sequence': label_api,
                        'Reply': reply,
                        'Pred API sequence': apis,
                        'Pred File': f'Word_Pred_File1/{set_name}/{args.exp_name}_{sess_id}_{turn_id}.docx',
                        'Label File': label_file,
                        'Prompt File': f'Word_Prompt_File/{set_name}/{args.exp_name}_{sess_id}_{turn_id}.txt'
                    }
                    writer.write(data)

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    # Word assistant
    parser.add_argument("--data_path", default="test", type=str,
                        help="The data path to load the instructions")
    parser.add_argument("--dataset", default="short", type=str,
                        help="short/long")
    parser.add_argument("--model_id", default="None", type=str,
                        help="Model identifier")
    parser.add_argument("--user_path", default='./Word_assistant_json/', type=str,
                        help="Path to user data")
    parser.add_argument("--save_path", default="test_docx_data", type=str,
                        help="Path to save intermediate docs")
    
    # mode
    parser.add_argument("--prepare", action='store_true',
                        help='Whether to prepare the data for the model')
    parser.add_argument("--eval", action='store_true',
                        help='Whether to evaluate the docx files generated by the model')
    parser.add_argument("--test", action='store_true',
                        help='Whether to test on the instruction data loaded from data_path')
    parser.add_argument("--tf", action='store_true',
                        help='Whether to use teacher forcing mode')
    parser.add_argument("--sess", action='store_true',
                        help='Whether to test from session level')
    parser.add_argument("--resume", action='store_true',
                        help='Whether to continue generation from the last unfinished instruction')
    
    # modeling
    parser.add_argument("--model", default="turbo", type=str,
                        help="turbo/gpt4/text3") 
    parser.add_argument("--planning", action='store_true',
                        help="Whether to apply the planning module") 
    parser.add_argument("--api_selection", action='store_true',
                        help="Whether to apply the api selection module") 
    parser.add_argument("--api_topk", default=10, type=int,
                        help="How many apis to retrieve from the api pool") 
    parser.add_argument("--content_selection", action='store_true',
                        help="Whether to apply the content selection module") 
    
    # api update/lack
    parser.add_argument("--api_lack", action='store_true',
                        help='Whether to test in the api lack setting')
    parser.add_argument("--api_update", action='store_true',
                        help='Whether to test in the api update setting')
    parser.add_argument("--second", action='store_true',
                        help='Second test')
    parser.add_argument("--robust", action='store_true',
                        help='Whether to test in robust data')
    parser.add_argument("--robust_num", default=0, type=int,
                        help="Which robusted data") 
    parser.add_argument("--noisy", action='store_true',
                        help='Whether to test in noisy data')

    args = parser.parse_args()

    args.exp_name = utils.prepare_exp_name(args)
    args.save_path = os.path.join(args.save_path, args.dataset)
    api_selection.prepare_embedding(args)
    word_assistant = modeling.word_assistant(args)

    regenerate(word_assistant, args)