from src import word_executor, word_reader, openai_api, prompt_factor, dataset, api_selection, utils, modeling, evaluate, content_selection
import argparse
import os
from tqdm import tqdm
import jsonlines
import logging

# 设置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def prepare_data(word_assistant, args):
    """准备数据集"""
    instructions, labels = dataset.load_data(args.user_path+args.data_path, args.dataset, args)
    logger.info(f"#Dialogues: {len(instructions)}")
    
    for idx, dialogue in enumerate(instructions): 
        if idx >= len(labels):
            logger.error(f"No labels found for dialogue {idx}")
            continue
            
        if args.dataset == 'long':
            word_assistant.load_docx(os.path.join(args.user_path+'long_docs',f'{idx}.docx'))
        else:
            word_assistant.load_docx(None)
        
        set_name = 'Edit_Word_template' if args.dataset == 'long' else 'Create_new_docs'
        
        # 创建必要的目录
        if args.api_lack:
            utils.makedir(args.user_path+f"Word_Base_File/{set_name}_API_lack/")
            utils.makedir(args.user_path+f"Word_Label_File/{set_name}_API_lack/")
        else:
            utils.makedir(args.user_path+f"Word_Base_File/{set_name}/")
            utils.makedir(args.user_path+f"Word_Label_File/{set_name}/")

        for step, instruction in enumerate(dialogue):
            if step >= len(labels[idx]):
                logger.error(f"No labels found for dialogue {idx} step {step}")
                continue
                
            instruction = instruction.split("##")[0]
            label_apis = utils.merge_list(labels[idx][:step])

            if args.dataset == 'long':
                word_assistant.load_docx(os.path.join(args.user_path+'long_docs',f'{idx}.docx'))
            else:
                word_assistant.load_docx(None)
            
            try:
                word_assistant.api_executor(label_apis, test=False)
                
                # 保存基础文件
                if args.api_lack:
                    word_executor.save_word(args.user_path+f"Word_Base_File/{set_name}_API_lack/{idx}_{step}.docx")
                else:
                    word_executor.save_word(args.user_path+f"Word_Base_File/{set_name}/{idx}_{step}.docx")

                word_assistant.api_executor(labels[idx][step], test=False)

                # 保存标签文件
                if args.api_lack:
                    word_executor.save_word(args.user_path+f"Word_Label_File/{set_name}_API_lack/{idx}_{step}.docx")
                else:
                    word_executor.save_word(args.user_path+f"Word_Label_File/{set_name}/{idx}_{step}.docx")
                logger.info(f"{idx}/{step} done!")
                
            except Exception as e:
                logger.error(f"Error processing dialogue {idx} step {step}: {e}")
                continue

def test(word_assistant, args):
    """测试模型"""
    set_name = 'Create_new_docs' if args.dataset == 'short' else 'Edit_Word_template'
    utils.makedir(args.user_path+f'Word_Pred_File/{set_name}')
    utils.makedir(args.user_path+f'Word_Prompt_File/{set_name}')
    
    for sess_id, session_path in enumerate(utils.sorted_list(args.user_path+f'Word_test_input/{set_name}')):
        try:
            # 每个新会话开始时创建新的 Word 文档
            doc = word_assistant.load_docx(None)
            if not doc:
                logger.error(f"Failed to create document for session {sess_id}")
                continue
                
            session = utils.parse_train_json(args.user_path+f'Word_test_input/{set_name}/{session_path}')
            chat_history = []
            for turn_id, turn in tqdm(enumerate(session)):
                logger.info(f"{sess_id}/{turn_id}")
                
                if args.resume:
                    if args.tf and os.path.exists(args.user_path+f'Word_Pred_File/{set_name}/{args.exp_name}_{sess_id}_{turn_id}.docx'):
                        logger.info('Exists!')
                        continue 
                    if args.sess and os.path.exists(args.user_path+f'Word_Pred_File/{set_name}/{args.exp_name}_{sess_id}_{len(session)-1}.docx'):
                        logger.info('Exists!')
                        continue 
                        
                turn_id, instruction, label_api, base_doc_path, label_doc_path, api_lack_base_doc_path, api_lack_label_doc_path = turn
                if turn_id == 0 and args.sess:
                    if args.api_lack:
                        word_assistant.load_docx(args.user_path+api_lack_base_doc_path)
                        label_file = api_lack_label_doc_path
                    else:
                        word_assistant.load_docx(args.user_path+base_doc_path)
                        label_file = label_doc_path
                        
                splitted_instruction = instruction.split("##")[0]
                if args.tf:
                    if args.api_lack:
                        word_assistant.load_docx(args.user_path+api_lack_base_doc_path)
                        label_file = api_lack_label_doc_path
                    else:
                        word_assistant.load_docx(args.user_path+base_doc_path)
                        label_file = label_doc_path
                        
                    word_assistant.load_chat_history([x[0] for x in chat_history],[x[1].strip(';').split(';') for x in chat_history])
                    prompt, reply = word_assistant.chat(splitted_instruction, doc_path=args.user_path+base_doc_path, verbose=False)
                    apis = utils.parse_api(reply)
                    word_assistant.api_executor(apis, test=True)
                    
                    word_executor.save_word(args.user_path+f'Word_Pred_File/{set_name}/{args.exp_name}_{sess_id}_{turn_id}.docx')
                    utils.write_lines([prompt],args.user_path+f'Word_Prompt_File/{set_name}/{args.exp_name}_{sess_id}_{turn_id}.txt')
                    
                    utils.makedir(f"Word_test_output/{set_name}")
                    with jsonlines.open(args.user_path+f"Word_test_output/{set_name}/{args.exp_name}_session_{sess_id}.json", mode='a') as writer:
                        data = {
                            'Turn': turn_id,
                            'User instruction': instruction,
                            'Feasible API sequence': label_api,
                            'Reply': reply,
                            'Pred API sequence': apis,
                            'Pred File': f'Word_Pred_File/{set_name}/{args.exp_name}_{sess_id}_{turn_id}.docx',
                            'Label File': label_file,
                            'Prompt File': f'Word_Prompt_File/{set_name}/{args.exp_name}_{sess_id}_{turn_id}.txt'
                        }
                        writer.write(data)
                    chat_history.append([splitted_instruction, label_api])
                
                elif args.sess:
                    prompt, reply = word_assistant.chat_v2(instruction, doc_path=None, verbose=False)
                    # apis = utils.parse_api(reply)
                    apis = reply
                    word_assistant.api_executor(apis, test=True)
                    word_executor.save_word(args.user_path+f'Word_Pred_File/{set_name}/{args.exp_name}_{sess_id}_{turn_id}.docx')
                    utils.write_lines([prompt],args.user_path+f'Word_Prompt_File/{set_name}/{args.exp_name}_{sess_id}_{turn_id}.txt')
                    
                    with jsonlines.open(args.user_path+f"Word_test_output/{set_name}/{args.exp_name}_session_{sess_id}.json", mode='a') as writer:
                        data = {
                            'Turn': turn_id,
                            'User instruction': instruction,
                            'Feasible API sequence': label_api,
                            'Reply': reply,
                            'Pred API sequence': apis,
                            'Pred File': f'Word_Pred_File/{set_name}/{args.exp_name}_{sess_id}_{turn_id}.docx',
                            'Label File': label_file,
                            'Prompt File': f'Word_Prompt_File/{set_name}/{args.exp_name}_{sess_id}_{turn_id}.txt'
                        }
                        writer.write(data)
        except Exception as e:
            logger.error(f"Error processing session {sess_id}: {e}")
            continue

def test_api_selection(word_assistant):
    """测试API选择模块
    
    用于验证API选择的准确性，记录未被正确选择的API
    """
    instructions, labels = dataset.load_data(args.data_path, args.dataset)
    f = open(f'test_system/api_selection_{args.api_topk}_{args.dataset}.txt', 'a+')
    cnt = 0
    
    for idx, dialogue in tqdm(enumerate(instructions)):
        for step, instruction in enumerate(dialogue):
            label_apis = labels[idx][step]
            instruction = instruction.split("##")[0]
            
            # 获取API选择结果
            selected_apis = word_assistant.api_selector(instruction)
            selected_apis = [x.name for x in selected_apis]
            
            # 检查每个标签API是否被正确选择
            for api in label_apis:
                api_name = api.split('(')[0]
                # 跳过位置相关的API
                if ('align' in api_name) or ('set_position' in api_name):
                    continue
                    
                if api_name not in selected_apis:
                    f.write(f'Case {idx}/{step}\n')
                    f.write(f'Instruction: {instruction}\n')
                    f.write(f'Missing API: {api_name}\n')
                    f.write(f'Selected APIs: {selected_apis}\n\n')
                    f.flush()
                    cnt += 1
                    
    logger.info(f"Total missing API cases: {cnt}")
    return cnt

def test_content_selection(word_assistant):
    """测试内容选择模块
    
    用于验证内容选择的效果，记录选择结果
    """
    instructions, labels = dataset.load_data(args.data_path, args.dataset)
    f = open(f'test_system/content_selection_{args.dataset}.txt', 'a+')
    
    for idx, dialogue in tqdm(enumerate(instructions)):
        for step, instruction in enumerate(dialogue):
            instruction = instruction.split("##")[0]
            
            # 生成内容选择提示
            prompt = prompt_factor.Word_content_selection_prompt.format(
                instruction=instruction
            )
            
            # 获取内容选择结果
            try:
                reply = openai_api.query_azure_openai(prompt, model='turbo')
                
                f.write(f'Case {idx}/{step}\n')
                f.write(f'Instruction: {instruction}\n')
                f.write(f'Selected content:\n{reply}\n\n')
                f.flush()
                
            except Exception as e:
                logger.error(f"Error in content selection for case {idx}/{step}: {e}")
                continue

def test_planning(word_assistant):
    """测试规划模块
    
    用于验证指令分解和规划的效果
    """
    instructions, labels = dataset.load_data(args.data_path, args.dataset)
    f = open(f'test_system/planning_{args.dataset}.txt', 'a+')
    
    for idx, dialogue in tqdm(enumerate(instructions)):
        for step, instruction in enumerate(dialogue):
            instruction = instruction.split("##")[0]
            
            try:
                # 获取规划结果
                planned_steps = word_assistant.planner(instruction)
                
                f.write(f'Case {idx}/{step}\n')
                f.write(f'Instruction: {instruction}\n')
                f.write('Planned steps:\n')
                for i, step in enumerate(planned_steps, 1):
                    f.write(f'{i}. {step}\n')
                f.write('\n')
                f.flush()
                
            except Exception as e:
                logger.error(f"Error in planning for case {idx}/{step}: {e}")
                continue

if __name__ == "__main__":
    # 1. 首先运行 API 选择测试
    from src.api_selection import run_api_selection_test
    # run_api_selection_test()
    # 2. 然后继续原来的主程序逻辑
    parser = argparse.ArgumentParser()
    parser.add_argument("--skip_api_test", action='store_true',
                    help="Skip API selection test")

    parser.add_argument("--data_path", default="test", type=str,
                        help="The data path to load the instructions")
    parser.add_argument("--dataset", default="short", type=str,
                        help="short/long")
    parser.add_argument("--model_id", default="None", type=str,
                        help="Model identifier")
    parser.add_argument("--user_path", default='./', type=str,
                        help="Path to user data")
    parser.add_argument("--save_path", default="test_docx_data", type=str,
                        help="Path to save intermediate docs")
    
    # mode
    parser.add_argument("--prepare", default=False, action='store_true',
                        help='Whether to prepare the data for the model')
    parser.add_argument("--eval", default=False, action='store_true',
                        help='Whether to evaluate the docx files generated by the model')
    parser.add_argument("--test", default=False, action='store_true',
                        help='Whether to test on the instruction data loaded from data_path')
    parser.add_argument("--tf", default=False, action='store_true',
                        help='Whether to use teacher forcing mode')
    parser.add_argument("--sess", default=False, action='store_true',
                        help='Whether to test from session level')
    parser.add_argument("--resume", default=False, action='store_true',
                        help='Whether to continue generation from the last unfinished instruction')
    
    # modeling
    parser.add_argument("--model", default="turbo", type=str,
                        help="turbo/gpt4/text3") 
    parser.add_argument("--planning", default=False, action='store_true',
                        help="Whether to apply the planning module") 
    parser.add_argument("--api_selection", default=False, action='store_true',
                        help="Whether to apply the api selection module") 
    parser.add_argument("--api_topk", default=10, type=int,
                        help="How many apis to retrieve from the api pool") 
    parser.add_argument("--content_selection", default=False, action='store_true',
                        help="Whether to apply the content selection module") 
    
    # api update/lack
    parser.add_argument("--api_lack", default=False, action='store_true',
                        help='Whether to test in the api lack setting')
    parser.add_argument("--api_update", default=False, action='store_true',
                        help='Whether to test in the api update setting')
    parser.add_argument("--second", default=False, action='store_true',
                        help='Second test')

    parser.add_argument("--robust", default=False, action='store_true',
                        help='Whether to test in robust data')
    parser.add_argument("--robust_num", default=0, type=int,
                        help="Which robusted data") 
    parser.add_argument("--noisy", default=False, action='store_true',
                        help='Whether to test in noisy data')
    
    args = parser.parse_args()

    # 设置日志
    logging.basicConfig(level=logging.INFO)
    logger = logging.getLogger(__name__)

    try:
        # 创建必要的目录
        os.makedirs("data", exist_ok=True)
        os.makedirs("test_docx_data", exist_ok=True)
        
        args.exp_name = utils.prepare_exp_name(args)
        args.save_path = os.path.join(args.save_path, args.dataset)
        
        # 初始化 word assistant
        api_selection.prepare_embedding(args)
        word_assistant = modeling.word_assistant(args)

        if args.prepare:
            logger.info("Starting data preparation...")
            prepare_data(word_assistant, args)
            logger.info("Data preparation completed")
            exit(0)
            
        if args.test:
            logger.info("Starting testing...")
            test(word_assistant, args)
            logger.info("Testing completed")
            exit(0)
            
        if args.eval:
            logger.info("Starting evaluation...")
            evaluate.eval(args)
            logger.info("Evaluation completed")
            exit(0)
        if not args.skip_api_test:
            from src.api_selection import run_api_selection_test
            run_api_selection_test()
            
    except Exception as e:
        logger.error(f"Error in main: {e}")
        raise

