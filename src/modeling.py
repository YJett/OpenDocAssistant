from asyncio.log import logger
from docx import Document
from src import (
    word_executor,
    word_reader,
    openai_api,
    prompt_factor,
    api_selection,
    utils,
    api_doc
)

class word_assistant(object):
    def __init__(self, args=None):
        self.chat_history = []
        self.args = args
        self.planning = args.planning
        self.api_selection = args.api_selection
        self.content_selection = args.content_selection
        self.model = args.model
        self.model_id = args.model_id
        self.doc = None
        self.prompt = ""
        
        class AgentArgs:
            def __init__(self, api_update, api_lack, dataset, model, api_topk):
                self.api_update = api_update
                self.api_lack = api_lack
                self.dataset = dataset
                self.model = model
                self.api_topk = api_topk

        agent_args = {
            "api_update": False,
            "api_lack": False,
            "dataset": 'short',
            "model": "gpt-3.5-turbo",
            "api_topk": 10
        }
        self.agent, _ = api_selection.prepare_embedding(AgentArgs(**agent_args))

    def planner(self, instruction):
        if not self.planning:
            return [instruction]
        else:
            print('Planning...')
            planning_prompt = prompt_factor.query_decomposition_prompt.format(instruction)
            self.prompt += planning_prompt + "\n\n"
            planning_reply = openai_api.query_openai(planning_prompt, model=self.model).strip()
            decomposed = planning_reply.split('\n')
            decomposed = [d.replace('</d>', '') for d in decomposed if (d != '</d>') and (d != '<d>')]
            print(f"{instruction}->{decomposed}")
            return decomposed

    def api_selector(self, instruction):
        if not self.api_selection:
            all_apis = api_selection.get_all_apis(self.args)
            return all_apis
        else:
            selected_apis = api_selection.get_selected_apis(instruction, self.args)
            print('Selecting APIs...')
            print([x.name for x in selected_apis])
            return selected_apis

    def content_selector(self, word_path, instruction, args, word):
        content, prompt = word_reader.get_content_by_instructions(word_path, instruction, args, word)
        self.prompt += prompt + '\n\n'
        return content

    def api_executor(self, apis, test=False):
        print('Executing APIs...')
        error_info = word_executor.API_executor(apis, test=test, args=self.args)
        if error_info != "":
            print(error_info)
        self.doc = word_executor.get_word()

    def load_docx(self, path):
        """加载或创建新的Word文档
        
        Args:
            path: 文档路径,None表示创建新文档
            
        Returns:
            Document对象
        """
        try:
            word_executor.set_word(path)  # 如果path是None，则创建新文档
            if path is None:
                word_executor.create_docx()  # 创建新文档并添加空白段落
            self.doc = word_executor.get_word()
            self.chat_history = []  # 重置对话历史
            return self.doc
        except Exception as e:
            logger.error(f"Error loading document: {e}")
            return None

    def load_chat_history(self, instructions, labels):
        """加载对话历史
        
        Args:
            instructions: 指令列表
            labels: 标签列表
        Returns:
            对话历史列表
        """
        history = []
        for idx, instruction in enumerate(instructions):
            if self.args.api_lack:
                label = api_doc.api_lack_mask(labels[idx])
            else:
                label = labels[idx]
            label_str = ";\n".join(label)
            history += [
                f"¬User¬\n{instruction}",
                f"¬AI¬:\n<code>\n{label_str};\n</code>",
            ]
        self.chat_history = history
        return history

    def chat_v2(self, user_instruction, doc_path=None, verbose=False):
        self.prompt = ""
        reply_list = []
        api_list, prompt = self.agent.analyze_request_with_prompt(user_instruction)
        apis = [part.strip() for part in api_list.split(';') if part.strip()]
        return prompt, apis
        
    def chat(self, user_instruction, doc_path=None, verbose=False):
        self.prompt = ""
        instruction_list = self.planner(user_instruction)
        reply_list = []
        for instruction in instruction_list:
            if verbose:
                print('Executing instruction: ', instruction)

            selected_apis = self.api_selector(instruction)
            API_string = "\n".join(map(str, selected_apis))
            if verbose:
                print(f"== Selected APIs ==\n{API_string}\n\n")

            word_content = self.content_selector(doc_path, instruction, self.args, self.doc)
            if verbose:
                print(word_content)
            
            prompt = prompt_factor.get_instruction_to_API_code_prompt2(
                API_string,
                word_content,
                self.chat_history,
                instruction,
                True,
                0,
            )

            exceeded = utils.check_token(self.model, prompt)
            if exceeded != 0:
                print(f'Exceeded:{exceeded}')
                truncated_word_content = utils.get_token(word_content, exceeded, self.model)
                prompt = prompt_factor.get_instruction_to_API_code_prompt2(
                    API_string,
                    truncated_word_content,
                    self.chat_history,
                    instruction,
                    True,
                    0,
                )

                exceeded = utils.check_token(self.model, prompt)
                if exceeded != 0:
                    print(f'Exceeded:{exceeded}')
                    truncated_API_string = utils.get_token(API_string, exceeded, self.model)
                    prompt = prompt_factor.get_instruction_to_API_code_prompt2(
                        truncated_API_string,
                        truncated_word_content,
                        self.chat_history,
                        instruction,
                        True,
                        0,
                    )
            self.prompt += prompt + '\n\n'
            if verbose:
                print(f"== Prompt ==\n{prompt}\n\n")

            try:
                reply = openai_api.query_azure_openai(prompt, model=self.model, id=self.model_id).strip()
                print('#### Reply:')
                print(reply)
                print('#### Parsed:')
                print(utils.parse_api(reply))
            except Exception as e:
                print("Query Failed:", e)
                reply = "Query Failed!"
            
            if verbose:
                print(f"== Reply from AI ==\n{reply}\n\n")

            self.chat_history += [
                f"¬User¬\n{instruction}",
                f"¬AI¬:\n{reply}",
            ]
            reply_list.append(reply)

        return self.prompt, "\n".join(reply_list)
