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
        self.word = None
        self.current_page_id = 0
        self.prompt = ""

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
        self.word = word_executor.get_word()
        self.current_page_id = word_executor.get_current_page_id()

    def load_chat_history(self, instructions, labels):
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

    def load_docx(self, docx_path=None):
        # 如果提供了文档路径，则加载该文档；否则，创建一个新的文档
        if docx_path:
            doc = Document(docx_path)
        else:
            doc = Document()

        # 返回当前文档对象
        return doc

    def chat(self, user_instruction, word_path=None, verbose=False):
        self.prompt = ""
        instruction_list = self.planner(user_instruction)
        reply_list = []

        # 加载 Word 文档
        doc = self.load_docx(word_path)
        current_paragraph = doc.add_paragraph()  # 新建一个段落

        for instruction in instruction_list:
            if verbose:
                print('Executing instruction: ', instruction)

            selected_apis = self.api_selector(instruction)
            API_string = "\n".join(map(str, selected_apis))
            if verbose:
                print(f"== Selected APIs ==\n{API_string}\n\n")

            word_content = self.content_selector(word_path, instruction, self.args, self.word)
            if verbose:
                print(word_content)

            prompt = prompt_factor.get_instruction_to_API_code_prompt2(
                API_string,
                word_content,
                self.chat_history,
                instruction,
                True,
                self.current_page_id,
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
                    self.current_page_id,
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
                        self.current_page_id,
                    )
            self.prompt += prompt + '\n\n'
            if verbose:
                print(f"== Prompt ==\n{prompt}\n\n")

            try:
                reply = openai_api.query_openai(prompt, model=self.model, id=self.model_id).strip()

                print('#### Reply:')
                print(reply)
                print('#### Parsed:')
                print(utils.parse_api(reply))
            except Exception as e:
                print("Query Failed!")
                print(e)
                reply = "Query Failed!"
            if verbose:
                print(f"== Reply from AI ==\n{reply}\n\n")

            # 将回复添加到 Word 文档中
            current_paragraph.add_run(reply)

            self.chat_history += [
                f"¬User¬\n{instruction}",
                f"¬AI¬:\n{reply}",
            ]
            reply_list.append(reply)

        # 保存文档
        doc.save("output.docx")  # 保存为 output.docx 文件
        return self.prompt, "\n".join(reply_list)
