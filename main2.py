from flask import Flask, render_template, request, jsonify
import os
from src import (
    api_selection,
    word_executor,
    modeling,
    utils
)

# 修改 Flask 应用配置
app = Flask(__name__,
    static_folder='static',    # 静态文件目录
    template_folder='templates' # 模板文件目录
)

# 初始化 API 选择器
class Args:
    def __init__(self):
        self.api_update = False
        self.api_lack = False
        self.dataset = 'short'
        self.model = "gpt-3.5-turbo"
        self.api_topk = 10
        self.planning = True
        self.api_selection = True
        self.content_selection = True
        self.model_id = None

args = Args()
word_assistant = None

@app.route('/')
def index():
    return render_template('opendoc/index.html')

@app.route('/doc')
def doc():
    return render_template('opendoc/doc.html')

@app.route('/api/analyze', methods=['POST'])
def analyze():
    global word_assistant
    
    try:
        data = request.json
        instruction = data.get('instruction', '')
        
        # 延迟初始化 word_assistant
        if word_assistant is None:
            word_assistant = modeling.word_assistant(args)
            word_assistant.load_docx(None)  # 创建新文档
            
        # 获取提示和API调用序列
        prompt, api_sequence = word_assistant.chat(instruction, verbose=True)
        
        # 解析API调用
        api_calls = utils.parse_api(api_sequence)
        
        # 执行API调用
        error_info = word_executor.API_executor(api_calls)
        
        # 获取当前文档
        doc = word_executor.get_word()
        
        # 保存文档
        temp_path = "static/temp/output.docx"
        os.makedirs(os.path.dirname(temp_path), exist_ok=True)
        doc.save(temp_path)
        
        return jsonify({
            'success': True,
            'api_calls': api_calls,
            'document_url': '/static/temp/output.docx',
            'error': error_info if error_info else None
        })
        
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        })

if __name__ == '__main__':
    # 确保 OpenAI API key 已设置
    if "OPENAI_API_KEY" not in os.environ:
        os.environ["OPENAI_API_KEY"] = "你的OpenAI API密钥"
    
    # 使用 5002 端口
    app.run(debug=True, port=5002)
