import openai
from openai import (
    RateLimitError,
    APIConnectionError,
    APIError as ServiceUnavailableError
)
import backoff
import os
import tiktoken
import time

# 设置 OpenAI
openai.api_key = os.getenv("OPENAI_API_KEY")
openai.api_base = "https://xiaoai.plus/v1"

# 设置代理
os.environ['http_proxy'] = "http://127.0.0.1:7890"
os.environ['https_proxy'] = "http://127.0.0.1:7890"

# 增加重试装饰器的等待时间和最大重试次数
@backoff.on_exception(backoff.expo, 
                     (RateLimitError, APIConnectionError, ServiceUnavailableError),
                     max_tries=8,
                     max_time=300)
def get_embedding_with_backoff(text, engine="text-embedding-ada-002"):
    """获取文本嵌入的函数，带有重试机制"""
    try:
        return openai.Embedding.create(
            input=[text], 
            engine=engine
        )["data"][0]["embedding"]
    except Exception as e:
        print(f"Error getting embedding: {e}")
        raise

@backoff.on_exception(backoff.expo, 
                     (RateLimitError, APIConnectionError, ServiceUnavailableError),
                     max_tries=8,
                     max_time=300)
def completions_with_backoff(**kwargs):
    return openai.Completion.create(**kwargs)

@backoff.on_exception(backoff.expo, 
                     (RateLimitError, APIConnectionError, ServiceUnavailableError),
                     max_tries=8,
                     max_time=300)
def chat_with_backoff(**kwargs):
    return openai.ChatCompletion.create(**kwargs)

@backoff.on_exception(backoff.expo, RateLimitError)
@backoff.on_exception(backoff.expo, APIConnectionError)
def embeddings_with_backoff(**kwargs):
    return openai.Embedding.create(**kwargs)

def query_azure_openai(query, model="gpt-3.5-turbo"):
    """
    调用 OpenAI API，根据用户选择的模型进行查询。
    """
    if model == 'text3':
        response = completions_with_backoff(
            model="text-davinci-003",
            prompt=query,
            temperature=0,
            max_tokens=1000,
            top_p=0.95,
            frequency_penalty=0,
            presence_penalty=0,
            stop=["<|im_end|>", "¬User¬", "</decomposed>", "</query>"]
        )
        return response["choices"][0]["text"]

    elif model == 'turbo':
        prompt = "<|im_start|>system\nYou are a helpful assistant.\n<|im_end|>\n<|im_start|>user\nHello!\n<|im_end|>\n<|im_start|>assistant\nHow can I help you?\n<|im_end|>\n<|im_start|>user\n{0}\n<|im_end|>\n<|im_start|>assistant\n".format(query)
        response = completions_with_backoff(
            model="gpt-35-turbo",
            prompt=prompt,
            temperature=0,
            max_tokens=1000,
            top_p=0.95,
            frequency_penalty=0,
            presence_penalty=0,
            stop=["<|im_end|>", "¬User¬", "</decomposed>", "</query>"])
        return response["choices"][0]["text"]

    elif model == 'gpt4':
        response = chat_with_backoff(
            model="gpt-4-32k",
            messages=[
                {"role": "system", "content": "You are a helpful AI assistant."},
                {"role": "user", "content": query},
            ],
            temperature=0,
            max_tokens=1000,
            top_p=0.95,
            frequency_penalty=0,
            presence_penalty=0,
            stop=["<|im_end|>", "¬User¬", "</decomposed>", "</query>"])
        return response["choices"][0]["message"]["content"]
    else:
        encoding = tiktoken.encoding_for_model("gpt-3.5-turbo")

        def truncate_text_with_token_count(text, max_tokens):
            num_tokens = len(encoding.encode(text))

            if num_tokens > max_tokens:
                tokens = text.split()
                truncated_tokens = tokens[:max_tokens]
                truncated_text = ' '.join(truncated_tokens)
                return truncated_text
            return text

        truncated_input = query
        completion = openai.ChatCompletion.create(
            model=model,
            messages=[
                {"role": "system", "content": "You are a helpful AI assistant."},
                {"role": "user", "content": truncated_input},
            ],
            temperature=0,
            top_p=0.95,
        )

        try:
            return completion.choices[0].message.content
        except:
            return ' '

def rewrite(prompt):
    response = chat_with_backoff(
        model="gpt-4",
        messages=[
            {"role": "system", "content": "You are a helpful AI assistant."},
            {"role": "user", "content": prompt},
        ],
        temperature=0.7,
        max_tokens=200
    )
    return response["choices"][0]["message"]["content"]

if __name__ == "__main__":
    prompts = """根据用户指令"请添加一个表格"，你可以基于以下API和Word文档内容完成任务。
当前你位于第1段。
请根据你拥有的功能完成用户指令。
不要生成超出用户指令的内容。
不要猜测用户可能在下一步指示什么，或者为他们生成额外的API调用。
不要使用Python循环调用API。你只能一次调用一个API。

如果用户没有指定要修改的页面，你可以直接开始使用API，而不需要先导航到其他页面。

你需要生成的代码应该能完成用户指令。多行代码应该放在<code>和</code>标签中，例如：
<code>
API();
API();
</code>

目前，你可以使用以下Word文档API：
API: create_and_save_docx(content, save_path)：此API用于创建并保存一个新的docx文档。
API: rename_docx_name(original_path, new_name)：此API用于重命名现有的docx文档，但不改变其径。
API: move_docx_to_new_path(original_path, new_directory)：此API用于将docx文件移动到一个新目录，但不改变文件名。
API: add_paragraph(doc, content, font_name='Arial', font_size=Pt(12), color=RGBColor(0, 0, 0), spacing=Pt(1.15))：此API用于向docx文档添加一个段落，并设置格式。
API: add_heading(doc, content, level=1, font_name='Arial', font_size=Pt(14), color=RGBColor(0, 0, 0), spacing=Pt(1.15))：此API用于向docx文档添加一个标题，并设置格式。
API: find_paragraphs_by_heading_and_content(doc, heading, content)：此API用于根据指定的标题和内容查找文档中的段落。
API: add_header(doc_path, text, section_index=None)：此API用于添加页眉。如果指定section_index，则仅为该节添加页眉；否则为所有节添加页眉。
API: delete_header(doc_path, section_index=None)：此API用于删除页眉。如果指定section_index，则仅删除该节的页眉；否则删除所有节的页眉。
API: add_footer(doc_path, text, section_index=None)：此API用于添加页脚。如果指定section_index，则仅为该节添加页脚；否则为所有节添加页脚。
API: delete_footer(doc_path, section_index=None)：此API用于删除页脚。如果指定section_index，则仅删除该节的页脚；否则删除所有节的页脚。
API: add_line_break(doc_path, paragraph_index, position=None)：此API用于在指定段落中添加换行符。
API: delete_line_break(doc_path, paragraph_index, break_index)：此API用于删除指定段落中的换行符。
API: add_paragraph(doc_path, text, position=None, style=None)：此API用于添加段落，支持自定义样式和插入位置。
API: add_style(doc_path, style_name, font_name=None, font_size=None, bold=None, italic=None, color=None)：此API用于添加新样式。
API: delete_style(doc_path, style_name)：此API用于删除指定样式。
API: add_page_numbers(doc_path)：此API用于为文档添加页码。
API: delete_page_numbers(doc_path)：此API用于删除文档的页码。
API: add_table_of_contents(doc_path, levels=3)：此API用于添加目录，支持指定级别。
API: delete_table_of_contents(doc_path)：此API用于删除目录。
API: add_watermark(doc_path, text)：此API用于添加文字水印。
API: delete_watermark(doc_path)：此API用于删除水印。
API: delete_paragraph(doc_path, paragraph_index)：此API用于删除指定索引的段落。
API: add_text(doc_path, paragraph_index, text, position=None, bold=False, italic=False, underline=False, color=None, size=None)：此 API 用于在指定段落位置添加文本，可以设置文本样式。
API: delete_text(doc_path, paragraph_index, start, end)：此 API 用于删除指定段落中从 start 到 end 的文本。
API: add_hyperlink(doc_path, paragraph_index, text, url, position=None)：此 API 用于在指定段落添加超链接。
API: delete_hyperlink(doc_path, paragraph_index, link_index)：此 API 用于删除指定段落中的超链接。
API: add_heading(doc_path, text, level, position=None)：此 API 用于在文档中添加标题。
API: delete_heading(doc_path, level, occurrence)：此 API 用于删除指定级别的第 n 个标题。
API: add_table(doc_path, rows, cols, position=None, style=None)：此 API 用于在文档中添加表格。
API: set_font_size(size)：此 API 用于设置当前段落字体大小。
API: set_font_color(color)：此 API 用于设置当前段落字体颜色。
API: set_font_bold()：此 API 用于将当前段落的字体设置为加粗。
API: set_font_italic()：此 API 用于将当前段落的字体设置为斜体。
API: set_font_underline()：此 API 用于在当前段落字体下添加下划线。
API: set_font_style(font_name)：此 API 用于设置当前段落的字体样式。
API: set_line_space(line_space_level)：此 API 用于设置当前段落的行间距。
API: text_align_left()：此 API 用于将当前段落左对齐。
API: text_align_center()：此 API 用于将当前段落居中对齐。
API: text_align_right()：此 API 用于将当前段落右对齐。
API: add_run(doc, text, font_size=12, font_color='000000', bold=False, italic=False, underline=False)：此 API 用于在当前段落中添加一个新的文本块，并设置文本的样式。
API: get_word()：此 API 用于获取当前操作的 Word 文档对象。
API: get_current_page_id()：此 API 用于获取当前文档的页编号。
API: init_docx()：此 API 用于初始化一个新的 Word 文档。
API: set_word(docx_path=None)：此 API 用于加载现有文档或初始化一个新的 Word 文档。
API: set_current_paragraph(index)：此 API 用于设置当前操作的段落。
API: get_table_count(doc)：此API用于获取文档中已存在的表格数量。
API: add_table(doc_path, rows, cols, position=None, style=None)：此API用于在文档中添加一个新的表格。
API: set_table_title(doc, title, font_size=12, bold=True, color='black', alignment=WD_PARAGRAPH_ALIGNMENT.CENTER)：此API用于为表格添加标题，并设置标题样式。
API: delete_table(doc_path, table_index)：此API用于删除指定索引的表格。
API: add_table_header(table, headers, font_size=12, bold=True, alignment=WD_ALIGN_PARAGRAPH.CENTER)：此API用于添加或修改表格的表头行。
API: set_cell_text(table, row, col, text, font_size=12, bold=False)：此API用于设置表格单元格的文本内容和样式。
API: add_table_row(doc_path, table_index)：此API用于向指定表格添加一行。
API: add_table_column(doc_path, table_index)：此API用于向指定表格添加一列。
API: delete_table_column(doc_path, table_index, col_index)：此API用于删除表格中的指定列。
API: add_list_item(doc_path, text, level=0, style='List Bullet')：此API用于在文档中添加一个列表项。
API: delete_list_item(doc_path, item_index)：此API用于删除指定的列表项及其所有子项。
API: set_column_width(table, col_idx, width_in_inches)：此API用于设置表格某列的宽度。
API: delete_table_row(doc_path, table_index, row_index)：此API用于删除表格中的指定行。
API: set_row_height(table, row_idx, height_in_inches)：此API用于设置表格某行的高度。
API: merge_cells(table, start_row, start_col, end_row, end_col)：此API用于合并表格中的指定单元格。
API: set_cell_bg_color(table, row, col, color_hex)：此API用于设置表格单元格的背景颜色。
API: align_cell_text(table, row, col, alignment=WD_ALIGN_PARAGRAPH.CENTER)：此API用于设置表格单元格中文本的对齐方式。
API: replace_picture(doc, old_picture_path, new_picture_path, width_inch=None, height_inch=None)：此API用于替换文档中的旧图片为新图片，并可调整新图片的大小。
API: add_image(doc_path, paragraph_index, image_path, width=None, height=None, position=None)：此API用于在指定段落中添加图片。
API: delete_image(doc_path, paragraph_index, image_index)：此API用于删除指定段落中的图片。
API: set_image_size(image, width=None, height=None)：此API用于设置插入图片的大小。
API: align_image_center(paragraph)：此API用于将图片居中对齐。
API: align_image_right(paragraph)：此API用于将图片右对齐。
API: add_caption(doc, text, font_size=10)：此API用于为图片添加文字说明。
API: create_line_chart(x, y, title='Line Chart', xlabel='X Axis', ylabel='Y Axis', save_path='line_chart.png')：此API用于创建并保存折线图。
API: create_bar_chart(x, y, title='Bar Chart', xlabel='X Axis', ylabel='Y Axis', save_path='bar_chart.png')：此API用于创建并保存柱状图。
API: create_pie_chart(labels, sizes, title='Pie Chart', save_path='pie_chart.png')：此API用于创建并保存饼图。
API: insert_image_into_docx(image_path, docx_path)：此API用于将图片插入到Word文档中。
API: generate_chart(chart_type, data, chart_title='Chart', docx_path='output.docx')：此API根据指定的图表类型和数据生成图表，并将图表插入到Word文档中。
API: set_width(width)：此API用于设置所选对象的宽度。
API: set_height(height)：此API用于设置所选对象的高度。
API: rotate_element(angle)：此API用于按指定角度旋转所选对象。
API: set_fill_color(color)：此API用于设置所选对象的填充颜色。


文档内容如下：
<文档开始>  
当前文档有1页，页面高度为6858，页面宽度为9144。   
[段落1]  
大小：高度=1143，宽度=8229  
文本：   

字体样式：粗体=无，斜体=无，下划线=无，大小=无，颜色=无，字体样式=无，行间距=无，文本对齐=无  
"""
    answer = query_azure_openai(prompts, model="gpt-3.5-turbo")
    print(answer)
