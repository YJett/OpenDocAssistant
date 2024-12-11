import docx
import matplotlib.pyplot as plt
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from src import api_doc,word_reader
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx import Document
from docx.shared import Inches, Pt, RGBColor
import os

import logging

# 设置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


# document
# 1. 创建并保存一个新的 docx 文档
def create_and_save_docx(content, save_path):
    try:
        doc = Document()
        doc.add_paragraph(content)
        doc.save(save_path)
        print(f"文档已成功创建并保存到: {save_path}")
        return True
    except Exception as e:
        print(f"发生错误: {e}")
        return False
    
def save_docx(path):
    """保存Word文档"""
    try:
        # 确保目录存在
        os.makedirs(os.path.dirname(path), exist_ok=True)
        doc = Document()
        doc.save(path)
        logger.info(f"Document saved to: {path}")
        return True
    except Exception as e:
        logger.error(f"Error saving document: {e}")
        return False    


# 2. 修改 docx 文件的名称（不改变路径）
def rename_docx_name(original_path, new_name):
    try:
        if not os.path.exists(original_path):
            print(f"错误: 文件 {original_path} 不存在！")
            return False
        directory = os.path.dirname(original_path)
        new_path = os.path.join(directory, new_name)
        os.rename(original_path, new_path)
        print(f"文件成功重命名为: {new_path}")
        return True
    except Exception as e:
        print(f"发生错误: {e}")
        return False


# 3. 修改 docx 文件的路径（不改变文件名）
def move_docx_to_new_path(original_path, new_directory):
    try:
        if not os.path.exists(original_path):
            print(f"错误: 文件 {original_path} 不存在！")
            return False
        file_name = os.path.basename(original_path)
        new_path = os.path.join(new_directory, file_name)
        os.rename(original_path, new_path)
        print(f"文件成功移动到新路径: {new_path}")
        return True
    except Exception as e:
        print(f"发生错误: {e}")
        return False


# 4. 创建段落并设置格式
def add_paragraph(text, style=None):
    """添加���落
    
    Args:
        text: 段落文本
        style: 段落样式(可选)
    
    Returns:
        新添加的段落对象
    """
    global doc, current_paragraph
    try:
        current_paragraph = doc.add_paragraph(text, style=style)
        return current_paragraph
    except Exception as e:
        logger.error(f"Error adding paragraph: {e}")
        return None




# 6. 根据标题查找段落
def find_paragraphs_by_heading_and_content(doc, heading, content):
    paragraphs = []
    found_heading = False
    for para in doc.paragraphs:
        if para.style.name.startswith('Heading') and para.text == heading:
            found_heading = True
            continue
        if found_heading and para.style.name.startswith('Normal') and para.text.strip() == content:
            paragraphs.append(para.text)
        if found_heading and para.style.name.startswith('Heading'):
            break
    if paragraphs:
        print(f"找到了标题 '{heading}' 下匹配内容 '{content}' 的段落: {paragraphs}")
    else:
        print(f"没有找到标题 '{heading}' 下匹配内容 '{content}' 的段落。")
    return paragraphs


def delete_header(doc_path, section_index=None):
    """
    删除页眉
    如果指定section_index,则只删除该节的页眉
    否则删除所有节的页眉
    """
    doc = Document(doc_path)
    if section_index is not None:
        if 0 <= section_index < len(doc.sections):
            section = doc.sections[section_index]
            # 断开与前一节的链接
            section.header.is_linked_to_previous = False
            section.header.paragraphs[0].text = ""
    else:
        for section in doc.sections:
            # 断开与前一节的链接
            section.header.is_linked_to_previous = False
            section.header.paragraphs[0].text = ""
    doc.save(doc_path)


def delete_footer(doc_path, section_index=None):
    """
    删除页脚
    如果指定section_index,则只删除该节的页脚
    否则删除所有节的页脚
    """
    doc = Document(doc_path)
    if section_index is not None:
        if 0 <= section_index < len(doc.sections):
            section = doc.sections[section_index]
            # 断开与前一节的链接
            section.footer.is_linked_to_previous = False
            section.footer.paragraphs[0].text = ""
    else:
        for section in doc.sections:
            # 断开与前一节的链接
            section.footer.is_linked_to_previous = False
            section.footer.paragraphs[0].text = ""
    doc.save(doc_path)


def add_line_break(doc_path, paragraph_index, position=None):
    """
    在指定段落中添加换行符

    参数:
    - paragraph_index: 段落索引
    - position: 在段落中插入的位置(如果不指定,则在段落末尾添加)
    """
    doc = Document(doc_path)
    if 0 <= paragraph_index < len(doc.paragraphs):
        paragraph = doc.paragraphs[paragraph_index]

        # 创建新的 run 并移动到指定位置
        new_run = paragraph.add_run()
        if position is not None and 0 <= position < len(paragraph.runs):
            paragraph._p.insert(position, new_run._r)
            paragraph.runs.insert(position, new_run)
            paragraph.runs.pop()

        new_run.add_break(WD_BREAK.LINE)

    doc.save(doc_path)


def delete_line_break(doc_path, paragraph_index, break_index):
    """
    删除指定段落中的换行符

    参数:
    - paragraph_index: 段落索引
    - break_index: 换行符在段落中的索引
    """
    doc = Document(doc_path)
    if 0 <= paragraph_index < len(doc.paragraphs):
        paragraph = doc.paragraphs[paragraph_index]
        breaks = paragraph._element.findall('.//w:br', namespaces={
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
        if 0 <= break_index < len(breaks):
            line_break = breaks[break_index]
            parent = line_break.getparent()
            if parent is not None:
                parent.remove(line_break)
                # 如果 run 为空，则删除该 run
                if not parent.getchildren():
                    parent.getparent().remove(parent)


def add_style(doc_path, style_name, font_name=None, font_size=None, bold=None, italic=None, color=None):
    """
    添加新样式

    参数:
    - style_name: 样式名称
    - font_name: 字名称
    - font_size: 字体大小(磅)
    - bold: 是否加粗
    - italic: 是否斜体
    - color: RGB颜色元组,如(255, 0, 0)

    返回: bool 是否添加成功
    """
    doc = Document(doc_path)

    # 检查样式是否已存在
    if style_name in doc.styles:
        return False

    try:
        style = doc.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
        if font_name:
            style.font.name = font_name
        if font_size:
            style.font.size = Pt(font_size)
        if bold is not None:
            style.font.bold = bold
        if italic is not None:
            style.font.italic = italic
        if color:
            style.font.color.rgb = RGBColor(*color)
        doc.save(doc_path)
        return True
    except Exception:
        return False


def delete_style(doc_path, style_name):
    """
    删除指定样式

    参数:
    - style_name: 要删除的样式名称

    返���: bool 是否删除成功
    """
    doc = Document(doc_path)
    try:
        # 将使用样式的段落重置为默认样式
        for paragraph in doc.paragraphs:
            if paragraph.style.name == style_name:
                paragraph.style = doc.styles['Normal']

        # 删除样式
        doc.styles._element.remove(doc.styles[style_name]._element)
        doc.save(doc_path)
        return True
    except Exception:
        return False


def add_page_numbers(doc_path):
    """添加页码（为了测试删除页码功能）"""
    doc = Document(doc_path)
    sections = doc.sections
    for section in sections:
        footer = section.footer
        paragraph = footer.paragraphs[0]
        run = paragraph.add_run()

        # 添加页码字段
        fldChar1 = OxmlElement('w:fldChar')
        fldChar1.set(qn('w:fldCharType'), 'begin')
        run._r.append(fldChar1)

        instrText = OxmlElement('w:instrText')
        instrText.text = "PAGE"
        run._r.append(instrText)

        fldChar2 = OxmlElement('w:fldChar')
        fldChar2.set(qn('w:fldCharType'), 'end')
        run._r.append(fldChar2)

    doc.save(doc_path)


def delete_page_numbers(doc_path):
    """
    删除页码
    """
    doc = Document(doc_path)
    for section in doc.sections:
        # 检查页眉
        for paragraph in section.header.paragraphs:
            if 'PAGE' in paragraph._element.xml:
                paragraph._element.getparent().remove(paragraph._element)

        # 检查页脚
        for paragraph in section.footer.paragraphs:
            if 'PAGE' in paragraph._element.xml:
                paragraph._element.getparent().remove(paragraph._element)

    doc.save(doc_path)


def add_table_of_contents(doc_path, levels=3):
    """
    添加目录

    参数:
    - levels: 目录级别(1-9)
    """
    doc = Document(doc_path)
    paragraph = doc.add_paragraph()
    run = paragraph.add_run()

    fldChar = OxmlElement('w:fldChar')
    fldChar.set(qn('w:fldCharType'), 'begin')
    run._element.append(fldChar)

    instrText = OxmlElement('w:instrText')
    instrText.text = f'TOC \\o "1-{levels}" \\h \\z \\u'
    run._element.append(instrText)

    fldChar = OxmlElement('w:fldChar')
    fldChar.set(qn('w:fldCharType'), 'end')
    run._element.append(fldChar)

    doc.save(doc_path)


def delete_table_of_contents(doc_path):
    """
    删除目录
    """
    doc = Document(doc_path)
    for paragraph in doc.paragraphs:
        if 'TOC' in paragraph._element.xml:
            paragraph._element.getparent().remove(paragraph._element)

    doc.save(doc_path)


def add_watermark(doc_path, text):
    """添加文字水印"""
    doc = Document(doc_path)
    for section in doc.sections:
        header = section.header
        paragraph = header.paragraphs[0]
        run = paragraph.add_run()

        # 创建水印
        hdr = parse_xml(f'''
            <w:pict xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml">
                <v:shape id="_x0000_i1025" type="#_x0000_t136" style="position:absolute;margin-left:0;margin-top:0;width:468pt;height:351pt;rotation:315;z-index:-251654144">
                    <v:textpath style="font-family:'Calibri';font-size:1pt" string="{text}"/>
                </v:shape>
            </w:pict>
            ''')

        run._r.append(hdr)

    doc.save(doc_path)


def delete_watermark(doc_path):
    """
    删除水印
    """
    doc = Document(doc_path)
    for section in doc.sections:
        header = section.header
        for paragraph in header.paragraphs:
            for run in paragraph.runs:
                if 'pict' in run._element.xml:
                    run._element.getparent().remove(run._element)

    doc.save(doc_path)


def delete_paragraph(doc_path, paragraph_index):
    """
    删除指定索引的段落
    """
    doc = Document(doc_path)
    if 0 <= paragraph_index < len(doc.paragraphs):
        p = doc.paragraphs[paragraph_index]._element
        p.getparent().remove(p)
    doc.save(doc_path)


# # 示例使用：
# content = "这是一个新的文档，包一些示例内容。"
# save_path = "new_document_with_headings.docx"
# create_and_save_docx(content, save_path)
#
# doc = Document()
# add_heading(doc, "一级标题", level=1, font_name='Calibri', font_size=Pt(16), color=RGBColor(0, 0, 255), spacing=Pt(1.5))
# add_heading(doc, "二级标题", level=2, font_name='Calibri', font_size=Pt(14), color=RGBColor(255, 0, 0), spacing=Pt(1.2))
# add_paragraph(doc, "这是一个普通的段落。", font_name='Times New Roman', font_size=Pt(11), color=RGBColor(255, 255, 0), spacing=Pt(1.1))
#
# add_heading(doc, "三级标题", level=3, font_name='Calibri', font_size=Pt(12), color=RGBColor(0, 255, 0), spacing=Pt(1.0))
# add_paragraph(doc, "这是一个普通的段落。", font_name='Times New Roman', font_size=Pt(11), color=RGBColor(255, 255, 0), spacing=Pt(1.1))
# add_paragraph(doc, "这是另一个段落，紧接在标题后面。", font_name='Arial', font_size=Pt(10), color=RGBColor(128, 0, 128), spacing=Pt(1.0))
# doc.save("new_document_with_headings.docx")
# content = "这是一普通的段落。"
# found_paragraphs = find_paragraphs_by_heading_and_content(doc, "二级标题",content)


# text


color2hex = {
    "blue": "0000FF", "light blue": "ADD8E6", "dark blue": "00008B",
    "green": "008000", "light green": "90EE90", "dark green": "006400",
    "yellow": "FFFF00", "light yellow": "FFFFE0", "dark yellow": "BDB76B",
    "orange": "FFA500", "light orange": "FFDAB9", "dark orange": "FF8C00",
    "red": "FF0000", "light red": "FFC0CB", "dark red": "8B0000",
    "black": "000000", "white": "FFFFFF", "purple": "800080", "pink": "FFC0CB",
}


def add_text(doc_path, paragraph_index, text, position=None, bold=False, italic=False, underline=False, color=None,
             size=None):
    doc = Document(doc_path)
    if 0 <= paragraph_index < len(doc.paragraphs):
        paragraph = doc.paragraphs[paragraph_index]
        if position is not None:
            # 分割现有文本并创建新的运行块
            old_text = paragraph.text
            paragraph.clear()
            if position > 0:
                run = paragraph.add_run(old_text[:position])
            run = paragraph.add_run(text)
            if position < len(old_text):
                paragraph.add_run(old_text[position:])
        else:
            run = paragraph.add_run(text)

        run.bold = bold
        run.italic = italic
        run.underline = underline
        if color:
            run.font.color.rgb = RGBColor(*color)
        if size:
            run.font.size = Pt(size)

    doc.save(doc_path)


def delete_text(doc_path, paragraph_index, start, end):
    doc = Document(doc_path)
    if 0 <= paragraph_index < len(doc.paragraphs):
        paragraph = doc.paragraphs[paragraph_index]
        runs = paragraph.runs
        text = paragraph.text

        # 保存现有运行块的格式信息
        formats = []
        current_pos = 0
        for run in runs:
            formats.append({
                'text': run.text,
                'start': current_pos,
                'end': current_pos + len(run.text),
                'bold': run.bold,
                'italic': run.italic,
                'underline': run.underline,
                'color': run.font.color.rgb,
                'size': run.font.size
            })
            current_pos += len(run.text)

        # 创建新文本
        new_text = text[:start] + text[end:]

        # 重新应用格式
        paragraph.clear()
        current_pos = 0
        for fmt in formats:
            if current_pos < len(new_text):
                # 计算新的文本范围
                new_start = max(0, fmt['start'] - (end - start))
                new_end = max(0, fmt['end'] - (end - start))
                if new_start < new_end:
                    run = paragraph.add_run(new_text[new_start:new_end])
                    run.bold = fmt['bold']
                    run.italic = fmt['italic']
                    run.underline = fmt['underline']
                    if fmt['color']:
                        run.font.color.rgb = fmt['color']
                    if fmt['size']:
                        run.font.size = fmt['size']
                    current_pos = new_end

    doc.save(doc_path)


from docx.oxml.shared import OxmlElement, qn


def add_hyperlink(doc_path, paragraph_index, text, url, position=None):
    """
    在指定段落中添加超链接

    参数:
    - paragraph_index: 段落索引
    - text: 链接显示的文本
    - url: 链接地址
    - position: 在段落中插入的位置(如果不指定,则在段落末尾添加)
    """
    doc = Document(doc_path)
    if 0 <= paragraph_index < len(doc.paragraphs):
        paragraph = doc.paragraphs[paragraph_index]

        # 创建新的运行块
        run = paragraph.add_run()
        run.text = text
        run.font.underline = True
        run.font.color.rgb = RGBColor(0, 0, 255)  # 蓝色

        # 添加 rStyle
        rStyle = OxmlElement('w:rStyle')
        rStyle.set(qn('w:val'), 'Hyperlink')
        run._element.rPr.append(rStyle)

        # 添加超链接关系
        r_id = doc.part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
        hyperlink = OxmlElement('w:hyperlink')
        hyperlink.set(qn('r:id'), r_id)
        hyperlink.append(run._element)

        # 处理插入位置
        if position is not None and 0 <= position <= len(paragraph._p):
            paragraph._p.insert(position, hyperlink)
        else:
            paragraph._p.append(hyperlink)

    doc.save(doc_path)


def delete_hyperlink(doc_path, paragraph_index, link_index):
    """
    删除指定段落中的超链接

    参数:
    - paragraph_index: 段落索引
    - link_index: 超链接在段落中的索引
    """
    doc = Document(doc_path)
    if 0 <= paragraph_index < len(doc.paragraphs):
        paragraph = doc.paragraphs[paragraph_index]
        hyperlinks = paragraph._p.xpath('.//w:hyperlink')
        if 0 <= link_index < len(hyperlinks):
            hyperlink = hyperlinks[link_index]
            for run in hyperlink.xpath('.//w:r'):
                hyperlink.addnext(run)
            hyperlink.getparent().remove(hyperlink)

    doc.save(doc_path)


def add_heading(doc_path, text, level, position=None):
    """
    添加标题

    参数:
    - level: 标题级别(1-9)
    - position: 如果指定,在特定位置插入标题
    """
    doc = Document(doc_path)
    if position is not None:
        heading = doc.add_heading(level=level)
        doc._body._body.insert(position, heading._element)
    else:
        heading = doc.add_heading(level=level)
    heading.text = text
    doc.save(doc_path)


def delete_heading(doc_path, level, occurrence):
    """
    删除指定级别的第n个标题

    参数:
    - level: 标题级别
    - occurrence: 第几个该级别的标题(从1开始)
    """
    doc = Document(doc_path)
    count = 0
    for i, paragraph in enumerate(doc.paragraphs):
        if paragraph.style.name.startswith(f'Heading {level}'):
            count += 1
            if count == occurrence:
                p = doc.paragraphs[i]._element
                p.getparent().remove(p)
                break
    doc.save(doc_path)




# 设置字体大小
def set_font_size(size):
    global current_paragraph
    for run in current_paragraph.runs:
        run.font.size = Pt(size)


# 设置字体颜色
def set_font_color(color):
    global current_paragraph
    for run in current_paragraph.runs:
        run.font.color.rgb = RGBColor.from_string(color2hex[color])


# 设置加粗字体
def set_font_bold():
    global current_paragraph
    for run in current_paragraph.runs:
        run.font.bold = True


# 设置斜体
def set_font_italic():
    global current_paragraph
    for run in current_paragraph.runs:
        run.font.italic = True


# 设置下划线
def set_font_underline():
    global current_paragraph
    for run in current_paragraph.runs:
        run.font.underline = True


# 设置字体样式
def set_font_style(font_name):
    global current_paragraph
    for run in current_paragraph.runs:
        run.font.name = font_name


# 设置行间距
def set_line_space(line_space_level):
    global current_paragraph
    current_paragraph.paragraph_format.line_spacing = Pt(line_space_level)


# 左对齐
def text_align_left():
    """文本左对齐"""
    global doc, current_paragraph
    try:
        if current_paragraph is None:
            current_paragraph = doc.add_paragraph()
        current_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        return current_paragraph
    except Exception as e:
        logger.error(f"Error aligning text left: {e}")
        return None


# 居中对齐
def text_align_center():
    """文本居中对齐"""
    global doc, current_paragraph
    try:
        if current_paragraph is None:
            current_paragraph = doc.add_paragraph()
        current_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        return current_paragraph
    except Exception as e:
        logger.error(f"Error aligning text center: {e}")
        return None


# 右对齐
def text_align_right():
    """文本右对"""
    global doc, current_paragraph
    try:
        if current_paragraph is None:
            current_paragraph = doc.add_paragraph()
        current_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        return current_paragraph
    except Exception as e:
        logger.error(f"Error aligning text right: {e}")
        return None


def add_run(doc, text, font_size=12, font_color='000000', bold=False, italic=False, underline=False):
    p = doc.paragraphs[-1]  # 添加到最新的段落
    run = p.add_run(text)
    run.font.size = Pt(font_size)
    run.font.color.rgb = RGBColor.from_string(font_color)
    run.font.bold = bold
    run.font.italic = italic
    run.font.underline = underline


doc = None


def get_word():
    global doc
    return doc


def get_current_page_id():
    global doc
    # 在Word中，我们可以使用段落的数量来模拟页面编号
    # 因为Word文档不是分页存储的，所以这里的"页面编号"可能不完全准确
    # 但可以作为一个参考
    return len(doc.paragraphs)


def init_docx():
    """初始化新文档"""
    doc = Document()
    word_reader.set_document(doc)  # 设置当前文档
    return doc


def check_api_in_list(api_call, api_list):
    """检查API调用是否在指定的API列表中
    
    Args:
        api_call: API调用字符
        api_list: API名称列表
    Returns:
        bool: 是否在列表中
    """
    api_name = api_call.split('(')[0]
    return api_name in api_list

def API_executor(lines, test=False, args=None):
    """行API调用序列
    
    Args:
        lines: API调用序列（字符串列表）
        test: 是否为测试模式
        args: 参数配置
    Returns:
        error_info: 错误信息字符串
    """
    error_info = ""
    for line in lines:
        # 测试模式外的API过
        if not test:
            # 跳过位置相关API
            if check_api_in_list(line, ["set_left", "set_top", "set_right", "set_bottom"]):
                continue
            # API缺失模式下的过滤
            if args and args.api_lack:
                if not check_api_in_list(line, api_doc.original_apis):
                    continue
        
        try:
            # 长文档模式的特殊处理
            if args and args.dataset == 'long' and line == 'add_heading()':
                eval("add_paragraph('', style='Heading 1')")
                continue
            
            # API缺失模式
            if args and args.api_lack:
                if line == 'seek_assistance()':
                    continue
                elif not check_api_in_list(line, api_doc.original_apis):
                    # 添加标记段落
                    eval("add_paragraph('API不可用: @@@@@@@@@@')")
                else:
                    eval(line)
            
            # API更新模式
            elif args and args.api_update:
                if check_api_in_list(line, api_doc.update_apis):
                    # 添加标记段落
                    eval("add_paragraph('API已更新: @@@@@@@@@@')")
                else:
                    eval(line)
            
            # 正模式
            else:
                eval(line)
                
        except Exception as e:
            error_msg = f"ERROR: {line}\n{str(e)}"
            print(error_msg)
            error_info += error_msg + "\n"
    
    return error_info

def set_word(docx_path=None):
    """设置当前操作的Word文档并初始化全局变量
    
    Args:
        docx_path: Word文档路径，如果为None则创建新文档
        
    全局变量:
        doc: 当前Word文档对象
        current_paragraph: 当前操作的段落
        current_table: 当前操作的表格
        current_picture: 当前操作的图片
        current_heading: 当前操作的标题
    """
    global doc, current_paragraph, current_table, current_picture, current_heading
    
    # 载或创建文
    if docx_path:
        doc = Document(docx_path)
    else:
        doc = Document()
    
    # 重置所有当前操作对象
    current_paragraph = None
    current_table = None
    current_picture = None
    current_heading = None
    
    # 设置文档阅读器的文档对象
    word_reader.set_document(doc)
    
    return doc

def get_word():
    """获取当前Word文档对象"""
    global doc
    return doc

def get_current_paragraph():
    """获取当前操作的段落"""
    global current_paragraph
    return current_paragraph

def set_current_paragraph(paragraph):
    """设置当前操作的段落"""
    global current_paragraph
    current_paragraph = paragraph

def get_current_table():
    """获取当前操作的表格"""
    global current_table
    return current_table

def set_current_table(table):
    """设置当前操作的表格"""
    global current_table
    current_table = table

def get_current_picture():
    """获取当前操作的图片"""
    global current_picture
    return current_picture

def set_current_picture(picture):
    """设置当前操作的图片"""
    global current_picture
    current_picture = picture

def get_current_heading():
    """获取当前操作的标题"""
    global current_heading
    return current_heading

def set_current_heading(heading):
    """设置当前操作的标题"""
    global current_heading
    current_heading = heading

def init_docx():
    """始化一个新的Word文档"""
    return set_word()

# # 示例使用
# doc = Document()
#
# # 全局变量
# current_paragraph = None  # 当前选中的段落
# paragraph = doc.add_paragraph("This is a sample paragraph.")
# current_paragraph = paragraph  # 置当前段落
#
# # 示例调用函数设置样式
# set_font_size(12)
# set_font_color("blue")
# set_font_bold()
# set_font_italic()
# set_font_underline()
# set_font_style("Arial")
# set_line_space(18)
# text_align_center()
#
# # 保存文档
# doc.save("text_document.docx")

# table

# 创建一个 Word 文档
doc = Document()


# 获取当前文档中已存在的表格数量
def get_table_count(doc):
    """
    获取文档中表格的数量，返回已添加表格的数量。
    """
    return len(doc.tables)


def add_table(doc_path, rows, cols, position=None, style=None):
    """
    添加表格

    参数:
    - rows: 行数
    - cols: 列数
    - position: 插入位置
    - style: 表格样式
    """
    doc = Document(doc_path)
    table = doc.add_table(rows=rows, cols=cols)
    if style:
        table.style = style
    if position is not None:
        doc._body._body.insert(position, table._element)
    doc.save(doc_path)

# 设置表格标题
def set_table_title(doc, title, font_size=12, bold=True, color="black", alignment=WD_PARAGRAPH_ALIGNMENT.CENTER):
    """
    设置表格标题并为表格添加编号。

    参数：
        doc (Document): docx 文档对象
        title (str): 表格标题文本
        font_size (int): 标字体大小
        bold (bool): 是否加粗
        color (str): 标题字体颜色
        alignment (WD_PARAGRAPH_ALIGNMENT): 标题文本对齐方式
    """
    table_count = get_table_count(doc) + 1  # 表格的编号是当前表格数 + 1
    table_title = f"Table {table_count}: {title}"

    # 在文档中添加标题段落，并设置居中对齐
    title_paragraph = doc.add_paragraph(table_title)
    title_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER  # 设置标题居中

    # 设置标题样式
    set_font_size(font_size)
    set_font_color(color)
    if bold:
        set_font_bold()
    set_line_space(15)  # 设置标题的行间距


def delete_table(doc_path, table_index):
    """
    删除指定索引的表格
    """
    doc = Document(doc_path)
    if 0 <= table_index < len(doc.tables):
        table = doc.tables[table_index]._element
        table.getparent().remove(table)
    doc.save(doc_path)


def add_table_header(table, headers, font_size=12, bold=True, alignment=WD_ALIGN_PARAGRAPH.CENTER):
    """
    添加或修改表格的表头行。

    参数：
        table (Table): 表格对象
        headers (list): 表头文本的列表
        font_size (int): 表头字体大小
        bold (bool): 否加粗
        alignment (WD_ALIGN_PARAGRAPH): 表头文本的对齐方式
    """
    if all(cell.text == '' for cell in table.rows[0].cells):
        header_row = table.rows[0]
    else:
        header_row = table.add_row().cells
        table.rows.insert(0, table.rows.pop())

    for idx, header in enumerate(headers):
        cell = header_row.cells[idx]
        cell.text = header
        for paragraph in cell.paragraphs:
            paragraph.alignment = alignment
            run = paragraph.runs[0]
            run.font.size = Pt(font_size)
            run.font.bold = bold


def set_cell_text(table, row, col, text, font_size=12, bold=False):
    """
    设置表格单元格的文本内容和样式。

    参数：
        table (Table): 表格对象
        row (int): 单元格的行索引
        col (int): 单元格的列索引
        text (str): 要设置的文本
        font_size (int): 字体大小
        bold (bool): 是否加粗
    """
    cell = table.cell(row, col)
    cell.text = text
    if cell.paragraphs:
        paragraph = cell.paragraphs[0]
        if paragraph.runs:
            run = paragraph.runs[0]
            run.font.size = Pt(font_size)
            run.font.bold = bold
        else:
            run = paragraph.add_run()
            run.text = text
            run.font.size = Pt(font_size)
            run.font.bold = bold


def add_table_row(doc_path, table_index):
    """
    向指定表格添加行
    """
    doc = Document(doc_path)
    if 0 <= table_index < len(doc.tables):
        doc.tables[table_index].add_row()
    doc.save(doc_path)


def add_table_column(doc_path, table_index):
    """
    向指定表格添加列
    """
    doc = Document(doc_path)
    if 0 <= table_index < len(doc.tables):
        table = doc.tables[table_index]
        for row in table.rows:
            cell = row._tr.add_tc()
            table.cell(row._index, len(row.cells) - 1)._tc = cell
            cell = table.cell(row._index, len(row.cells) - 1)
            cell.width = Inches(1)  # 设置新列的宽度，可以根据需要调整
    doc.save(doc_path)


def delete_table_column(doc_path, table_index, col_index):
    """
    删除表格中的指定列
    """
    doc = Document(doc_path)
    if 0 <= table_index < len(doc.tables):
        table = doc.tables[table_index]
        for row in table.rows:
            if 0 <= col_index < len(row.cells):
                tc = row.cells[col_index]._element
                tc.getparent().remove(tc)
    doc.save(doc_path)


def add_list_item(doc_path, text, level=0, style='List Bullet'):
    """
    添加列表项

    参数:
    - level: 缩进级别
    - style: 列表样式('List Bullet'或'List Number')
    """
    doc = Document(doc_path)
    paragraph = doc.add_paragraph(style=style)
    paragraph.text = text
    paragraph.paragraph_format.left_indent = Pt(level * 18)
    doc.save(doc_path)


def delete_list_item(doc_path, item_index):
    """
    删除指定的列表项及其所有子项
    """
    doc = Document(doc_path)
    count = 0
    to_delete = []
    current_level = None

    for i, paragraph in enumerate(doc.paragraphs):
        if paragraph.style.name.startswith('List'):
            if count == item_index:
                current_level = paragraph.paragraph_format.left_indent
                to_delete.append(i)
            elif current_level is not None:
                if paragraph.paragraph_format.left_indent > current_level:
                    to_delete.append(i)
                else:
                    break
            count += 1

    for i in reversed(to_delete):
        p = doc.paragraphs[i]._element
        p.getparent().remove(p)

    doc.save(doc_path)


def set_column_width(table, col_idx, width_in_inches):
    """
    设置表格某列的宽度。

    参数：
        table (Table): 表格对象
        col_idx (int): 列索引
        width_in_inches (float): 列宽以英寸为单位）
    """
    for cell in table.columns[col_idx].cells:
        cell.width = Inches(width_in_inches)


def delete_table_row(doc_path, table_index, row_index):
    """
    删除表格中的指定行
    """
    doc = Document(doc_path)
    if 0 <= table_index < len(doc.tables):
        table = doc.tables[table_index]
        if 0 <= row_index < len(table.rows):
            tr = table.rows[row_index]._element
            tr.getparent().remove(tr)
    doc.save(doc_path)


def add_table_column(doc_path, table_index):
    """
    向指定表格添加列
    """
    doc = Document(doc_path)
    if 0 <= table_index < len(doc.tables):
        table = doc.tables[table_index]
        for row in table.rows:
            cell = row._tr.add_tc()
            table.cell(row._index, len(row.cells) - 1)._tc = cell
            cell = table.cell(row._index, len(row.cells) - 1)
            cell.width = Inches(1)  # 设置新列的宽度，可以根据需要调整
    doc.save(doc_path)


def set_row_height(table, row_idx, height_in_inches):
    """
    设置表格某行的高度。

    参数：
        table (Table): 表格对象
        row_idx (int): 行索引
        height_in_inches (float): 行高（以英寸为单位）
    """
    table.rows[row_idx].height = Inches(height_in_inches)


def merge_cells(table, start_row, start_col, end_row, end_col):
    """
    合并表格的单元格。

    参数：
        table (Table): 表格对象
        start_row (int): 起始行索引
        start_col (int): 起始列索引
        end_row (int): 结束行索引
        end_col (int): 结束列索引
    """
    cell = table.cell(start_row, start_col)
    other_cell = table.cell(end_row, end_col)
    cell.merge(other_cell)


def set_cell_bg_color(table, row, col, color_hex):
    """
    设置单元格背景颜色。

    参数：
        table (Table): 表格对象
        row (int): 行索引
        col (int): 列索引
        color_hex (str): 颜色的十六进制代码（例如，'FF5733' 表示橙色）
    """
    cell = table.cell(row, col)
    shading_elm = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), color_hex))
    cell._element.get_or_add_tcPr().append(shading_elm)


def align_cell_text(table, row, col, alignment=WD_ALIGN_PARAGRAPH.CENTER):
    """
    设置单元格文本的对齐方式。

    参数：
        table (Table): 表格对象
        row (int): 行索引
        col (int): 列索引
        alignment (WD_ALIGN_PARAGRAPH): 对齐方式（如：WD_ALIGN_PARAGRAPH.CENTER、WD_ALIGN_PARAGRAPH.LEFT、WD_ALIGN_PARAGRAPH.RIGHT）
    """
    cell = table.cell(row, col)
    for paragraph in cell.paragraphs:
        paragraph.alignment = alignment


# 注意：在调用set_three_line_table函数之前，确保表格已经添加了所有行和单元格内容。
if __name__ == "__main__":
    # # 创建一个表格
    # # 创建一个表格
    # table = add_table(doc, 3, 3)
    #
    # # 设置三线表的边框
    # set_table_borders(table, border_color="FF0000", border_width=2, is_threeline=True)
    #
    # # 保存文档
    # doc.save("three_line_table1.docx")
    pass


# picture


# 替换图片并调整大小的函数
def replace_picture(doc, old_picture_path, new_picture_path, width_inch=None, height_inch=None):
    """
    替换文档中的旧图片为新图片，并调整图片的大小。
    :param doc: 文档对象
    :param old_picture_path: 旧图片的路径
    :param new_picture_path: 新图片的路径
    :param width_inch: 新图片的宽度，单位为英寸（可选）
    :param height_inch: 新图片的高度，单位为英寸（可选）
    """
    # 遍历文档中的段落，查找并删除旧图片
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            if run._r.pic:  # 检查当前run是否包含图片
                # 检查图片路径是否与旧图片路径匹配
                if old_picture_path in str(run._r.pic):
                    # 如果找到匹配的图片，清空该段落的内容（删除旧图片）
                    paragraph.clear()  # 清除该段落内容
                    break

    # 检查新图片是否存在
    if os.path.exists(new_picture_path):
        # 插入新图片
        picture = doc.add_picture(new_picture_path)

        # 如果指定了宽度，调整图片宽度
        if width_inch:
            picture.width = Inches(width_inch)

        # 如果指定了高度，调整图片高度
        if height_inch:
            picture.height = Inches(height_inch)
    else:
        print(f"错误: 找不到路为 {new_picture_path} 的文件。")


def add_image(image_path, width=None, height=None):
    """添加图片"""
    global doc, current_paragraph
    try:
        # 打印当前工作目录
        current_working_dir = os.getcwd()
        logger.info(f"Current working directory: {current_working_dir}")
        
        # 检查图片路径
        if not os.path.isabs(image_path):
            # 获取当前脚本的目录
            current_dir = os.path.dirname(os.path.abspath(__file__))
            project_root = os.path.dirname(current_dir)
            logger.info(f"Script directory: {current_dir}")
            logger.info(f"Project root: {project_root}")
            
            # 尝试在多个位置查找图片
            possible_paths = [
                os.path.join(project_root, image_path),
                os.path.join(project_root, 'images', image_path),
                os.path.join(project_root, 'resources', image_path),
                os.path.join(project_root, 'assets', image_path),
                image_path,
                os.path.join('images', image_path),
                os.path.join('resources', image_path),
                os.path.join('assets', image_path)
            ]
            
            # 检查每个可能的路径
            logger.info("Checking possible paths:")
            for path in possible_paths:
                logger.info(f"Checking path: {path} - Exists: {os.path.exists(path)}")
                if os.path.exists(path):
                    image_path = path
                    logger.info(f"Found image at: {path}")
                    break
            else:
                raise FileNotFoundError(f"Image not found in any of these locations: {possible_paths}")
        
        # 创建新段落并添加图片
        current_paragraph = doc.add_paragraph()
        if width and height:
            image = current_paragraph.add_run().add_picture(image_path, width=Inches(width), height=Inches(height))
        elif width:
            image = current_paragraph.add_run().add_picture(image_path, width=Inches(width))
        elif height:
            image = current_paragraph.add_run().add_picture(image_path, height=Inches(height))
        else:
            image = current_paragraph.add_run().add_picture(image_path)
            
        return image, current_paragraph
    except Exception as e:
        logger.error(f"Error adding image: {e}")
        logger.error(f"Image path attempted: {image_path}")
        return None, None


def delete_image(doc_path, paragraph_index, image_index):
    """
    删除指定段落中的图片

    参数:
    - paragraph_index: 段落索引
    - image_index: 图片在段落中的索引
    """
    doc = Document(doc_path)
    if 0 <= paragraph_index < len(doc.paragraphs):
        paragraph = doc.paragraphs[paragraph_index]
        images = paragraph._element.findall('.//w:drawing', namespaces={
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
        if 0 <= image_index < len(images):
            image = images[image_index]
            parent = image.getparent()
            if parent is not None:
                parent.remove(image)
                # 如果 run 为空，则删除该 run
                if not parent.getchildren():
                    parent.getparent().remove(parent)

    doc.save(doc_path)


def set_image_size(image, width=None, height=None):
    """
    设置图片大小。

    :param image: 图片对象（在插入图片后返回的对象）
    :param width: 图片宽度（单位：英寸）
    :param height: 图片高度（单位：英）
    """
    if width:
        image.width = Inches(width)
    if height:
        image.height = Inches(height)


def align_image_left():
    """图片左对齐"""
    global current_paragraph
    try:
        if current_paragraph is None:
            logger.error("No paragraph selected")
            return None
        current_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        return current_paragraph
    except Exception as e:
        logger.error(f"Error aligning image left: {e}")
        return None


def align_image_right():
    """图片右对齐"""
    global current_paragraph
    try:
        if current_paragraph is None:
            logger.error("No paragraph selected")
            return None
        current_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        return current_paragraph
    except Exception as e:
        logger.error(f"Error aligning image right: {e}")
        return None


def align_image_center():
    """图片居中对齐"""
    global current_paragraph
    try:
        if current_paragraph is None:
            logger.error("No paragraph selected")
            return None
        current_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        return current_paragraph
    except Exception as e:
        logger.error(f"Error aligning image center: {e}")
        return None


def add_caption(doc, text, font_size=10):
    """
    为图片添加文字说明。

    :param doc: Document ���象
    :param text: 图片说明文字
    :param font_size: 字体大小（单位：Pt）
    """
    caption_paragraph = doc.add_paragraph()
    run = caption_paragraph.add_run(text)
    run.font.size = Pt(font_size)


# # 示例操作：插入图片并调整大小和对齐
# image_path = '/Users/ywootae/Desktop/研/半身.jpg'
# image_paragraph = doc.add_paragraph()
# image = insert_image(doc, image_path, width=2, height=2)  # 插入图片并设置大小
# align_image_center(image_paragraph)  # 居中对齐图片
# add_caption(doc, "图片说明文本", font_size=10)  # 添加说明文字
#
# # 保存文档
# doc.save('output.docx')

# # 替换图片例使用
# new_picture_path = "/Users/ywootae/Desktop/研/半身.jpg"  # 这里替换为实际的旧图片路径
# old_picture_path = "/Users/ywootae/Desktop/研/冠.jpg"  # 这里替换为实际的新图片路径
#
# # 替换图片并将新图片大小调整为 3 英宽，2 英寸高
# replace_picture(doc, old_picture_path, new_picture_path, width_inch=3, height_inch=2)
#
# # 保存修改后的文档
# doc.save("output_with_resized_picture.docx")


# chart

# 1. 创建折线图
def create_line_chart(x, y, title="Line Chart", xlabel="X Axis", ylabel="Y Axis", save_path='line_chart.png'):
    plt.figure(figsize=(8, 5))
    plt.plot(x, y, marker='o', linestyle='-', color='b', label='Line')
    plt.title(title)
    plt.xlabel(xlabel)
    plt.ylabel(ylabel)
    plt.legend()
    plt.grid(True)
    plt.savefig(save_path)
    plt.close()
    print(f"Line chart created and saved as '{save_path}'")


# 2. 创建柱状图
def create_bar_chart(x, y, title="Bar Chart", xlabel="X Axis", ylabel="Y Axis", save_path='bar_chart.png'):
    plt.figure(figsize=(8, 5))
    plt.bar(x, y, color='g', label='Bar')
    plt.title(title)
    plt.xlabel(xlabel)
    plt.ylabel(ylabel)
    plt.legend()
    plt.grid(True)
    plt.savefig(save_path)
    plt.close()
    print(f"Bar chart created and saved as '{save_path}'")


# 3. 创建饼图
def create_pie_chart(labels, sizes, title="Pie Chart", save_path='pie_chart.png'):
    plt.figure(figsize=(8, 8))
    plt.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=90, colors=['#ff9999', '#66b3ff', '#99ff99', '#ffcc99'])
    plt.title(title)
    plt.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.
    plt.savefig(save_path)
    plt.close()
    print(f"Pie chart created and saved as '{save_path}'")


# 4. 将图片插入到 Word 文档中
def insert_image_into_docx(image_path, docx_path):
    """
    将图片插入到 Word 文档中。

    参数：
        image_path (str): 图片的路径
        docx_path (str): 要保存的 Word 文档路径
    """
    # 创建一个新的 Word 文档
    doc = Document()

    # 插入图表图片
    doc.add_picture(image_path, width=Inches(6), height=Inches(4))  # 调整图片大小

    # 保存文档
    doc.save(docx_path)
    print(f"Word document saved as {docx_path}")


# 5. 主函数
def generate_chart(chart_type, data, chart_title="Chart", docx_path='output.docx'):
    # 解析数据
    if chart_type == "line":
        x, y = data
        create_line_chart(x, y, title=chart_title, save_path='line_chart.png')
        insert_image_into_docx('line_chart.png', docx_path)
    elif chart_type == "bar":
        x, y = data
        create_bar_chart(x, y, title=chart_title, save_path='bar_chart.png')
        insert_image_into_docx('bar_chart.png', docx_path)
    elif chart_type == "pie":
        labels, sizes = data
        create_pie_chart(labels, sizes, title=chart_title, save_path='pie_chart.png')
        insert_image_into_docx('pie_chart.png', docx_path)
    else:
        print("Unsupported chart type!")

# # 示例调用
# chart_data = {
#     "line": ([1, 2, 3, 4], [19.2, 21.4, 16.7, 18.5]),
#     "bar": (['A', 'B', 'C'], [12, 17, 8]),
#     "pie": (['Red', 'Blue', 'Green'], [40, 30, 30])
# }
#
# # 生成折线图并插入到 Word 文档
# generate_chart("line", chart_data["line"], "Line Chart Example", 'line_chart_output.docx')
#
# # 生成柱状图并插入到 Word 文档
# generate_chart("bar", chart_data["bar"], "Bar Chart Example", 'bar_chart_output.docx')
#
# # 生成饼图并插入到 Word 文档
# generate_chart("pie", chart_data["pie"], "Pie Chart Example", 'pie_chart_output.docx')

def add_heading(text, level=1):
    """添加标题
    
    Args:
        text: 标题文本
        level: 标题级别(1-9)
    
    Returns:
        新添加的标题段落对象
    """
    global doc, current_heading
    try:
        current_heading = doc.add_heading(text, level=level)
        return current_heading
    except Exception as e:
        logger.error(f"Error adding heading: {e}")
        return None

def add_footer(text):
    """添加页脚
    
    Args:
        text: 页脚文本
    """
    global doc
    try:
        section = doc.sections[0]
        footer = section.footer
        paragraph = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        paragraph.text = text
        return paragraph
    except Exception as e:
        logger.error(f"Error adding footer: {e}")
        return None

def add_header(text):
    """添加页眉
    
    Args:
        text: 页眉文本
    """
    global doc
    try:
        section = doc.sections[0]
        header = section.header
        paragraph = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
        paragraph.text = text
        return paragraph
    except Exception as e:
        logger.error(f"Error adding header: {e}")
        return None

def add_table(rows, cols, style=None):
    """添加表格
    
    Args:
        rows: 行数
        cols: 列数
        style: 表格样式(可选)
    
    Returns:
        新添加的表格对象
    """
    global doc, current_table
    try:
        current_table = doc.add_table(rows=rows, cols=cols, style=style)
        return current_table
    except Exception as e:
        logger.error(f"Error adding table: {e}")
        return None

def set_cell_text(row, col, text):
    """设置表格单元格文本
    
    Args:
        row: 行索引
        col: 列索引
        text: 单元格文本
    """
    global current_table
    try:
        if current_table is None:
            logger.error("No table selected")
            return None
        cell = current_table.cell(row, col)
        cell.text = text
        return cell
    except Exception as e:
        logger.error(f"Error setting cell text: {e}")
        return None

def set_table_style(style_name):
    """设置表格样式
    
    Args:
        style_name: 样式名称
    """
    global current_table
    try:
        if current_table is None:
            logger.error("No table selected")
            return None
        current_table.style = style_name
        return current_table
    except Exception as e:
        logger.error(f"Error setting table style: {e}")
        return None

def add_shape(shape_type):
    """添加形状
    
    Args:
        shape_type: 形状类型 ('rectangle', 'oval', 'line', etc.)
    
    Returns:
        新添加的形状对象
    """
    global doc, current_paragraph
    try:
        current_paragraph = doc.add_paragraph()
        shape = current_paragraph._element.add_shape(
            shape_type,
            width=Inches(1),
            height=Inches(1)
        )
        return shape
    except Exception as e:
        logger.error(f"Error adding shape: {e}")
        return None

def set_fill_color(color):
    """设置形状填充颜色
    
    Args:
        color: 颜色名称或RGB值
    """
    global current_paragraph
    try:
        if current_paragraph is None:
            logger.error("No shape selected")
            return None
        
        # 将颜色名称转换为RGB值
        color_map = {
            'red': (255, 0, 0),
            'green': (0, 255, 0),
            'blue': (0, 0, 255),
            'yellow': (255, 255, 0),
            'purple': (128, 0, 128),
            'orange': (255, 165, 0)
        }
        
        rgb = color_map.get(color.lower(), color)
        if isinstance(rgb, str):
            # 如果是十六进制颜色值
            rgb = tuple(int(rgb.lstrip('#')[i:i+2], 16) for i in (0, 2, 4))
            
        shape = current_paragraph._element.find('.//v:shape', namespaces={'v': 'urn:schemas-microsoft-com:vml'})
        if shape is not None:
            shape.set('fillcolor', f'#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}')
            return True
    except Exception as e:
        logger.error(f"Error setting fill color: {e}")
    return None

def set_background_color(color):
    """设置文档背景颜色
    
    Args:
        color: 颜色名称或RGB值
    """
    global doc
    try:
        # 颜色映射表
        color_map = {
            'red': (255, 0, 0),
            'green': (0, 255, 0),
            'blue': (0, 0, 255),
            'yellow': (255, 255, 0),
            'purple': (128, 0, 128),
            'orange': (255, 165, 0),
            'light gray': (211, 211, 211),
            'light blue': (173, 216, 230),
            'light green': (144, 238, 144)
        }
        
        # 获取RGB值
        rgb = color_map.get(color.lower(), color)
        if isinstance(rgb, str):
            # 如果是十六进制颜色值
            rgb = tuple(int(rgb.lstrip('#')[i:i+2], 16) for i in (0, 2, 4))
            
        # 设置背景色
        background = doc.sections[0]._sectPr.xpath("./w:background")
        if not background:
            background = OxmlElement('w:background')
            doc.sections[0]._sectPr.append(background)
        background.set(qn('w:color'), f'{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}')
        return True
        
    except Exception as e:
        logger.error(f"Error setting background color: {e}")
        return None