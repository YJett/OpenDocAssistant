import argparse


class API(object):
    def __init__(self, name: str, parameters: str, description: str,
                 parameter_description: str = "", 
                 composition_instruction: str = "", 
                 example: str = "", 
                 api_desc: str = "",
                 type: str = "",
                 implementation = None) -> None:
        """初始化一个API文档对象。
        
        参数:
            name: API的名称
            parameters: 参数签名
            description: API的通用描述
            parameter_description: 参数的详细描述
            composition_instruction: API的使用说明
            example: 使用示例
            api_desc: 简短描述
            type: API类型
            implementation: 实现细节
        """
        self.name = name
        self.parameters = parameters  
        self.description = description
        self.parameter_description = parameter_description
        self.composition_instruction = composition_instruction
        self.example = example
        self.api_desc = api_desc

        # self.implementation = implementation
        # self.type = type

    def __str__(self):
        infos = [f"API: {self.name}{self.parameters}: {self.description}", self.parameter_description,
                 self.composition_instruction, self.example]
        infos = [item for item in infos if item != ""]
        return '\n'.join(infos)


# document
document_APIs = [
    API(name="create_docx", 
        parameters="()", 
        description="此API创建新的Word文档",
        api_desc="创建新的Word文档"
        ),

    API(
        name="rename_docx_name",
        parameters="(original_path, new_name)",
        description="此 API 重命名现有的 docx 文档，但不改变其路径。",
        parameter_description="original_path 是文档当前的路径，new_name 是新文件的名称。",
        composition_instruction="你应该调用 rename_docx_name() 来重命名一个已存在的文档。",
        api_desc="重命名文档"
    ),

    API(
        name="move_docx_to_new_path",
        parameters="(original_path, new_directory)",
        description="此 API 将 docx 文件移动到一个新目录，但不改变文件名。",
        parameter_description="original_path 是文档的当前路径，new_directory 是文档将要移动到的新目录。",
        composition_instruction="你应该调用 move_docx_to_new_path() 来将文档移动到新的目录。",
        api_desc="移动文档"
    ),

    API(
        name="find_paragraphs_by_heading_and_content",
        parameters="(doc, heading, content)",
        description="此 API 根据指定的标题和内容查找文档中的段落。",
        parameter_description="doc 是 Document 对象，heading 是标题文本，content 是段落内容。",
        composition_instruction="你应该调用 find_paragraphs_by_heading_and_content() 来查找文档中匹配指定标题和内容的段落。",
        api_desc="查找段落，搜索内容"
    ),

    API(
        name="add_header",
        parameters="(doc_path, text, section_index=None)",
        description="此 API 添加页眉。如果指定 section_index，则仅为该节添加页眉；否则为所有节添加页眉。",
        parameter_description="doc_path 是文档路径，text 是要添加的页眉文本，section_index 是可选的节索引。",
        composition_instruction="你应该调用 add_header() 来在文档中添加页眉。",
        api_desc="添加页眉"
    ),

    API(
        name="delete_header",
        parameters="(doc_path, section_index=None)",
        description="此 API 删除页眉。如果指定 section_index，则仅删除该节的页眉；否则删除所有节的页眉。",
        parameter_description="doc_path 是文档路径，section_index 是可选的节索引。",
        composition_instruction="你应该调用 delete_header() 来删除文档中的页眉。",
        api_desc="删除页眉"
    ),

    API(
        name="add_footer",
        parameters="(doc_path, text, section_index=None)",
        description="此 API 添加页脚。如果指定 section_index，则仅为该节添加页脚；否则为所有节添加页脚。",
        parameter_description="doc_path 是文档路径，text 是要添加的页脚文本，section_index 是可选的节索引。",
        composition_instruction="你应该调用 add_footer() 来在文档中添加页脚。",
        api_desc="添加页脚"
    ),

    API(
        name="delete_footer",
        parameters="(doc_path, section_index=None)",
        description="此 API 删除页脚。如果指定 section_index，则仅删除该节的页脚；否则删除所有节的页脚。",
        parameter_description="doc_path 是文档路径，section_index 是可选的节索引。",
        composition_instruction="你应该调用 delete_footer() 来删除文档中的页脚。",
        api_desc="删除页脚"
    ),

    API(
        name="add_line_break",
        parameters="(doc_path, paragraph_index, position=None)",
        description="此 API 在指定段落中添加换行符。",
        parameter_description="doc_path 是文档路径，paragraph_index 是段落索引，position 是段落中插入的位置（可选）。",
        composition_instruction="你应该调用 add_line_break() 在文档的段落中添加换行符。",
        api_desc="添加换行符"
    ),

    API(
        name="delete_line_break",
        parameters="(doc_path, paragraph_index, break_index)",
        description="此 API 删除指定段落中的换行符。",
        parameter_description="doc_path 是文档路径，paragraph_index 是段落索引，break_index 是换行符在段落中的索引。",
        composition_instruction="你应该调用 delete_line_break() 来删除段落中的换行符。",
        api_desc="删除换行符"
    ),

    API(
        name="add_heading",
        parameters="(doc_path, text, level, position=None)",
        description="此API用于添加标题",
        parameter_description=(
            "doc_path: 文档路径\n"
            "text: 标题文本\n"
            "level: 标题级别(1-9)\n"
            "position: 插入位置(可选)"
        ),
        composition_instruction=(
            "你应该调用 add_heading() 来添加标题。\n"
            "level参数决定标题的级别，1是最高级标题。\n"
            "如果需要在特定位置插入标题，可以通过position参数指定。"
        ),
        example="add_heading('doc.docx', '第一章', level=1)",
        api_desc="添加标题"
    ),

    API(
        name="add_style",
        parameters="(doc_path, style_name, font_name=None, font_size=None, bold=None, italic=None, color=None)",
        description="此 API 添加新样式。",
        parameter_description="doc_path 是文档路径，style_name 是样式名称，font_name 是字体名称，font_size 是字体大小，bold 和 italic 控制粗体和斜体，color 是颜色。",
        composition_instruction="你应该调用 add_style() 来为文档创建新的样式。",
        api_desc="添加样式"
    ),

    API(
        name="delete_style",
        parameters="(doc_path, style_name)",
        description="此 API 删除指定样式。",
        parameter_description="doc_path 是文档路径，style_name 是要删除的样式名称。",
        composition_instruction="你应该调用 delete_style() 来从文档中删除指定的样式。",
        api_desc="删除样式"
    ),

    API(
        name="add_page_numbers",
        parameters="(doc_path)",
        description="此 API 为文档添加页码。",
        parameter_description="doc_path 是文档路径。",
        composition_instruction="你应该调用 add_page_numbers() 来为文档添加页码。",
        api_desc="添加页码"
    ),

    API(
        name="delete_page_numbers",
        parameters="(doc_path)",
        description="此 API 删除文档的页码。",
        parameter_description="doc_path 是文档路径。",
        composition_instruction="你应该调用 delete_page_numbers() 来从文档中删除页码。",
        api_desc="删除页码"
    ),

    API(
        name="add_table_of_contents",
        parameters="(doc_path, levels=3)",
        description="此 API 添加目录，支持指定级别。",
        parameter_description="doc_path 是文档路径，levels 是目录级别（1-9）。",
        composition_instruction="你应该调用 add_table_of_contents() 来为文档添加目录。",
        api_desc="添加目录"
    ),

    API(
        name="delete_table_of_contents",
        parameters="(doc_path)",
        description="此 API 删除目录。",
        parameter_description="doc_path 是文档路径。",
        composition_instruction="你应该调用 delete_table_of_contents() 来删除文档中的目录。",
        api_desc="删除目录"
    ),

    API(
        name="add_watermark",
        parameters="(doc_path, text)",
        description="此 API 添加文字水印。",
        parameter_description="doc_path 是文档路径，text 是水印文本。",
        composition_instruction="你应该调用 add_watermark() 来为文档添加水印。",
        api_desc="添加水印"
    ),

    API(
        name="delete_watermark",
        parameters="(doc_path)",
        description="此 API 删除水印。",
        parameter_description="doc_path 是文档路径。",
        composition_instruction="你应该调用 delete_watermark() 来从文档中删除水印。",
        api_desc="删除水印"
    ),
    API(
        name="delete_paragraph",
        parameters="(doc_path, paragraph_index)",
        description="此 API 用于删除指定索引的段落。",
        parameter_description="doc_path 是文档的文件路径，paragraph_index 是要删除的段落的索引。",
        composition_instruction="你应该调用 delete_paragraph() 来从文档中删除指定索引的段落。",
        api_desc="删除指定段落"
    )
]

# text
text_APIs = [
    API(
        name="add_paragraph",
        parameters="(text, style=None)",
        description="此API用于添加段落并设置其样式",
        parameter_description="text: 段落文本内容\nstyle: 段落样式(可选)",
        composition_instruction="你应该调用 add_paragraph() 来添加一个新的段落。如果需要特定样式，可以通过style参数指定。",
        example="add_paragraph('这是一个新段落', style='Normal')",
        api_desc="添加段落"
    ),

    API(
        name="add_text",
        parameters="(doc_path, paragraph_index, text, position=None, bold=False, italic=False, underline=False, color=None, size=None)",
        description="此 API 用于在指定段落位置添加文本，可以设置文本样式。",
        parameter_description="doc_path 是文档路径，paragraph_index 是段落索引，text 是要添加的文本，position 是插入位置，bold, italic, underline 分别表示是否加粗、斜体、下划线，color 是颜色，size 是字体大小。",
        composition_instruction="你应该调用 add_text() 来向文档中的指定段落添加文本。",
        api_desc="添加文本"
    ),

    API(
        name="delete_text",
        parameters="(doc_path, paragraph_index, start, end)",
        description="此 API 用于删除指定段落中从 start 到 end 的文本。",
        parameter_description="doc_path 是文档路径，paragraph_index 是段落索引，start 和 end 分别为要删除的文本范围。",
        composition_instruction="你应该调用 delete_text() 来删除指定段落中的文本。",
        api_desc="删除文本"
    ),

    API(
        name="add_hyperlink",
        parameters="(doc_path, paragraph_index, text, url, position=None)",
        description="此 API 用于在指定段落添加超链接。",
        parameter_description="doc_path 是文档路径，paragraph_index 是段落索引，text 是超链接文本，url 是超链接地址，position 是插入位置。",
        composition_instruction="你应该调用 add_hyperlink() 来向指定段落添加超链接。",
        api_desc="添加超链接"
    ),

    API(
        name="delete_hyperlink",
        parameters="(doc_path, paragraph_index, link_index)",
        description="此 API 用于删除指定段落中的超链接。",
        parameter_description="doc_path 是文档路径，paragraph_index 是段落索引，link_index 是要删除的超链接索引。",
        composition_instruction="你应该调用 delete_hyperlink() 来删除指定段落中的超链接。",
        api_desc="删除超链接"
    ),

    API(
        name="delete_heading",
        parameters="(doc_path, level, occurrence)",
        description="此 API 用于删除指定级别的第 n 个标题。",
        parameter_description="doc_path 是文档路径，level 是标题级别，occurrence 是要删除的该级别标题的第几个。",
        composition_instruction="你应该调用 delete_heading() 来删除指定级别的标题。",
        api_desc="删除标题"
    ),

    API(
        name="add_table",
        parameters="(doc_path, rows, cols, position=None, style=None)",
        description="此 API 用于在文档中添加表格。",
        parameter_description="doc_path 是文档路径，rows 是行数，cols 是列数，position 是插入位置，style 是表格样式。",
        composition_instruction="你应该调用 add_table() 来在文档中添加表格。",
        api_desc="添加表格"
    ),

    API(
        name="set_font_size",
        parameters="(size)",
        description="此 API 用于设置当前段落字体大小。",
        parameter_description="size 是字体大小。",
        composition_instruction="你应该调用 set_font_size() 来设置当前段落的字体大小。",
        api_desc="设置字体大小"
    ),

    API(
        name="set_font_color",
        parameters="(color)",
        description="此 API 用于设置当前段落字体颜色。",
        parameter_description="color 是字体颜色代码。",
        composition_instruction="你应该调用 set_font_color() 来设置当前段落的字体颜色。",
        api_desc="设置字体颜色"
    ),

    API(
        name="set_font_bold",
        parameters="()",
        description="此 API 用于将当前段落的字体设置为加粗。",
        parameter_description="无参数。",
        composition_instruction="你应该调用 set_font_bold() 来将当前段落的字体加粗。",
        api_desc="设置加粗字体"
    ),

    API(
        name="set_font_italic",
        parameters="()",
        description="此 API 用于将当前段落的字体设置为斜体。",
        parameter_description="无参数。",
        composition_instruction="你应该调用 set_font_italic() 来将当前段落的字体设置为斜体。",
        api_desc="设置斜体"
    ),

    API(
        name="set_font_underline",
        parameters="()",
        description="此 API 用于在当前段落字体下添加下划线。",
        parameter_description="无参数。",
        composition_instruction="你应该调用 set_font_underline() 来为当前段落字体添加下划线。",
        api_desc="设置下划线"
    ),

    API(
        name="set_font_style",
        parameters="(font_name)",
        description="此 API 用于设置当前段落的字体样式。",
        parameter_description="font_name 是字体名称。",
        composition_instruction="你应该调用 set_font_style() 来设置当前段落的字体样式。",
        api_desc="设置字体样式"
    ),

    API(
        name="set_line_space",
        parameters="(line_space_level)",
        description="此 API 用于设置当前段落的行间距。",
        parameter_description="line_space_level 是行间距大小。",
        composition_instruction="你应该调用 set_line_space() 来设置当前段落的行间距。",
        api_desc="设置行间距"
    ),

    API(
        name="text_align_left",
        parameters="()",
        description="此 API 用于将当前段落左对齐。",
        parameter_description="无参数。",
        composition_instruction="你应该调用 text_align_left() 来将当前段落左对齐。",
        api_desc="左对齐"
    ),

    API(
        name="text_align_center",
        parameters="()",
        description="此 API 用于将当前段落居中对齐。",
        parameter_description="无参数。",
        composition_instruction="你应该调用 text_align_center() 来将当前段落居中对齐。",
        api_desc="居中对齐"
    ),

    API(
        name="text_align_right",
        parameters="()",
        description="此 API 用于将当前段落右对齐。",
        parameter_description="无参数。",
        composition_instruction="你应该调用 text_align_right() 来将当前段落右对齐。",
        api_desc="右对齐"
    ),

    API(
        name="add_run",
        parameters="(doc, text, font_size=12, font_color='000000', bold=False, italic=False, underline=False)",
        description="此 API 用于在当前段落中添加一个新的文本块，并设置文本的样式。",
        parameter_description="doc 是目标文档对象，text 是要添加的文本，font_size 是字体大小，font_color 是字体颜色，bold, italic, underline 分别表示是否加粗、斜体、下划线。",
        composition_instruction="你应该调用 add_run() 来在段落中添加一个新的文本块并设置样式。",
        api_desc="添加带样式的文本块"
    ),

    API(
        name="get_word",
        parameters="()",
        description="此 API 用于获取当前操作的 Word 文档对象。",
        parameter_description="无参数。",
        composition_instruction="你应该调用 get_word() 来获取当前的 Word 文档对象。",
        api_desc="获取当前文档对象"
    ),

    API(
        name="get_current_page_id",
        parameters="()",
        description="此 API 用于获取当前文档的页编号。",
        parameter_description="无参数。",
        composition_instruction="你应该调用 get_current_page_id() 来获取当前文档的页编号。",
        api_desc="获取当前页编号"
    ),


    API(
        name="set_word",
        parameters="(docx_path=None)",
        description="此 API 用于加载现有文档或初始化一个新的Word文档。",
        parameter_description="docx_path 是要加载的文档路径，如果未指定，则会初始化一个新文档。",
        composition_instruction="你应该调用 set_word() 来加载现有文档或初始化一个新文档。",
        api_desc="加载或初始化文档"
    ),

    API(
        name="set_current_paragraph",
        parameters="(index)",
        description="此 API 用于设置当前操作的段落。",
        parameter_description="index 是段落的索引。",
        composition_instruction="你应该调用 set_current_paragraph() 来设置当前操作的段落。",
        api_desc="设置当前段落"
    )

]

# table
table_APIs = [
    API(
        name="get_table_count",
        parameters="(doc)",
        description="此 API 用于获取文档中已存在的表格数量。",
        parameter_description="doc 是目标文档对象。",
        composition_instruction="你应该调用 get_table_count() 来获取文档中已有表格的数量。",
        api_desc="获取文档中表格的数量"
    ),

    API(
        name="add_table",
        parameters="(doc_path, rows, cols, position=None, style=None)",
        description="此 API 用于在文档中添加一个新的表格。",
        parameter_description="doc_path 是文档路径，rows 和 cols 分别为表格的行数和列数，position 是表格插入的位置，style 是表格样式。",
        composition_instruction="你应该调用 add_table() 来向文档中插入一个新的表格。",
        api_desc="在文档中添加表格"
    ),

    API(
        name="set_table_title",
        parameters="(doc, title, font_size=12, bold=True, color='black', alignment=WD_PARAGRAPH_ALIGNMENT.CENTER)",
        description="此 API 用于为表格添加标题，并设置标题样式。",
        parameter_description="doc 是目标文档对象，title 是标题文本，font_size 是字体大小，bold 表示是否加粗，color 是字体颜色，alignment 是标题的对齐方式。",
        composition_instruction="你应该调用 set_table_title() 来设置表格的标题并进行样式调整。",
        api_desc="设置表格标题并添加编号"
    ),

    API(
        name="delete_table",
        parameters="(doc_path, table_index)",
        description="此 API 用于删除指定索引的表格。",
        parameter_description="doc_path 是文档路径，table_index 是表格的索引。",
        composition_instruction="你应该调用 delete_table() 来删除指定的表格。",
        api_desc="删除指定索引的表格"
    ),

    API(
        name="add_table_header",
        parameters="(table, headers, font_size=12, bold=True, alignment=WD_ALIGN_PARAGRAPH.CENTER)",
        description="此 API 用于添加或修改表格的表头行。",
        parameter_description="table 是目标表格对象，headers 是表头文本的列表，font_size 是字体大小，bold 表示是否加粗，alignment 是对齐方式。",
        composition_instruction="你应该调用 add_table_header() 来为表格添加或修改表头行。",
        api_desc="添加或修改表格表头"
    ),

    API(
        name="set_cell_text",
        parameters="(table, row, col, text, font_size=12, bold=False)",
        description="此 API 用于设置表格单元格的文本内容和样式。",
        parameter_description="table 是目标表格对象，row 和 col 分别是单元格的行索引和列索引，text 是要设置的文本，font_size 是字体大小，bold 表示是否加粗。",
        composition_instruction="你应该调用 set_cell_text() 来修改表格单元格的文本内容和样式。",
        api_desc="设置表格单元格文本和样式"
    ),

    API(
        name="add_table_row",
        parameters="(doc_path, table_index)",
        description="此 API 用于向指定表格添加一行。",
        parameter_description="doc_path 是文档路径，table_index 是表格的索引。",
        composition_instruction="你应该调用 add_table_row() 来向指定表格中添加一行。",
        api_desc="向表格添加一行"
    ),

    API(
        name="add_table_column",
        parameters="(doc_path, table_index)",
        description="此 API 用于向指定表格添加一列。",
        parameter_description="doc_path 是文档路径，table_index 是表格的索引。",
        composition_instruction="你应该调用 add_table_column() 来向指定表格添加一列。",
        api_desc="向表格添加一列"
    ),

    API(
        name="delete_table_column",
        parameters="(doc_path, table_index, col_index)",
        description="此 API 用于删除表格中的指定列。",
        parameter_description="doc_path 是文档路径，table_index 是表格的索引，col_index 是列的索引。",
        composition_instruction="你应该调用 delete_table_column() 来删除指定列。",
        api_desc="删除指定列"
    ),

    API(
        name="add_list_item",
        parameters="(doc_path, text, level=0, style='List Bullet')",
        description="此 API 用于在文档中添加一个列表项。",
        parameter_description="doc_path 是文档路径，text 是列表项文本，level 是缩进级别，style 是列表样式。",
        composition_instruction="你应该调用 add_list_item() 来在文档中插入列表项。",
        api_desc="向文档添加列表项"
    ),

    API(
        name="delete_list_item",
        parameters="(doc_path, item_index)",
        description="此 API 用于删除指定的列表项及其所有子项。",
        parameter_description="doc_path 是文档路径，item_index 是要删除的列表项的索引。",
        composition_instruction="你应该调用 delete_list_item() 来删除指定的列表项。",
        api_desc="删除指定列表项及其子项"
    ),

    API(
        name="set_column_width",
        parameters="(table, col_idx, width_in_inches)",
        description="此 API 用于设置表格某列的宽度。",
        parameter_description="table 是目标表格对象，col_idx 是列索引，width_in_inches 是列宽（单位：英寸）。",
        composition_instruction="你应该调用 set_column_width() 来调整表格某列的宽度。",
        api_desc="设置表格列的宽度"
    ),

    API(
        name="delete_table_row",
        parameters="(doc_path, table_index, row_index)",
        description="此 API 用于删除表格中的指定行。",
        parameter_description="doc_path 是文档路径，table_index 是表格的索引，row_index 是行的索引。",
        composition_instruction="你应该调用 delete_table_row() 来删除指定的行。",
        api_desc="删除指定行"
    ),

    API(
        name="set_row_height",
        parameters="(table, row_idx, height_in_inches)",
        description="此 API 用于设置表格某行的高度。",
        parameter_description="table 是目标表格对象，row_idx 是行索引，height_in_inches 是行高（单位：英寸）。",
        composition_instruction="你应该调用 set_row_height() 来调整表格某行的高度。",
        api_desc="设置表格行的高度"
    ),

    API(
        name="merge_cells",
        parameters="(table, start_row, start_col, end_row, end_col)",
        description="此 API 用于合并表格中的指定单元格。",
        parameter_description="table 是目标表格对象，start_row, start_col 是起始单元格的行列索引，end_row, end_col 是结束单元格的行列索引。",
        composition_instruction="你应该调用 merge_cells() 来合并指定的单元格。",
        api_desc="合并表格单元格"
    ),

    API(
        name="set_cell_bg_color",
        parameters="(table, row, col, color_hex)",
        description="此 API 用于设置表格单元格的背景颜色。",
        parameter_description="table 是目标表格对象，row 和 col 是单元格的行列索引，color_hex 是背景颜色的十六进制代码。",
        composition_instruction="你应该调用 set_cell_bg_color() 来设置单元格的背景颜色。",
        api_desc="设置单元格背景颜色"
    ),

    API(
        name="align_cell_text",
        parameters="(table, row, col, alignment=WD_ALIGN_PARAGRAPH.CENTER)",
        description="此 API 用于设置表格单元格中文本的对齐方式。",
        parameter_description="table 是目标表格对象，row 和 col 是单元格的行列索引，alignment 是文本的对齐方式。",
        composition_instruction="你应该调用 align_cell_text() 来调整单元格中文本的对齐方式。",
        api_desc="设置单元格中文本的对齐方式"
    )

]

# picture
picture_APIs = [

    API(
        name="replace_picture",
        parameters="(doc, old_picture_path, new_picture_path, width_inch=None, height_inch=None)",
        description="此 API 用于替换文档中的旧图片为新图片，并可调整新图片的大小。",
        parameter_description="doc 是目标文档对象，old_picture_path 是旧图片的路径，new_picture_path 是新图片的路径，width_inch 是新图片的宽度（可选），height_inch 是新图片的高度（可选）。",
        composition_instruction="你应该调用 replace_picture() 来替换文档中的旧图片并调整新图片的大小。",
        api_desc="替换文档中的旧图片为新图片，并调整图片的大小"
    ),

    API(
        name="add_image",
        parameters="(doc_path, paragraph_index, image_path, width=None, height=None, position=None)",
        description="此 API 用于在指定段落中添加图片。",
        parameter_description="doc_path 是文档路径，paragraph_index 是段落的索引，image_path 是图片文件路径，width 是图片宽度（单位：英寸），height 是图片高度（单位：英寸），position 是图片插入的位置（可选）。",
        composition_instruction="你应该调用 add_image() 来在指定段落中插入图片。",
        api_desc="在指定段落中添加图片"
    ),

    API(
        name="delete_image",
        parameters="(doc_path, paragraph_index, image_index)",
        description="此 API 用于删除指定段落中的图片。",
        parameter_description="doc_path 是文档路径，paragraph_index 是段落的索引，image_index 是图片在段落中的索引。",
        composition_instruction="你应该调用 delete_image() 来删除指定段落中的图片。",
        api_desc="删除指定段落中的图片"
    ),

    API(
        name="set_image_size",
        parameters="(image, width=None, height=None)",
        description="此 API 用于设置插入图片的大小。",
        parameter_description="image 是插入图片后返回的图片对象，width 是图片宽度（单位：英寸），height 是图片高度（单位：英寸）。",
        composition_instruction="你应该调用 set_image_size() 来调整图片的大小。",
        api_desc="设置图片的大小"
    ),

    API(
        name="align_image_center",
        parameters="(paragraph)",
        description="此 API 用于将图片居中对齐。",
        parameter_description="paragraph 是包含图片的段落对象。",
        composition_instruction="你应该调用 align_image_center() 来将图片居中对齐。",
        api_desc="将图片居中对齐"
    ),

    API(
        name="align_image_right",
        parameters="(paragraph)",
        description="此 API 用于将图片右对齐。",
        parameter_description="paragraph 是包含图片的段落对象。",
        composition_instruction="你应该调用 align_image_right() 来将图片右对齐。",
        api_desc="将图片右对齐"
    ),

    API(
        name="add_caption",
        parameters="(doc, text, font_size=10)",
        description="此 API 用于为图片添加文字说明。",
        parameter_description="doc 是文档对象，text 是图片说明文字，font_size 是字体大小（单位：Pt）。",
        composition_instruction="你应该调用 add_caption() 来为图片添加说明文字。",
        api_desc="为图片添加文字说明"
    )

]

# chart
chart_APIs = [
    API(
        name="create_line_chart",
        parameters="(x, y, title='Line Chart', xlabel='X Axis', ylabel='Y Axis', save_path='line_chart.png')",
        description="此 API 用于创建并保存折线图。",
        parameter_description="x 是数据的横坐标，y 是数据的纵坐标，title 是图表标题，xlabel 和 ylabel 分别是 X 和 Y 轴的标签，save_path 是保存图表的文件路径。",
        composition_instruction="你应该调用 create_line_chart() 来生成并保存一个折线图。",
        api_desc="创建并保存折线图"
    ),

    API(
        name="create_bar_chart",
        parameters="(x, y, title='Bar Chart', xlabel='X Axis', ylabel='Y Axis', save_path='bar_chart.png')",
        description="此 API 用于创建并保存柱状图。",
        parameter_description="x 是数据的横坐标，y 是数据的纵坐标，title 是图表标题，xlabel 和 ylabel 分别是 X 和 Y 轴的标签，save_path 是保存图表的文件路径。",
        composition_instruction="你应该调用 create_bar_chart() 来生成并保存一个柱状图。",
        api_desc="创建并保存柱状图"
    ),

    API(
        name="create_pie_chart",
        parameters="(labels, sizes, title='Pie Chart', save_path='pie_chart.png')",
        description="此 API 用于创建并保存饼图。",
        parameter_description="labels 是饼图的标签，sizes 是饼图中每个部分的大小，title 是图表标题，save_path 是保存图表的文件路径。",
        composition_instruction="你应该调用 create_pie_chart() 来生成并保存一个饼图。",
        api_desc="创建并保存饼图"
    ),

    API(
        name="insert_image_into_docx",
        parameters="(image_path, docx_path)",
        description="此 API 用于将图片插入到 Word 文档中。",
        parameter_description="image_path 是要插入的图片路径，docx_path 是要保存的 Word 文档路径。",
        composition_instruction="你应该调用 insert_image_into_docx() 将图表图片插入到 Word 文档。",
        api_desc="将图片插入到 Word 文档中"
    ),

    API(
        name="generate_chart",
        parameters="(chart_type, data, chart_title='Chart', docx_path='output.docx')",
        description="此 API 根据指定的图表类型和数据生成图表，并将图表插入到 Word 文档中。",
        parameter_description="chart_type 是图表类型（'line'、'bar' 或 'pie'），data 是图表的数据，chart_title 是图表的标题，docx_path 是生成的 Word 文档保存路径。",
        composition_instruction="你应该调用 generate_chart() 来生成图表并将其插入到 Word 文档。",
        api_desc="根据指定类型生成图表并插入到 Word 文档中"
    )

]

# basic
basic_APIs = [
    API(name="set_width", parameters="(width)",
        description="此 API 用于设置所选对象的宽度。",
        parameter_description="该 API 接受一个参数 'width'，表示对象的宽度，单位为厘米（浮动值）。",
        composition_instruction="在更改对象宽度之前，请先选择一个对象。",
        api_desc="图片和形状的宽度"),

    API(name="set_height", parameters="(height)",
        description="此 API 用于设置所选对象的高度。",
        parameter_description="该 API 接受一个参数 'height'，表示对象的高度，单位为厘米（浮动值）。",
        composition_instruction="在更改对象高度之前，请先选择一个对象。",
        api_desc="图片和形状的高度"),

    API(name="rotate_element", parameters="(angle)",
        description="此 API 用于按指定角度旋转所选对象。",
        parameter_description="该 API 接受一个参数 'angle'，表示旋转角度，单位为整数，顺时针旋转。",
        composition_instruction="在旋转对象之前，请先选择一个对象。",
        api_desc="旋转"),

    API(name="set_fill_color", parameters="(color)",
        description="此 API 用于设置所选对象的填充颜色。",
        parameter_description="该 API 接受一个参数 'color'，表示要设置的填充颜色，可以是 '红色'、'紫色'、'蓝色'、'绿色'、'黄色' 或 '橙色'。",
        composition_instruction="可以设置内容、标题或文本框的填充颜色。",
        api_desc="填充颜色：红色，紫色，蓝色，绿色，黄色，橙色"),

    # API(name="align_top_right_corner", parameters="()",
    #     description="此 API 将所选对象移动到文档的右上角。",
    #     composition_instruction="在调用 align_top_right_corner() 之前，请先选择一个对象。"),
    #
    # API(name="align_top_left_corner", parameters="()",
    #     description="此 API 将所选对象移动到文档的左上角。",
    #     composition_instruction="在调用 align_top_left_corner() 之前，请先选择一个对象。"),
    #
    # API(name="align_bottom_right_corner", parameters="()",
    #     description="此 API 将所选对象移动到文档的右下角。",
    #     composition_instruction="在调用 align_bottom_right_corner() 之前，请先选择一个对象。"),
    #
    # API(name="align_bottom_left_corner", parameters="()",
    #     description="此 API 将所选对象移动到文档的左下角。",
    #     composition_instruction="在调用 align_bottom_left_corner() 之前，请先选择一个对象。"),
    #
    # API(name="align_slide_top", parameters="()",
    #     description="此 API 将所选对象移动到文档的顶部。",
    #     composition_instruction="在调用 align_slide_top() 之前，请先选择一个对象。"),
    #
    # API(name="align_slide_bottom", parameters="()",
    #     description="此 API 将所选对象移动到文档的底部。",
    #     composition_instruction="在调用 align_slide_bottom() 之前，请先选择一个对象。"),
    #
    # API(name="align_slide_left", parameters="()",
    #     description="此 API 将所选对象移动到文档的左侧。",
    #     composition_instruction="在调用 align_slide_left() 之前，请先选择一个对象。"),
    #
    # API(name="align_slide_right", parameters="()",
    #     description="此 API 将所选对象移动到文档的右侧。",
    #     composition_instruction="在调用 align_slide_right() 之前，请先选择一个对象。"),
    #
    # API(name="align_slide_center", parameters="()",
    #     description="此 API 将所选对象移动到文档的中心。",
    #     composition_instruction="在调用 align_slide_center() 之前，请先选择一个对象。"),

    # ??
    # API(name="set_left",
    #     parameters="(left)",
    #     description="此API用于移动和改变对象的位置。它设置选定对象的最左边点的x坐标。",
    #     parameter_description="它接受一个参数，即设置的x坐标。",
    #     composition_instruction="你应该先选择一个对象，然后才能更改它的最左边点位置。",
    #     api_desc="移动对象，设置左边、中间、右边的位置。"),
    #
    # API(name="set_top",
    #     parameters="(top)",
    #     description="此API用于移动和改变对象的位置。它设置选定对象的最上边点的y坐标。",
    #     parameter_description="它接受一个参数，即设置的y坐标。",
    #     composition_instruction="你应该先选择一个对象，然后才能更改它的最上边点位置。",
    #     api_desc="移动对象，设置顶部、中间、底部的位置。")

]
#
# # shape
# shape_APIs = [
#     API(name="insert_rectangle", parameters="()",
#         description="This API inserts a rectangle or square shape onto the slide.", api_desc="rectangle, square"),
#     API(name="insert_right_arrow", parameters="()", description="This API inserts an arrow shape onto the slide.",
#         api_desc="arrow"),
#     API(name="insert_rounded_rectangle", parameters="()",
#         description="This API inserts a rounded rectangle shape onto the slide.", api_desc="rounded rectangle"),
#     API(name="insert_triangle", parameters="()", description="This API inserts a triangle shape onto the slide.",
#         api_desc="triangle"),
#     API(name="insert_callout", parameters="()", description="This API inserts a callout shape onto the slide.",
#         api_desc="callout"),
#     API(name="insert_cloud", parameters="()", description="This API inserts a cloud shape onto the slide.",
#         api_desc="cloud"),
#     API(name="insert_star", parameters="()", description="This API inserts a star shape onto the current slide.",
#         api_desc="star"),
#     API(name="insert_circle", parameters="()",
#         description="This API inserts a circle or oval shape into the current slide.", api_desc="circle, oval"),
# ]

lack_APIs = [
    API(name="seek_assistance", parameters="()",
        description="This API requests human help when the computer is unsure about the result or lacks the necessary API to fulfill the user's instruction.",
        api_desc="assistance"),
]

import random

random.seed(42)


def random_permutation(lst):
    shuffled = lst.copy()
    random.shuffle(shuffled)
    return shuffled


def get_all_APIs(args):
    all_apis =  basic_APIs + text_APIs + picture_APIs+ table_APIs + chart_APIs
    if args.api_update:
        all_apis += update_APIs
        all_apis = random_permutation(all_apis)
    if args.api_lack:
        all_apis = [x for x in all_apis if x.name in original_apis]
        all_apis += lack_APIs
    return all_apis



def get_API_name(apis):
    return [api.name + api.parameters for api in apis]


def get_API_desc(apis):
    return [api.api_desc for api in apis]


def get_must_APIs(args):
    print(args)
    if args.dataset == 'long':
        # 长文档处理场景
        must_APIs = [
            document_APIs[0],  # create_and_save_docx - 创建和保存文档
            text_APIs[0],      # add_text - 添加文本
            text_APIs[4],      # add_heading - 添加标题
            table_APIs[0]      # get_table_count - 获取表格数量
        ]
    else:
        # 常规文档处理场景
        must_APIs = [ 
            document_APIs[0],  # create_and_save_docx - 创建和保存文档
            text_APIs[0],      # add_text - 添加文本
            text_APIs[4],      # add_heading - 添加标题
            picture_APIs[0],   # replace_picture - 替换图片
            table_APIs[0]      # get_table_count - 获取表格数量
        ]
    
    if args.api_lack:
        must_APIs += lack_APIs
    return must_APIs


def api_lack_mask(apis):
    ans = []  # 创建结果列表
    for api in apis:  # 遍历每个API调用
        # 从API调用中提取API名称（去掉参数部分）
        api_name = api.split('(')[0]  
        
        # 检查API是否在原始API列表中
        if not api_name in original_apis:
            # 如果API不存在，替换为seek_assistance()
            ans.append("seek_assistance()")  
        else:
            # 如果API存在，保持原样
            ans.append(api)
    return ans


# update

update_APIs = [
    # API(name="insert_icon", parameters="(icon_id)", description="This API inserts an icon onto the slide.",
    #     parameter_description="It takes one parameter 'icon_id', the ID or name of the icon to insert as a string.",
    #     api_desc="icon"),
    # API(name="insert_3d_model", parameters="(model_id)", description="This API inserts a 3D model onto the slide.",
    #     parameter_description="It takes one parameter 'model_id', the ID or name of the 3D model to insert as a string.",
    #     api_desc="3D model"),
    # API(name="insert_smart_art_list", parameters="(smart_art_type_id)",
    #     description="This API inserts a SmartArt list onto the slide.",
    #     parameter_description="It takes one parameter 'smart_art_type_id', the ID or name of the SmartArt list type to insert as a string.",
    #     api_desc="SmartArt list"),
    # API(name="insert_smart_art_process", parameters="(smart_art_type_id)",
    #     description="This API inserts a SmartArt process diagram onto the slide.",
    #     parameter_description="It takes one parameter 'smart_art_type_id', the ID or name of the SmartArt process type to insert as a string.",
    #     api_desc="SmartArt process"),
    # API(name="insert_smart_art_cycle", parameters="(smart_art_type_id)",
    #     description="This API inserts a SmartArt cycle diagram onto the slide.",
    #     parameter_description="It takes one parameter 'smart_art_type_id', the ID or name of the SmartArt cycle type to insert as a string.",
    #     api_desc="SmartArt cycle"),
    # API(name="insert_smart_art_pyramid", parameters="(smart_art_type_id)",
    #     description="This API inserts a SmartArt pyramid diagram onto the slide.",
    #     parameter_description="It takes one parameter 'smart_art_type_id', the ID or name of the SmartArt pyramid type to insert as a string.",
    #     api_desc="SmartArt pyramid"),
    # API(name="insert_smart_art_relationship", parameters="(smart_art_type_id)",
    #     description="This API inserts a SmartArt relationship diagram onto the slide.",
    #     parameter_description="It takes one parameter 'smart_art_type_id', the ID or name of the SmartArt relationship type to insert as a string.",
    #     api_desc="SmartArt relationship"),
    # API(name="insert_link", parameters="(link)", description="This API inserts a hyperlink onto the slide.",
    #     parameter_description="It takes one parameter 'link', the URL or path of the hyperlink as a string.",
    #     api_desc="hyperlink"),
    # API(name="insert_comment", parameters="(comment)", description="This API inserts a comment onto the slide.",
    #     parameter_description="It takes one parameter 'comment', the text of the comment to insert as a string.",
    #     api_desc="comment"),
    # API(name="insert_symbol", parameters="(symbol)", description="This API inserts a symbol onto the slide.",
    #     parameter_description="It takes one parameter 'symbol', the symbol character to insert as a string.",
    #     api_desc="symbol"),
    # API(name="insert_equation", parameters="(equation)", description="This API inserts an equation onto the slide.",
    #     parameter_description="It takes one parameter 'equation', the mathematical equation to insert as a string.",
    #     api_desc="equation"),
    # API(name="insert_audio", parameters="(url)", description="This API inserts an audio file onto the slide.",
    #     parameter_description="It takes one parameter 'url', the URL or path of the audio file to insert as a string.",
    #     api_desc="audio"),
    # API(name="insert_video", parameters="(url)", description="This API inserts a video onto the slide.",
    #     parameter_description="It takes one parameter 'url', the URL or path of the video file to insert as a string.",
    #     api_desc="video"),
    # API(name="insert_transition", parameters="(transition_id)",
    #     description="This API sets the slide transition effect for the current slide.",
    #     parameter_description="It takes one parameter 'transition_id', the ID or name of the transition effect to set as a string.",
    #     api_desc="slide transition"),
    # API(name="set_transition_duration", parameters="(time)",
    #     description="This API sets the duration of the slide transition effect.",
    #     parameter_description="It takes one parameter 'time', the duration of the transition effect in seconds as an integer or float.",
    #     api_desc="transition duration"),
    # API(name="set_transition_sound", parameters="(url)",
    #     description="This API sets the sound for the slide transition effect.",
    #     parameter_description="It takes one parameter 'url', the URL or path of the sound file to set for the transition as a string.",
    #     api_desc="transition sound"),
    # API(name="set_transition_after_time", parameters="(time)",
    #     description="This API sets the delay time before the slide transition starts.",
    #     parameter_description="It takes one parameter 'time', the delay time in seconds as an integer or float.",
    #     api_desc="transition delay"),
    # API(name="increase_transition_after_time", parameters="()",
    #     description="This API increases the delay time before the slide transition starts.",
    #     api_desc="increase transition delay"),
    # API(name="decrease_transition_after_time", parameters="()",
    #     description="This API decreases the delay time before the slide transition starts.",
    #     api_desc="decrease transition delay"),
    # API(name="increase_transition_duration", parameters="()",
    #     description="This API increases the duration of the slide transition effect.",
    #     api_desc="increase transition duration"),
    # API(name="decrease_transition_duration", parameters="()",
    #     description="This API decreases the duration of the slide transition effect.",
    #     api_desc="decrease transition duration"),
    # API(name="insert_entrance_animation", parameters="(animation_id)",
    #     description="This API adds an entrance animation to the selected object or text on the slide.",
    #     parameter_description="It takes one parameter 'animation_id', the ID or name of the entrance animation to apply as a string.",
    #     api_desc="entrance animation"),
    # API(name="insert_emphasis_animation", parameters="(animation_id)",
    #     description="This API adds an emphasis animation to the selected object or text on the slide.",
    #     parameter_description="It takes one parameter 'animation_id', the ID or name of the emphasis animation to apply as a string.",
    #     api_desc="emphasis animation"),
    # API(name="insert_exit_animation", parameters="(animation_id)",
    #     description="This API adds an exit animation to the selected object or text on the slide.",
    #     parameter_description="It takes one parameter 'animation_id', the ID or name of the exit animation to apply as a string.",
    #     api_desc="exit animation"),
    # API(name="insert_path_animation", parameters="(animation_id)",
    #     description="This API adds a motion path animation to the selected object on the slide.",
    #     parameter_description="It takes one parameter 'animation_id', the ID or name of the motion path animation to apply as a string.",
    #     api_desc="motion path animation"),
    # API(name="delete_animation", parameters="()",
    #     description="This API removes any animation applied to the selected object or text on the slide.",
    #     api_desc="remove animation"),
    # API(name="set_animation_start_time", parameters="(time)",
    #     description="This API sets the start time of the animation for the selected object or text on the slide.",
    #     parameter_description="It takes one parameter 'time', the start time of the animation in seconds as an integer or float.",
    #     api_desc="animation start time"),
    # API(name="set_animation_duration", parameters="(time)",
    #     description="This API sets the duration of the animation for the selected object or text on the slide.",
    #     parameter_description="It takes one parameter 'time', the duration of the animation in seconds as an integer or float.",
    #     api_desc="animation duration"),
    # API(name="start_record", parameters="()", description="This API starts recording the slide show.",
    #     api_desc="start recording slide show"),
    # API(name="end_record", parameters="()", description="This API stops recording the slide show.",
    #     api_desc="stop recording slide show"),
    # API(name="get_screenshot", parameters="()", description="This API captures a screenshot of the current slide.",
    #     api_desc="capture slide screenshot"),
    # API(name="set_font_crossline", parameters="()",
    #     description="This API adds a crossline to the selected font in the text on the slide.",
    #     api_desc="add font crossline"),
    # API(name="delete_comment", parameters="(comment_id)",
    #     description="This API deletes a specific comment from the slide.",
    #     parameter_description="It takes one parameter 'comment_id', the ID or name of the comment to delete as a string.",
    #     api_desc="delete comment"),
    # API(name="play_from_start", parameters="()",
    #     description="This API starts playing the slide show from the beginning.",
    #     api_desc="play slide show from start"),
    # API(name="play_from_current_slide", parameters="()",
    #     description="This API starts playing the slide show from the current slide.",
    #     api_desc="play slide show from current slide"),
    # API(name="play_from_slide", parameters="(slide_id)",
    #     description="This API starts playing the slide show from a specific slide.",
    #     parameter_description="It takes one parameter 'slide_id', the ID or number of the slide to start playing from as a string or integer.",
    #     api_desc="play slide show from specific slide"),
    # API(name="insert_hierarchy_chart", parameters="(data, series)",
    #     description="This API inserts a hierarchy chart onto the slide.",
    #     parameter_description="It takes two parameters, 'data' as a list of numbers and 'series' as a list of strings.",
    #     api_desc="hierarchy chart"),
    # API(name="insert_statistical_chart", parameters="(data, series)",
    #     description="This API inserts a statistical chart onto the slide.",
    #     parameter_description="It takes two parameters, 'data' as a list of numbers and 'series' as a list of strings.",
    #     api_desc="statistical chart"),
    # API(name="insert_scatter_chart", parameters="(data, series)",
    #     description="This API inserts a scatter chart onto the slide.",
    #     parameter_description="It takes two parameters, 'data' as a list of numbers and 'series' as a list of strings.",
    #     api_desc="scatter chart"),
    # API(name="insert_combo_chart", parameters="(data, series)",
    #     description="This API inserts a combo chart onto the slide.",
    #     parameter_description="It takes two parameters, 'data' as a list of numbers and 'series' as a list of strings.",
    #     api_desc="combo chart"),
    # API(name="insert_map", parameters="()", description="This API inserts a map shape onto the slide.", api_desc="map"),
    # API(name="insert_left_arrow", parameters="()", description="This API inserts a left arrow shape onto the slide.",
    #     api_desc="left arrow"),
    # API(name="insert_down_arrow", parameters="()", description="This API inserts a down arrow shape onto the slide.",
    #     api_desc="down arrow"),
    # API(name="insert_up_arrow", parameters="()", description="This API inserts an up arrow shape onto the slide.",
    #     api_desc="up arrow"),
    # API(name="insert_pentagon", parameters="()", description="This API inserts a pentagon shape onto the slide.",
    #     api_desc="pentagon"),
    # API(name="insert_trapezoid", parameters="()", description="This API inserts a trapezoid shape onto the slide.",
    #     api_desc="trapezoid"),
    # API(name="insert_smile_face", parameters="()", description="This API inserts a smiley face shape onto the slide.",
    #     api_desc="smiley face"),
    # API(name="insert_heart_shape", parameters="()", description="This API inserts a heart shape onto the slide.",
    #     api_desc="heart shape"),
    # API(name="insert_lightening_shape", parameters="()",
    #     description="This API inserts a lightning shape onto the slide.", api_desc="lightning shape"),
    # API(name="insert_stop_shape", parameters="()", description="This API inserts a stop sign shape onto the slide.",
    #     api_desc="stop sign shape"),
    # #
    # API(name="insert_flow_chart", parameters="(flow_chart_type)",
    #     description="This API inserts a flow chart shape onto the slide.",
    #     parameter_description="It takes one parameter 'flow_chart_type', the type of flow chart shape to insert as a string.",
    #     api_desc="flow chart"),
    # API(name="insert_moon", parameters="()", description="This API inserts a moon shape onto the slide.",
    #     api_desc="moon shape"),
    # API(name="insert_sun", parameters="()", description="This API inserts a sun shape onto the slide.",
    #     api_desc="sun shape"),
    # API(name="insert_ellipse", parameters="()", description="This API inserts an ellipse shape onto the slide.",
    #     api_desc="ellipse shape"),
    # API(name="group_shapes", parameters="(shape_ids)", description="This API groups multiple shapes on the slide.",
    #     parameter_description="It takes one parameter 'shape_ids', a list containing the IDs or names of the shapes to group.",
    #     api_desc="group shapes"),
    # API(name="ungroup_shapes", parameters="(group_id)",
    #     description="This API ungroups a set of grouped shapes on the slide.",
    #     parameter_description="It takes one parameter 'group_id', the ID or name of the shape group to ungroup.",
    #     api_desc="ungroup shapes"),
    # API(name="set_border_color", parameters="(shape_id, color)",
    #     description="This API sets the border color of a specific shape.",
    #     parameter_description="It takes two parameters, 'shape_id' as the ID or name of the shape and 'color' for the border color.",
    #     api_desc="set border color"),
    # API(name="lock_shape", parameters="(shape_id)",
    #     description="This API locks a specific shape, preventing it from being edited or moved.",
    #     parameter_description="It takes one parameter 'shape_id', the ID or name of the shape to lock.",
    #     api_desc="lock shape"),
    # API(name="unlock_shape", parameters="(shape_id)",
    #     description="This API unlocks a specific shape, allowing it to be edited or moved.",
    #     parameter_description="It takes one parameter 'shape_id', the ID or name of the shape to unlock.",
    #     api_desc="unlock shape"),
    # API(name="set_shape_opacity", parameters="(shape_id, opacity)",
    #     description="This API sets the opacity of a specific shape.",
    #     parameter_description="It takes two parameters, 'shape_id' as the ID or name of the shape and 'opacity' as a value between 0 (transparent) to 1 (opaque).",
    #     api_desc="set shape opacity"),
    # API(name="send_to_back", parameters="(shape_id)",
    #     description="This API sends a specific shape to the back of the slide.",
    #     parameter_description="It takes one parameter 'shape_id', the ID or name of the shape to send to the back.",
    #     api_desc="send shape to back"),
    # API(name="bring_to_front", parameters="(shape_id)",
    #     description="This API brings a specific shape to the front of the slide.",
    #     parameter_description="It takes one parameter 'shape_id', the ID or name of the shape to bring to the front.",
    #     api_desc="bring shape to front"),
    # API(name="distribute_horizontally", parameters="(shape_ids)",
    #     description="This API distributes shapes equally in horizontal spacing.",
    #     parameter_description="It takes one parameter 'shape_ids', a list containing the IDs or names of the shapes to distribute.",
    #     api_desc="distribute shapes horizontally"),
    # API(name="distribute_vertically", parameters="(shape_ids)",
    #     description="This API distributes shapes equally in vertical spacing.",
    #     parameter_description="It takes one parameter 'shape_ids', a list containing the IDs or names of the shapes to distribute.",
    #     api_desc="distribute shapes vertically"),
    # API(name="export_slide_as_image", parameters="(file_name, format)",
    #     description="This API exports the current slide as an image in the specified format.",
    #     parameter_description="It takes two parameters, 'file_name' as the name of the file to save the image to and 'format' as the desired image format (e.g., 'JPEG', 'PNG').",
    #     api_desc="export slide as image"),
    # API(name="set_shape_gradient", parameters="(shape_id, gradient_type, colors)",
    #     description="This API sets a gradient fill on a shape.",
    #     parameter_description="Takes 'shape_id' as the shape's ID or name, 'gradient_type' as the gradient's type (linear, radial, etc.), and 'colors' as a list of colors for the gradient.",
    #     api_desc="set shape gradient"),
    # API(name="merge_cells", parameters="(table_id, start_row, start_column, end_row, end_column)",
    #     description="Merges cells in a table.",
    #     parameter_description="Parameters specify the table and the range of cells to merge.",
    #     api_desc="merge table cells"),
    API(name="split_cell", parameters="(table_id, row, column)",
        description="拆分已合并的单元格。",
        parameter_description="接受 'table_id' 表示表格的 ID 或名称，以及 'row' 和 'column' 表示单元格的位置。",
        api_desc="拆分表格单元格"),

    API(name="set_table_style", parameters="(table_id, style)",
        description="为表格应用样式。",
        parameter_description="接受 'table_id' 表示表格的 ID 或名称，以及 'style' 表示样式名称。",
        api_desc="设置表格样式"),

    API(name="insert_hyperlink", parameters="(text, link)",
        description="在文档中插入超链接。",
        parameter_description="接受 'text' 作为可见文本，'link' 作为实际的超链接。",
        api_desc="插入超链接"),

    API(name="insert_image_gallery", parameters="(image_paths)",
        description="在文档中插入图像库。",
        parameter_description="接受 'image_paths' 作为图像路径的列表。",
        api_desc="插入图像库"),

    API(name="flip_shape", parameters="(shape_id, direction)",
        description="水平或垂直翻转形状。",
        parameter_description="接受 'shape_id' 表示形状的 ID 或名称，'direction' 为翻转方向，'horizontal' 表示水平，'vertical' 表示垂直。",
        api_desc="翻转形状"),

    API(name="insert_date_time", parameters="(format, position)",
        description="在文档中插入日期和时间。",
        parameter_description="接受 'format' 表示日期/时间格式，'position' 表示插入位置。",
        api_desc="插入日期和时间"),

    API(name="embed_fonts", parameters="()",
        description="将使用的字体嵌入到文档中。",
        api_desc="嵌入字体"),

    API(name="set_text_wrap", parameters="(shape_id, wrap_type)",
        description="设置文本在形状内部的环绕方式。",
        parameter_description="接受 'shape_id' 表示形状的 ID 或名称，'wrap_type' 为文本环绕类型，如 'square', 'tight' 等。",
        api_desc="设置文本环绕"),

    API(name="lock_aspect_ratio", parameters="(shape_id)",
        description="锁定形状的纵横比。",
        parameter_description="接受 'shape_id' 表示形状的 ID 或名称。",
        api_desc="锁定形状纵横比"),

    API(name="unlock_aspect_ratio", parameters="(shape_id)",
        description="解锁形状的纵横比。",
        parameter_description="接受 'shape_id' 表示形状的 ID 或名称。",
        api_desc="解锁形状纵横比")
]

update_apis = [x.name for x in update_APIs]

original_apis = [
    "choose_title",
    "choose_content",
    "choose_picture",
    "choose_shape",
    "choose_textbox",
    "set_background_color",
    "insert_text",
    "insert_picture",
    "insert_rectangle",
    "insert_right_arrow",
    "insert_line_chart",
    "insert_bar_chart",
    "insert_pie_chart",
    "choose_table",
    "choose_table_cell",
    "insert_table",
    "set_font_color",
    "set_font_size",
    "set_font_bold",
    "set_height",
    "set_width",
    "set_left",
    "set_top",
]

# api_selector
# - get_all_APIs
# - get_selected_APIs
# api_executor

# api update:


# api lack:
