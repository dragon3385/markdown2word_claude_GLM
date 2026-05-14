# 文档样式配置
# 本配置文件定义了Word文档的样式规范，包括页面布局、字体样式和表格格式

# 页面设置（单位：厘米）
# 采用A4纸张标准尺寸，并设置页边距
PAGE_CONFIG = {
    'height': 29.7,  # A4纸高度（ISO 216标准）
    'width': 21.0,   # A4纸宽度（ISO 216标准）
    'margin_left': 2.8,   # 左页边距
    'margin_right': 2.6,  # 右页边距
    'margin_top': 3.5,    # 上页边距
    'margin_bottom': 3.5  # 下页边距
}

# 字体设置
# 定义文档中不同层级标题和正文的字体样式
# 包括字体名称、大小、对齐方式等属性
FONT_CONFIG = {
    'title': {  # 文档标题样式
        'name': '方正小标宋_GBK',  # 字体名称（需确保系统已安装该字体）
        'size': 22,               # 字号（单位：磅）
        'alignment': 'center'      # 对齐方式：居中
    },
    'heading1': {  # 一级标题样式
        'name': '方正黑体_GBK',    # 字体名称
        'size': 16                # 字号（单位：磅）
    },
    'heading2': {  # 二级标题样式
        'name': '方正楷体_GBK',    # 字体名称
        'size': 16                # 字号（单位：磅）
    },
    'heading3': {  # 三级标题样式
        'name': '方正仿宋_GBK',    # 字体名称
        'size': 16                # 字号（单位：磅）
    },
    'heading4': {  # 四级标题样式
        'name': '方正仿宋_GBK',    # 字体名称
        'size': 16                # 字号（单位：磅）
    },
    'body': {  # 正文样式
        'name': '方正仿宋_GBK',    # 字体名称
        'size': 16,               # 字号（单位：磅）
        'line_spacing': 28,       # 行距（单位：磅）
        'first_line_indent': 32   # 首行缩进（单位：磅，值为字号的2倍）
    }
}

# 表格设置
# 定义表格的基本样式，包括对齐方式、边距和边框样式
TABLE_CONFIG = {
    'alignment': 'center',        # 表格对齐方式：居中
    'cell_margin': 0,            # 单元格内边距（单位：磅）
    'border_color': (0, 0, 0),   # 边框颜色：黑色（RGB值）
    'border_style': 'single'      # 边框样式：单实线
}