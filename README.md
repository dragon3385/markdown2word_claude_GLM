# Markdown to Word 转换工具

将 Markdown 文件转换为符合中国公文排版规范的 Word (.docx) 文档。

## 功能特性

- 将 Markdown 标题层级（h1-h5）转换为对应格式的 Word 标题，并自动添加公文序号
- 支持正文、粗体、斜体、删除线、行内代码等行内格式
- 支持有序列表、无序列表
- 支持表格（自动边框、居中）
- 支持本地图片和网络图片
- 支持分隔线、代码块、引用块

## 快速开始

### 环境要求

- Python 3.8+
- Windows / macOS / Linux

### 安装

```bash
# 克隆项目
git clone <repository-url>
cd markdown2word_claude_GLM

# 创建虚拟环境（推荐）
python -m venv .venv

# 激活虚拟环境
# Windows:
.venv\Scripts\activate
# macOS / Linux:
source .venv/bin/activate

# 安装依赖
pip install -r requirements.txt
```

### 基本用法

```bash
# 最简用法：输入 .md 文件，自动生成同名 .docx 文件
python cli.py document.md

# 指定输出路径
python cli.py document.md -o output.docx
```

### 参数说明

| 参数 | 说明 |
|------|------|
| `input` | 输入 Markdown 文件路径（必需） |
| `-o, --output` | 输出 Word 文件路径（可选，默认与输入同名同目录） |

## 格式规范

生成的 Word 文档遵循以下排版规范：

### 页面设置

- 纸张：A4（210mm x 297mm）
- 左边距：2.8cm，右边距：2.6cm
- 上下边距：各 3.5cm

### 标题层级

| Markdown | 字体 | 字号 | 序号格式 |
|----------|------|------|----------|
| `#` (h1) | 方正小标宋_GBK | 22pt | 无（文档标题，居中） |
| `##` (h2) | 方正黑体_GBK | 16pt | 一、二、三、 |
| `###` (h3) | 方正楷体_GBK | 16pt | （一）（二）（三） |
| `####` (h4) | 方正仿宋_GBK | 16pt | 1. 2. 3. |
| `#####` (h5) | 方正仿宋_GBK | 16pt | （1）（2）（3） |

- h2-h5 标题具有 28pt 固定行距、首行缩进 2 字符
- 若 Markdown 中标题已包含序号（如 `## 一、工作目标`），自动检测并跳过重复编号

### 正文

- 字体：方正仿宋_GBK，16pt
- 行距：28pt（固定值）
- 首行缩进：2 字符（32pt）

### 列表

- 有序列表：`1. 2. 3.` 格式
- 无序列表：`●` 格式
- 列表项同样具有首行缩进和 28pt 行距

### 表格

- 黑色单实线边框
- 表格居中
- 单元格内无首行缩进
- 表头单元格自动加粗

### 图片

- 支持本地文件路径（相对于 Markdown 文件所在目录）
- 支持 URL 图片（自动下载，用后清理临时文件）
- 图片居中显示，最大宽度 6 英寸
- 行距为 1 倍，确保图片完整显示

## 示例

以下 Markdown 内容：

```markdown
# 关于开展年度工作总结的通知

## 一、工作目标

各部门要认真总结工作完成情况，主要包括以下方面：

### （一）具体要求

1. 总结报告字数不少于3000字
2. 数据要真实准确

### 时间安排

| 阶段 | 时间 | 内容 |
|------|------|------|
| 第一阶段 | 12月1日-10日 | 个人总结 |
```

将转换为一份格式规范的 Word 文档，其中：
- "关于开展年度工作总结的通知" 以方正小标宋 22pt 居中显示
- "一、工作目标" 以方正黑体 16pt 显示（检测到已有编号，不重复添加）
- "（一）具体要求" 以方正楷体 16pt 显示（检测到已有编号）
- "时间安排" 以方正楷体 16pt 显示，自动添加 "（二）" 编号
- 表格带有黑色边框、居中对齐

## 项目结构

```
markdown2word_claude_GLM/
  config.py        样式配置（页面、字体、表格）
  styles.py        样式桥接层（配置 → python-docx）
  converter.py     核心转换引擎（Markdown AST → Word）
  cli.py           命令行入口
  requirements.txt 依赖声明
  .venv/           虚拟环境（本地）
  tests/           测试文件
    samples/
      example.md   测试用 Markdown 样本
      tt.png       测试用图片
```

## 配置说明

样式配置集中在 `config.py` 中，可按需修改：

- `PAGE_CONFIG` — 页面尺寸和边距
- `FONT_CONFIG` — 各级标题和正文的字体、字号、行距、缩进
- `TABLE_CONFIG` — 表格边框和对齐方式

修改 `config.py` 后重新运行转换即可生效，无需修改代码。

## 依赖

| 库 | 用途 |
|----|------|
| [mistune](https://mistune.lepture.com/) 3.x | Markdown 解析 |
| [python-docx](https://python-docx.readthedocs.io/) 1.x | Word 文档生成 |
| [requests](https://docs.python-requests.org/) 2.x | 网络图片下载 |

## 注意事项

- 字体（方正小标宋_GBK、方正黑体_GBK 等）需要在打开文档的电脑上已安装
- 图片使用相对路径时，基于 Markdown 文件所在目录解析
- 文件编码默认 UTF-8
