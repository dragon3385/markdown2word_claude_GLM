"""Markdown to Word CLI 入口。"""

import argparse
import sys
import os

from converter import convert_markdown_to_docx


def main():
    parser = argparse.ArgumentParser(
        description='将 Markdown 文件转换为 Word (.docx) 文档'
    )
    parser.add_argument(
        'input',
        help='输入 Markdown 文件路径'
    )
    parser.add_argument(
        '-o', '--output',
        help='输出 Word 文件路径（默认与输入同目录、同文件名、.docx 后缀）'
    )
    args = parser.parse_args()

    if not os.path.isfile(args.input):
        print(f'错误：文件不存在: {args.input}', file=sys.stderr)
        sys.exit(1)

    if not args.input.lower().endswith('.md'):
        print(f'警告：输入文件不是 .md 文件: {args.input}', file=sys.stderr)

    try:
        output = convert_markdown_to_docx(args.input, args.output)
        print(f'转换完成: {output}')
    except Exception as e:
        print(f'错误：转换失败: {e}', file=sys.stderr)
        sys.exit(1)


if __name__ == '__main__':
    main()
