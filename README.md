# 问卷回答整理工具

## 简介
这是一个用于整理问卷回答的Python工具，支持Excel（`.xlsx`, `.xls`）和CSV（`.csv`）文件格式。它会自动提取包含"感想"的列及其后续列，并按问题分类整理回答内容。

## 功能
1. 自动搜索当前目录下的Excel/CSV文件
2. 交互式选择文件和标识列
3. 智能识别问卷问题列（包含"感想"的列）
4. 生成格式化的文本结果文件

## 使用方法
1. 将问卷文件（Excel或CSV）放在程序所在目录
2. 使用python运行代码，或直接运行[release中的exe文件](https://github.com/ArcPen/rwz-helper/releases)
3. 按照界面提示选择文件和标识列
4. 程序会自动生成`[原文件名]_整理结果.txt`

## 依赖
- Python 3.6+
- pandas
- openpyxl (用于处理Excel文件)
- tkinter (图形化界面)

安装依赖：
```bash
pip install pandas openpyxl
```

## 输出格式
生成的文本文件按问题分组，每个回答以`- [标识符]: [回答内容]`的格式列出。

## 注意事项
- CSV文件会自动尝试多种编码（utf-8, gbk等）
- 程序会跳过空回答和标记为"(空)"的回答
- 标识列建议选择花名
