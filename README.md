# 批量文档翻译器

一个基于 Python 和 OpenAI API 的批量文档翻译工具，支持 Word 文档（.docx 和 .doc 格式）的批量翻译。

## 功能特点

- 📁 **批量处理**：支持多选 Word 文档进行批量翻译
- 🔄 **双向翻译**：支持中英互译，可自动检测语言方向
- 💾 **格式保持**：保持原文段落结构，输出 TXT 和 DOCX 格式
- 🎯 **灵活配置**：可自定义 API 端点、模型和输出目录
- 📊 **进度显示**：实时显示翻译进度和详细日志
- ⚡ **多线程处理**：后台线程处理，界面不卡顿

## 技术依赖

```bash
pip install python-docx docx2txt requests pywin32
使用说明
配置 API

输入 OpenAI API Key

设置 API 端点（默认：https://api.openai.com/v1）

选择模型（默认：gpt-3.5-turbo）

选择翻译方向

auto：自动检测语言方向

en->zh：英文翻译为中文

zh->en：中文翻译为英文

选择文件

点击"选择文件（多选）"按钮

支持 .docx 和 .doc 格式

开始翻译

点击"开始翻译"按钮

实时查看进度和日志

可随时停止翻译过程

输出文件
翻译完成后，在输出目录生成：

{原文件名}_translated.txt：纯文本格式

{原文件名}_translated.docx：Word 文档格式

注意事项
处理 .doc 文件需要 Windows 系统和 pywin32 库

确保 API Key 有效且有足够额度

网络连接稳定以保证翻译质量