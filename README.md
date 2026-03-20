# 微信公众号文章抓取与 Word 导出工具

这是一个基于 Python + Streamlit 的本地网页工具。输入一个或多个公开可访问的微信公众号文章链接后，程序会抓取文章内容并导出为 Word 文档。

## 功能特性

- 支持单篇或批量输入公众号文章链接
- 支持抓取标题、作者、公众号名称、发布时间和正文
- 支持导出为 .docx 文件
- 支持批量打包为 zip 下载
- 支持正文预览后再导出
- 支持设置字体、字号、正文行距
- 支持是否保留图片
- 支持 requests 失败时回退到 Playwright
- 对非法链接、网络超时、解析失败、图片下载失败、Word 导出失败做了异常处理

## 项目结构

```text
.
├─ app.py
├─ parser.py
├─ docx_exporter.py
├─ utils.py
├─ requirements.txt
└─ README.md
```

## 安装依赖

```bash
pip install -r requirements.txt
```

## 启动方式

```bash
streamlit run app.py
```

默认访问地址通常为：

```text
http://localhost:8501
```

## Playwright 浏览器安装

如果你希望在 requests 抓取失败时启用 Playwright 回退，请先安装浏览器：

```bash
python -m playwright install chromium
```

## 使用说明

1. 打开网页后，在文本框中粘贴一个或多个公众号文章链接，每行一个。
2. 可选填写导出目录、导出文件名、行距、字体和字号。
3. 选择是否保留图片，以及是否启用 Playwright 回退。
4. 点击“开始抓取并导出”。
5. 先确认页面中的正文预览是否正常。
6. 点击“确认并生成 Word 文件”后下载 docx，或在批量场景下下载 zip。

## 说明与限制

- 仅处理用户主动提供的、可公开访问的微信公众号文章链接
- 不支持绕过登录、绕过权限、绕过付费限制
- 工具仅用于个人学习、整理和导出公开内容
- 微信公众号页面结构可能变化，若部分页面抓取失败可尝试启用 Playwright 回退
